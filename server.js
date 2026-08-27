import express from "express";
import fetch from "node-fetch";
import archiver from "archiver";
import { v4 as uuid } from "uuid";
import { fromPath } from "pdf2pic";
import puppeteer from "puppeteer";
import fs from "fs";
import path from "path";
import os from "os";

const app = express();
app.use(express.json({ limit: "2mb" }));

/**
 * POST /bulk-download
 * body: {
 *   urls: string[],           // mix of direct .pdf links and/or HTML report page links
 *   cookie?: string,          // optional "name=value; name2=value2" cookie header for authenticated pages
 *   headers?: Record<string,string> // optional extra headers (e.g. Authorization)
 * }
 * returns: a .zip stream of all images found/rendered
 */
app.post("/bulk-download", async (req, res) => {
  const { urls, cookie, headers = {} } = req.body || {};
  if (!Array.isArray(urls) || urls.length === 0) {
    return res.status(400).json({ error: "Provide 'urls' as a non-empty array." });
  }

  const workDir = path.join(os.tmpdir(), `job-${uuid()}`);
  fs.mkdirSync(workDir, { recursive: true });

  res.setHeader("Content-Type", "application/zip");
  res.setHeader("Content-Disposition", `attachment; filename="images.zip"`);
  const archive = archiver("zip", { zlib: { level: 9 } });
  archive.pipe(res);

  const reqHeaders = { ...headers, ...(cookie ? { Cookie: cookie } : {}) };

  let browser;
  try {
    for (let i = 0; i < urls.length; i++) {
      const url = urls[i];
      const label = `item-${i + 1}`;
      try {
        if (url.toLowerCase().includes(".pdf")) {
          await handlePdf(url, reqHeaders, workDir, label, archive);
        } else {
          browser = browser || (await puppeteer.launch({ args: ["--no-sandbox"] }));
          await handleHtmlPage(url, reqHeaders, browser, label, archive);
        }
      } catch (err) {
        // Don't kill the whole batch if one URL fails — log it inside the zip instead.
        archive.append(`Failed to fetch ${url}: ${err.message}`, { name: `${label}-ERROR.txt` });
      }
    }
  } finally {
    if (browser) await browser.close();
    archive.finalize();
    fs.rm(workDir, { recursive: true, force: true }, () => {});
  }
});

async function handlePdf(url, reqHeaders, workDir, label, archive) {
  const resp = await fetch(url, { headers: reqHeaders });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  const pdfPath = path.join(workDir, `${label}.pdf`);
  fs.writeFileSync(pdfPath, Buffer.from(await resp.arrayBuffer()));

  // Also keep the raw PDF in the zip, in case that's wanted too.
  archive.append(fs.createReadStream(pdfPath), { name: `${label}/original.pdf` });

  // Render each page to a PNG.
  const converter = fromPath(pdfPath, {
    density: 200,
    saveFilename: label,
    savePath: workDir,
    format: "png",
    width: 1600,
    height: 2200,
  });
  const pageCount = await getPdfPageCount(pdfPath);
  for (let p = 1; p <= pageCount; p++) {
    const out = await converter(p, { responseType: "image" });
    archive.append(fs.createReadStream(out.path), { name: `${label}/page-${p}.png` });
  }
}

async function getPdfPageCount(pdfPath) {
  const { PDFDocument } = await import("pdf-lib");
  const bytes = fs.readFileSync(pdfPath);
  const doc = await PDFDocument.load(bytes, { ignoreEncryption: true });
  return doc.getPageCount();
}

async function handleHtmlPage(url, reqHeaders, browser, label, archive) {
  const page = await browser.newPage();
  if (reqHeaders.Cookie) {
    const domain = new URL(url).hostname;
    const cookies = reqHeaders.Cookie.split(";").map((c) => {
      const [name, ...rest] = c.trim().split("=");
      return { name, value: rest.join("="), domain, path: "/" };
    });
    await page.setCookie(...cookies);
  }
  if (reqHeaders["Authorization"]) {
    await page.setExtraHTTPHeaders({ Authorization: reqHeaders["Authorization"] });
  }
  await page.goto(url, { waitUntil: "networkidle2", timeout: 60000 });

  const imgUrls = await page.$$eval("img", (imgs) =>
    imgs.map((img) => img.src).filter(Boolean)
  );

  for (let i = 0; i < imgUrls.length; i++) {
    try {
      const resp = await fetch(imgUrls[i], { headers: reqHeaders });
      if (!resp.ok) continue;
      const buf = Buffer.from(await resp.arrayBuffer());
      const ext = path.extname(new URL(imgUrls[i]).pathname).split("?")[0] || ".jpg";
      archive.append(buf, { name: `${label}/img-${i + 1}${ext}` });
    } catch {
      // skip broken image
    }
  }
  await page.close();
}

app.get("/", (_req, res) => {
  res.send(`<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8" />
  <title>Bulk Image Downloader</title>
  <style>
    body { font-family: system-ui, sans-serif; max-width: 640px; margin: 40px auto; padding: 0 16px; color: #222; }
    h1 { font-size: 1.4rem; }
    textarea { width: 100%; height: 140px; font-family: monospace; font-size: 0.9rem; padding: 8px; box-sizing: border-box; }
    input[type=text] { width: 100%; padding: 8px; box-sizing: border-box; font-family: monospace; }
    label { display: block; margin-top: 16px; font-weight: 600; }
    button { margin-top: 20px; padding: 10px 20px; font-size: 1rem; cursor: pointer; }
    #status { margin-top: 16px; color: #555; }
    small { color: #666; }
  </style>
</head>
<body>
  <h1>Bulk Image Downloader</h1>
  <p><small>Paste one URL per line — direct .pdf report links or HTML report pages. If a page needs login, paste your session cookie below.</small></p>

  <label for="urls">Report URLs (one per line)</label>
  <textarea id="urls" placeholder="https://.../pdf_image/report.pdf
https://.../admin/inspectionReports?id=...&type=..."></textarea>

  <label for="cookie">Cookie (optional, for authenticated admin pages)</label>
  <input type="text" id="cookie" placeholder="session=abc123; other=xyz" />

  <button id="go">Download ZIP</button>
  <div id="status"></div>

  <script>
    document.getElementById("go").addEventListener("click", async () => {
      const urls = document.getElementById("urls").value
        .split("\\n").map(s => s.trim()).filter(Boolean);
      const cookie = document.getElementById("cookie").value.trim();
      const statusEl = document.getElementById("status");

      if (urls.length === 0) {
        statusEl.textContent = "Add at least one URL first.";
        return;
      }

      statusEl.textContent = "Working... this can take a while for PDFs / many pages.";
      document.getElementById("go").disabled = true;

      try {
        const resp = await fetch("/bulk-download", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ urls, cookie: cookie || undefined })
        });
        if (!resp.ok) throw new Error("Server returned " + resp.status);
        const blob = await resp.blob();
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = "images.zip";
        link.click();
        statusEl.textContent = "Done — check your downloads for images.zip.";
      } catch (err) {
        statusEl.textContent = "Error: " + err.message;
      } finally {
        document.getElementById("go").disabled = false;
      }
    });
  </script>
</body>
</html>`);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Listening on ${PORT}`));
