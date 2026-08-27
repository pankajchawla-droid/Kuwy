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

app.get("/", (_req, res) => res.send("Bulk image downloader is running."));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Listening on ${PORT}`));
