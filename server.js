import express from "express";
import fetch from "node-fetch";
import archiver from "archiver";
import { v4 as uuid } from "uuid";
import { fromPath } from "pdf2pic";
import puppeteer from "puppeteer";
import fs from "fs";
import path from "path";
import os from "os";
import multer from "multer";
import * as XLSX from "xlsx";

const app = express();
app.use(express.json({ limit: "2mb" }));
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });

// ---------- in-memory job store ----------
const jobs = new Map(); // jobId -> { status, current, total, log: [], zipPath, error, createdAt }
const JOB_TTL_MS = 30 * 60 * 1000; // clean up finished jobs after 30 min

setInterval(() => {
  const now = Date.now();
  for (const [id, job] of jobs) {
    if (now - job.createdAt > JOB_TTL_MS) {
      if (job.zipPath) fs.rm(job.zipPath, { force: true }, () => {});
      jobs.delete(id);
    }
  }
}, 5 * 60 * 1000);

function sanitizeFolderName(name, fallback) {
  const cleaned = String(name || "").trim().replace(/[\\/:*?"<>|]/g, "-");
  return cleaned || fallback;
}

function log(job, message) {
  job.log.push({ t: Date.now(), message });
}

// ---------- core per-row handlers ----------
async function handlePdf(url, reqHeaders, workDir, label, archive) {
  const resp = await fetch(url, { headers: reqHeaders });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  const pdfPath = path.join(workDir, `${label}-${uuid()}.pdf`);
  fs.writeFileSync(pdfPath, Buffer.from(await resp.arrayBuffer()));

  archive.append(fs.createReadStream(pdfPath), { name: `${label}/original.pdf` });

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

  const imgUrls = await page.$$eval("img", (imgs) => imgs.map((img) => img.src).filter(Boolean));

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

// ---------- job runner ----------
async function runJob(jobId, entries, cookie) {
  const job = jobs.get(jobId);
  const reqHeaders = cookie ? { Cookie: cookie } : {};
  const workDir = path.join(os.tmpdir(), `job-${jobId}`);
  fs.mkdirSync(workDir, { recursive: true });
  const zipPath = path.join(os.tmpdir(), `out-${jobId}.zip`);
  job.zipPath = zipPath;

  const output = fs.createWriteStream(zipPath);
  const archive = archiver("zip", { zlib: { level: 9 } });
  archive.pipe(output);

  const zipDone = new Promise((resolve, reject) => {
    output.on("close", resolve);
    archive.on("error", reject);
  });

  let browser;
  try {
    for (let i = 0; i < entries.length; i++) {
      const entry = entries[i];
      const label = sanitizeFolderName(entry.folder, `row-${i + 1}`);
      const url = String(entry.url).trim();
      job.current = i + 1;
      log(job, `Fetching "${label}"...`);
      try {
        if (url.toLowerCase().includes(".pdf")) {
          await handlePdf(url, reqHeaders, workDir, label, archive);
        } else {
          browser = browser || (await puppeteer.launch({ args: ["--no-sandbox"] }));
          await handleHtmlPage(url, reqHeaders, browser, label, archive);
        }
        log(job, `Done: "${label}"`);
      } catch (err) {
        log(job, `Failed "${label}": ${err.message}`);
        archive.append(`Failed to fetch ${url}: ${err.message}`, { name: `${label}-ERROR.txt` });
      }
    }
  } finally {
    if (browser) await browser.close();
    archive.finalize();
    await zipDone;
    fs.rm(workDir, { recursive: true, force: true }, () => {});
  }

  job.status = "done";
  log(job, "All tasks completed.");
}

function startJob(entries) {
  const jobId = uuid();
  jobs.set(jobId, {
    status: "processing",
    current: 0,
    total: entries.length,
    log: [],
    zipPath: null,
    createdAt: Date.now(),
  });
  return jobId;
}

// ---------- routes ----------
app.post("/jobs/from-excel", upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "Upload an .xlsx file under field name 'file'." });
  const cookie = (req.body && req.body.cookie) || "";

  let rows;
  try {
    const workbook = XLSX.read(req.file.buffer, { type: "buffer" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  } catch (err) {
    return res.status(400).json({ error: `Could not read Excel file: ${err.message}` });
  }

  if (rows.length && !/^https?:\/\//i.test(String(rows[0][1] || ""))) {
    rows = rows.slice(1);
  }

  const entries = rows
    .map((r) => ({ folder: r[0], url: r[1] }))
    .filter((r) => r.url && /^https?:\/\//i.test(String(r.url)));

  if (entries.length === 0) {
    return res.status(400).json({ error: "No valid rows found. Column A = folder name, Column B = URL." });
  }

  const jobId = startJob(entries);
  runJob(jobId, entries, cookie).catch((err) => {
    const job = jobs.get(jobId);
    if (job) {
      job.status = "error";
      job.error = err.message;
      log(job, `Job failed: ${err.message}`);
    }
  });
  res.json({ jobId });
});

app.post("/jobs/from-urls", async (req, res) => {
  const { urls, cookie } = req.body || {};
  if (!Array.isArray(urls) || urls.length === 0) {
    return res.status(400).json({ error: "Provide 'urls' as a non-empty array." });
  }
  const entries = urls.map((u, i) => ({ folder: `item-${i + 1}`, url: u }));
  const jobId = startJob(entries);
  runJob(jobId, entries, cookie || "").catch((err) => {
    const job = jobs.get(jobId);
    if (job) {
      job.status = "error";
      job.error = err.message;
      log(job, `Job failed: ${err.message}`);
    }
  });
  res.json({ jobId });
});

app.get("/jobs/:id", (req, res) => {
  const job = jobs.get(req.params.id);
  if (!job) return res.status(404).json({ error: "Job not found (it may have expired)." });
  res.json({
    status: job.status,
    current: job.current,
    total: job.total,
    log: job.log.slice(-100).map((l) => l.message),
    error: job.error || null,
  });
});

app.get("/jobs/:id/download", (req, res) => {
  const job = jobs.get(req.params.id);
  if (!job || job.status !== "done" || !job.zipPath || !fs.existsSync(job.zipPath)) {
    return res.status(404).json({ error: "Zip not ready or job not found." });
  }
  res.download(job.zipPath, "images.zip");
});

const HOMEPAGE_HTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>Inspection Ticket — Bulk Report Downloader</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Rajdhani:wght@600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
<style>
  :root{
    --asphalt:#14171A;
    --surface:#1D2126;
    --surface-2:#20262C;
    --ink:#ECEAE3;
    --muted:#8A9099;
    --orange:#FF6A13;
    --green:#3FA672;
    --rule:#2B3138;
  }
  *{box-sizing:border-box;}
  body{
    margin:0; min-height:100vh; background:var(--asphalt);
    background-image:
      radial-gradient(circle at 1px 1px, #262C32 1px, transparent 0);
    background-size: 22px 22px;
    color:var(--ink); font-family:'Inter',sans-serif;
    display:flex; align-items:flex-start; justify-content:center; padding:48px 16px;
  }
  .ticket{
    width:100%; max-width:620px; background:var(--surface);
    border:1px solid var(--rule); border-radius:6px; position:relative; overflow:hidden;
  }
  .ticket::before{ /* perforation strip */
    content:""; display:block; height:16px;
    background-image: radial-gradient(circle, var(--asphalt) 5px, transparent 5.5px);
    background-size: 20px 20px; background-position: 6px -6px;
    background-color: var(--orange);
  }
  .head{ padding:22px 28px 18px; border-bottom:1px dashed var(--rule); }
  .eyebrow{ font-family:'JetBrains Mono',monospace; font-size:12px; letter-spacing:.14em; color:var(--orange); text-transform:uppercase; }
  h1{ font-family:'Rajdhani',sans-serif; font-weight:700; font-size:28px; margin:4px 0 6px; letter-spacing:.02em; }
  .sub{ color:var(--muted); font-size:14px; line-height:1.5; margin:0; }

  .body{ padding:24px 28px 28px; }
  .tabs{ display:flex; gap:4px; margin-bottom:18px; }
  .tab{
    font-family:'Rajdhani',sans-serif; font-weight:700; font-size:15px; letter-spacing:.03em;
    background:none; border:1px solid var(--rule); color:var(--muted); padding:8px 16px;
    cursor:pointer; border-radius:4px 4px 0 0; border-bottom:none;
  }
  .tab.active{ color:var(--ink); background:var(--surface-2); border-color:var(--rule); position:relative; top:1px; }

  .panel{ background:var(--surface-2); border:1px solid var(--rule); border-radius:0 4px 4px 4px; padding:20px; }

  label{ display:block; font-family:'JetBrains Mono',monospace; font-size:12px; color:var(--muted); text-transform:uppercase; letter-spacing:.08em; margin-bottom:8px; }
  textarea, input[type=text]{
    width:100%; background:var(--asphalt); border:1px solid var(--rule); border-radius:4px;
    color:var(--ink); font-family:'JetBrains Mono',monospace; font-size:13px; padding:10px;
  }
  textarea{ height:120px; resize:vertical; }
  .field{ margin-bottom:16px; }
  .hint{ font-size:12px; color:var(--muted); margin:6px 0 0; line-height:1.5; }

  .dropzone{
    border:2px dashed var(--rule); border-radius:6px; padding:28px 16px; text-align:center;
    cursor:pointer; transition:border-color .15s, background .15s;
  }
  .dropzone:hover, .dropzone.drag{ border-color:var(--orange); background:rgba(255,106,19,0.06); }
  .dropzone .icon{ font-family:'Rajdhani',sans-serif; font-weight:700; font-size:32px; color:var(--orange); }
  .dropzone .primary{ font-weight:600; margin:6px 0 2px; }
  .dropzone .secondary{ color:var(--muted); font-size:13px; }
  .filename{ margin-top:10px; font-family:'JetBrains Mono',monospace; font-size:13px; color:var(--green); display:none; }

  button.cta{
    width:100%; margin-top:16px; padding:13px; border:none; border-radius:4px;
    background:var(--orange); color:#1a1400; font-family:'Rajdhani',sans-serif;
    font-weight:700; font-size:16px; letter-spacing:.03em; cursor:pointer;
  }
  button.cta:disabled{ opacity:.5; cursor:not-allowed; }
  button.cta.secondary{ background:transparent; border:1px solid var(--rule); color:var(--ink); }

  .gauge-wrap{ display:flex; align-items:center; gap:18px; margin-top:20px; }
  .gauge{ width:88px; height:52px; flex-shrink:0; }
  .needle{ transform-origin:44px 44px; transition:transform .4s ease; }
  .progress-text{ font-family:'JetBrains Mono',monospace; }
  .progress-pct{ font-size:26px; font-weight:600; color:var(--orange); }
  .progress-label{ font-size:12px; color:var(--muted); margin-top:2px; }

  .log{
    margin-top:18px; max-height:180px; overflow-y:auto; font-family:'JetBrains Mono',monospace;
    font-size:12.5px; border-top:1px dashed var(--rule); padding-top:12px; display:none;
  }
  .log-line{ padding:3px 0; color:var(--muted); }
  .log-line.err{ color:#E6714B; }
  .log-line.ok{ color:var(--green); }

  .stamp{
    display:none; margin-top:18px; border:3px solid var(--green); color:var(--green);
    font-family:'Rajdhani',sans-serif; font-weight:700; font-size:15px; letter-spacing:.12em;
    text-transform:uppercase; text-align:center; padding:10px; border-radius:6px;
    transform:rotate(-2deg);
  }

  .foot{ padding:14px 28px; border-top:1px dashed var(--rule); display:flex; justify-content:space-between; align-items:center; }
  .made-by{
    font-family:'JetBrains Mono',monospace; font-size:11px; color:var(--muted);
    border:1px solid var(--rule); padding:4px 10px; border-radius:20px; letter-spacing:.05em;
  }
  .made-by b{ color:var(--orange); }
  @media (max-width:480px){ .ticket{border-radius:0;} body{padding:0;} }
</style>
</head>
<body>
  <div class="ticket">
    <div class="head">
      <div class="eyebrow">Report Retrieval · Bulk Job</div>
      <h1>Inspection Ticket</h1>
      <p class="sub">Pull page images out of vehicle report PDFs and admin report pages, in bulk. One ticket in, one zip out.</p>
    </div>

    <div class="body">
      <div class="tabs">
        <button class="tab active" id="tabExcel">Excel Sheet</button>
        <button class="tab" id="tabPaste">Paste URLs</button>
      </div>

      <div class="panel" id="panelExcel">
        <div class="field">
          <label>Report sheet</label>
          <div class="dropzone" id="dropzone">
            <div class="icon">⛽</div>
            <div class="primary">Drop your .xlsx here</div>
            <div class="secondary">or click to browse — Column A: folder name · Column B: URL</div>
            <div class="filename" id="filename"></div>
          </div>
          <input type="file" id="xlsxFile" accept=".xlsx,.xls" style="display:none" />
          <p class="hint">First row is auto-detected as a header and skipped if column B isn't a link.</p>
        </div>
        <div class="field">
          <label>Session cookie (optional)</label>
          <input type="text" id="cookieExcel" placeholder="session=abc123; other=xyz" />
          <p class="hint">Needed for admin/report pages that require you to be logged in.</p>
        </div>
        <button class="cta" id="goExcel">Start Processing</button>
      </div>

      <div class="panel" id="panelPaste" style="display:none;">
        <div class="field">
          <label>Report URLs (one per line)</label>
          <textarea id="urls" placeholder="https://.../pdf_image/report.pdf
https://.../admin/inspectionReports?id=...&type=..."></textarea>
        </div>
        <div class="field">
          <label>Session cookie (optional)</label>
          <input type="text" id="cookiePaste" placeholder="session=abc123; other=xyz" />
        </div>
        <button class="cta" id="goPaste">Start Processing</button>
      </div>

      <div class="gauge-wrap" id="gaugeWrap" style="display:none;">
        <svg class="gauge" viewBox="0 0 88 52">
          <path d="M4,48 A40,40 0 0 1 84,48" fill="none" stroke="#2B3138" stroke-width="6" stroke-linecap="round"/>
          <path id="gaugeFill" d="M4,48 A40,40 0 0 1 84,48" fill="none" stroke="#FF6A13" stroke-width="6" stroke-linecap="round" stroke-dasharray="0 126"/>
          <line class="needle" id="needle" x1="44" y1="44" x2="44" y2="14" stroke="#ECEAE3" stroke-width="3" stroke-linecap="round"/>
          <circle cx="44" cy="44" r="4" fill="#ECEAE3"/>
        </svg>
        <div class="progress-text">
          <div class="progress-pct" id="pctText">0%</div>
          <div class="progress-label" id="progressLabel">Waiting to start…</div>
        </div>
      </div>

      <div class="log" id="log"></div>

      <div class="stamp" id="stamp">✓ Batch Complete</div>

      <button class="cta secondary" id="downloadBtn" style="display:none;">Download All Files (ZIP)</button>
      <button class="cta secondary" id="resetBtn" style="display:none;">Process Another Sheet</button>
    </div>

    <div class="foot">
      <span class="hint" style="margin:0;">Runs entirely on your own Render service.</span>
      <span class="made-by">Made by <b>PC</b></span>
    </div>
  </div>

<script>
  const tabExcel = document.getElementById('tabExcel');
  const tabPaste = document.getElementById('tabPaste');
  const panelExcel = document.getElementById('panelExcel');
  const panelPaste = document.getElementById('panelPaste');
  tabExcel.onclick = () => { tabExcel.classList.add('active'); tabPaste.classList.remove('active'); panelExcel.style.display='block'; panelPaste.style.display='none'; };
  tabPaste.onclick = () => { tabPaste.classList.add('active'); tabExcel.classList.remove('active'); panelPaste.style.display='block'; panelExcel.style.display='none'; };

  const dropzone = document.getElementById('dropzone');
  const xlsxFile = document.getElementById('xlsxFile');
  const filenameEl = document.getElementById('filename');
  dropzone.onclick = () => xlsxFile.click();
  ['dragover','dragenter'].forEach(evt => dropzone.addEventListener(evt, e => { e.preventDefault(); dropzone.classList.add('drag'); }));
  ['dragleave','drop'].forEach(evt => dropzone.addEventListener(evt, e => { e.preventDefault(); dropzone.classList.remove('drag'); }));
  dropzone.addEventListener('drop', e => {
    if (e.dataTransfer.files.length) { xlsxFile.files = e.dataTransfer.files; showFilename(); }
  });
  xlsxFile.addEventListener('change', showFilename);
  function showFilename(){
    if (xlsxFile.files.length){ filenameEl.textContent = '📄 ' + xlsxFile.files[0].name; filenameEl.style.display='block'; }
  }

  const gaugeWrap = document.getElementById('gaugeWrap');
  const gaugeFill = document.getElementById('gaugeFill');
  const needle = document.getElementById('needle');
  const pctText = document.getElementById('pctText');
  const progressLabel = document.getElementById('progressLabel');
  const logEl = document.getElementById('log');
  const stampEl = document.getElementById('stamp');
  const downloadBtn = document.getElementById('downloadBtn');
  const resetBtn = document.getElementById('resetBtn');

  const CIRC = 126; // approx arc length for dasharray

  function setProgress(current, total){
    const pct = total ? Math.round((current/total)*100) : 0;
    pctText.textContent = pct + '%';
    progressLabel.textContent = 'Processing ' + current + ' of ' + total;
    gaugeFill.setAttribute('stroke-dasharray', (CIRC*pct/100) + ' ' + CIRC);
    const angle = -90 + (pct/100)*180; // -90deg to +90deg sweep
    needle.setAttribute('transform', 'rotate(' + angle + ' 44 44)');
  }

  function addLog(message){
    const line = document.createElement('div');
    line.className = 'log-line' + (/failed/i.test(message) ? ' err' : /done|completed/i.test(message) ? ' ok' : '');
    line.textContent = message;
    logEl.appendChild(line);
    logEl.scrollTop = logEl.scrollHeight;
  }

  function resetUI(){
    gaugeWrap.style.display='none'; logEl.style.display='none'; logEl.innerHTML='';
    stampEl.style.display='none'; downloadBtn.style.display='none'; resetBtn.style.display='none';
    document.getElementById('goExcel').disabled = false;
    document.getElementById('goPaste').disabled = false;
    setProgress(0,0);
  }

  async function pollJob(jobId){
    gaugeWrap.style.display='flex'; logEl.style.display='block';
    let seenLogs = 0;
    while (true){
      const resp = await fetch('/jobs/' + jobId);
      if (!resp.ok){ addLog('Lost track of job — it may have expired.'); break; }
      const data = await resp.json();
      setProgress(data.current, data.total);
      for (let i = seenLogs; i < data.log.length; i++) addLog(data.log[i]);
      seenLogs = data.log.length;

      if (data.status === 'done'){
        stampEl.style.display='block';
        downloadBtn.style.display='block';
        downloadBtn.onclick = () => { window.location.href = '/jobs/' + jobId + '/download'; };
        resetBtn.style.display='block';
        break;
      }
      if (data.status === 'error'){
        addLog('Job failed: ' + (data.error || 'unknown error'));
        resetBtn.style.display='block';
        break;
      }
      await new Promise(r => setTimeout(r, 900));
    }
    document.getElementById('goExcel').disabled = false;
    document.getElementById('goPaste').disabled = false;
  }

  document.getElementById('goExcel').onclick = async () => {
    if (!xlsxFile.files.length){ addLog('Choose an .xlsx file first.'); logEl.style.display='block'; return; }
    resetUI();
    document.getElementById('goExcel').disabled = true;
    const formData = new FormData();
    formData.append('file', xlsxFile.files[0]);
    const cookie = document.getElementById('cookieExcel').value.trim();
    if (cookie) formData.append('cookie', cookie);
    try {
      const resp = await fetch('/jobs/from-excel', { method:'POST', body: formData });
      const data = await resp.json();
      if (!resp.ok) throw new Error(data.error || 'Failed to start job');
      pollJob(data.jobId);
    } catch (err) {
      addLog('Error: ' + err.message); logEl.style.display='block';
      document.getElementById('goExcel').disabled = false;
    }
  };

  document.getElementById('goPaste').onclick = async () => {
    const urls = document.getElementById('urls').value.split('\\n').map(s=>s.trim()).filter(Boolean);
    if (!urls.length){ addLog('Add at least one URL first.'); logEl.style.display='block'; return; }
    resetUI();
    document.getElementById('goPaste').disabled = true;
    const cookie = document.getElementById('cookiePaste').value.trim();
    try {
      const resp = await fetch('/jobs/from-urls', {
        method:'POST', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({ urls, cookie: cookie || undefined })
      });
      const data = await resp.json();
      if (!resp.ok) throw new Error(data.error || 'Failed to start job');
      pollJob(data.jobId);
    } catch (err) {
      addLog('Error: ' + err.message); logEl.style.display='block';
      document.getElementById('goPaste').disabled = false;
    }
  };

  resetBtn.onclick = resetUI;
</script>
</body>
</html>`;

app.get("/", (_req, res) => {
  res.send(HOMEPAGE_HTML);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Listening on ${PORT}`));

