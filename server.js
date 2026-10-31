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
import crypto from "crypto";
import { execFile } from "child_process";
import { promisify } from "util";

const execFileAsync = promisify(execFile);

const SKIP_EXTENSIONS = new Set([".svg", ".webp"]);

function hashBuffer(buf) {
  return crypto.createHash("sha256").update(buf).digest("hex");
}

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
async function handlePdf(url, reqHeaders, workDir, label, archive, seenHashes, job) {
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
  let kept = 0;
  for (let p = 1; p <= pageCount; p++) {
    const out = await converter(p, { responseType: "image" });
    const buf = fs.readFileSync(out.path);
    const hash = hashBuffer(buf);
    if (seenHashes.has(hash)) {
      log(job, `Skipped duplicate page image in "${label}" (page ${p}).`);
      continue;
    }
    seenHashes.add(hash);
    kept++;
    archive.append(buf, { name: `${label}/page-${p}.png` });
  }
  return kept;
}

async function getPdfPageCount(pdfPath) {
  const { PDFDocument } = await import("pdf-lib");
  const bytes = fs.readFileSync(pdfPath);
  const doc = await PDFDocument.load(bytes, { ignoreEncryption: true });
  return doc.getPageCount();
}

async function handleHtmlPage(url, reqHeaders, browser, label, archive, seenHashes, job) {
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

  let kept = 0;
  let skippedType = 0;
  let skippedDupe = 0;
  for (let i = 0; i < imgUrls.length; i++) {
    try {
      const rawExt = path.extname(new URL(imgUrls[i]).pathname).split("?")[0].toLowerCase();
      if (SKIP_EXTENSIONS.has(rawExt)) {
        skippedType++;
        continue;
      }
      const resp = await fetch(imgUrls[i], { headers: reqHeaders });
      if (!resp.ok) continue;
      const buf = Buffer.from(await resp.arrayBuffer());
      const hash = hashBuffer(buf);
      if (seenHashes.has(hash)) {
        skippedDupe++;
        continue;
      }
      seenHashes.add(hash);
      const ext = rawExt || ".jpg";
      archive.append(buf, { name: `${label}/img-${i + 1}${ext}` });
      kept++;
    } catch {
      // skip broken image
    }
  }
  if (skippedType) log(job, `Skipped ${skippedType} .svg/.webp image(s) in "${label}".`);
  if (skippedDupe) log(job, `Skipped ${skippedDupe} duplicate image(s) in "${label}".`);
  await page.close();
  return kept;
}

// ---------- Nextcloud / ownCloud public share links ----------
function isNextcloudShare(url) {
  // Matches both classic "/index.php/s/<token>" and pretty-URL "/s/<token>" share links.
  return /\/(?:index\.php\/)?s\/[A-Za-z0-9_-]+(?:[/?#]|$)/i.test(url);
}

// ---------- generic "direct file endpoint" links (any vendor whose URL is clearly ----------
// ---------- a raw-file download rather than an HTML page — e.g. /export/DownloadFiles/123, ----------
// ---------- or a query-string action like ?action=download_case_images ----------
function isDirectDownloadLink(url) {
  try {
    const u = new URL(url);
    const p = u.pathname.toLowerCase();
    const action = (u.searchParams.get("action") || "").toLowerCase();
    return p.includes("download") || action.includes("download");
  } catch {
    return false;
  }
}

const BROWSER_LIKE_HEADERS = {
  "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
  "Accept": "*/*",
};

// Given an already-fetched buffer + its content-type, save it into the archive the right way
// (folder-zip / single PDF / single image). Shared by the Nextcloud handler and the generic
// direct-download handler so both link types get identical, tested handling.
// Minimum pixel size (on the longer side) for an embedded PDF image to be treated as real
// content (a photo) rather than template chrome (icons, logos, flat-color bars, QR codes).
// Tune this up/down if real photos get filtered out or icons still slip through.
const MIN_EMBEDDED_IMAGE_DIMENSION = 180;

// Reads width/height straight from a PNG or JPEG file's own header bytes, without decoding
// the image — fast, and avoids adding an image-processing dependency. Returns null if the
// format isn't recognized or the header can't be parsed (caller should then keep the image
// rather than risk dropping real content on a parsing miss).
function readImageDimensions(buf) {
  try {
    // PNG: 8-byte signature, then an IHDR chunk whose data starts with width(4) + height(4).
    if (buf.length > 24 && buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4e && buf[3] === 0x47) {
      return { width: buf.readUInt32BE(16), height: buf.readUInt32BE(20) };
    }
    // JPEG: scan markers for an SOFn segment, which holds height then width.
    if (buf.length > 4 && buf[0] === 0xff && buf[1] === 0xd8) {
      let pos = 2;
      while (pos + 9 < buf.length) {
        if (buf[pos] !== 0xff) { pos++; continue; }
        const marker = buf[pos + 1];
        if (marker === 0xd8 || marker === 0xd9) { pos += 2; continue; }
        if (marker >= 0xd0 && marker <= 0xd7) { pos += 2; continue; }
        const segLen = buf.readUInt16BE(pos + 2);
        const isSOF = marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
        if (isSOF) {
          return { height: buf.readUInt16BE(pos + 5), width: buf.readUInt16BE(pos + 7) };
        }
        pos += 2 + segLen;
      }
    }
  } catch {
    // fall through to null below
  }
  return null;
}

// Minimum number of distinct colors (after quantizing) an image needs to be treated as a
// real photo rather than template chrome (logos, flat-color bars, icons). Real photos have
// rich, continuous-tone detail even at modest size; flat graphics use only a handful of
// colors no matter how large they're rendered. Tune if real low-detail photos get dropped,
// or flat graphics still slip through.
const MIN_DISTINCT_COLORS = 40;

// Decodes a PNG/JPEG and returns how many visually-distinct colors it contains (sampled and
// quantized for speed), or null if the format isn't supported / decoding fails — callers
// should treat null as "keep the image" rather than risk dropping real content.
async function countDistinctColors(buf, ext) {
  try {
    let width, height, data; // data: flat RGBA (or RGB) byte array
    if (ext === ".png") {
      const { PNG } = await import("pngjs");
      const png = PNG.sync.read(buf);
      width = png.width; height = png.height; data = png.data; // RGBA
    } else if (ext === ".jpg" || ext === ".jpeg") {
      const jpeg = (await import("jpeg-js")).default;
      const decoded = jpeg.decode(buf, { useTArray: true });
      width = decoded.width; height = decoded.height; data = decoded.data; // RGBA
    } else {
      return null; // unsupported format for decoding (e.g. ppm/jp2 from pdfimages) — keep it
    }

    const totalPixels = width * height;
    if (totalPixels === 0) return 0;
    const channels = data.length / totalPixels; // 4 for RGBA, 3 for RGB
    const maxSamples = 4000;
    const step = Math.max(1, Math.floor(totalPixels / maxSamples));
    const seen = new Set();
    for (let i = 0; i < totalPixels; i += step) {
      const off = i * channels;
      // Quantize each channel to collapse near-identical shades (compression noise, anti-aliasing).
      const r = data[off] >> 4, g = data[off + 1] >> 4, b = data[off + 2] >> 4;
      seen.add((r << 8) | (g << 4) | b);
    }
    return seen.size;
  } catch {
    return null;
  }
}

// ---------- optional: arrange vehicle photos front, left, right, back, other ----------
// Needs an Anthropic API key in the ANTHROPIC_API_KEY environment variable. Each kept photo is
// shrunk and sent to a Claude model, which names the view. Without the key nothing here runs.
const VIEW_ORDER = ["front", "left", "right", "back", "other"];
const ANTHROPIC_API_URL = process.env.ANTHROPIC_API_URL || "https://api.anthropic.com/v1/messages";
const VIEW_RULES =
  "Say which view of the vehicle's exterior each photo shows: front, left, right, back, or other. " +
  "Use other for anything that is not an exterior view of the whole vehicle (interior, dashboard, engine bay, " +
  "tyre or wheel close-up, odometer, documents, labels, people). " +
  "For a diagonal view, answer front or back when that end of the vehicle faces the camera, otherwise left or right. " +
  "Left and right mean the vehicle's own left and right side.";

function batchPrompt(n) {
  return (
    `These ${n} photo${n === 1 ? " comes" : "s come"} from a vehicle inspection report, labelled Image 1` +
    `${n > 1 ? " to Image " + n : ""}. ${VIEW_RULES} Answer with exactly ${n} line${n === 1 ? "" : "s"}, ` +
    `one per image, written like "1: front" or "2: other", and nothing else.`
  );
}

async function prepareForClassification(buf, ext, workDir) {
  const inPath = path.join(workDir, `cls-${uuid()}${ext}`);
  const outPath = path.join(workDir, `cls-${uuid()}.jpg`);
  try {
    fs.writeFileSync(inPath, buf);
    await execFileAsync("gm", ["convert", inPath, "-resize", "512x512>", "-quality", "70", outPath]);
    return { data: fs.readFileSync(outPath), mediaType: "image/jpeg" };
  } catch {
    if (buf.length < 4 * 1024 * 1024 && [".jpg", ".jpeg", ".png"].includes(ext)) {
      return { data: buf, mediaType: ext === ".png" ? "image/png" : "image/jpeg" };
    }
    return null;
  }
}

function orderingModel() {
  return process.env.ANTHROPIC_MODEL || "claude-haiku-5-5";
}

// Which AI service arranges the photos: Claude if ANTHROPIC_API_KEY is set, otherwise Google
// Gemini (which has a free tier) if GEMINI_API_KEY is set. null means ordering is off.
function envAiConfig() {
  if (process.env.ANTHROPIC_API_KEY) return { provider: "claude", key: process.env.ANTHROPIC_API_KEY };
  if (process.env.GEMINI_API_KEY) return { provider: "gemini", key: process.env.GEMINI_API_KEY };
  return null;
}

function orderingProvider() {
  const c = envAiConfig();
  return c ? c.provider : null;
}

// A key pasted into the page: sk-ant-... is a Claude key, anything else is treated as a Gemini key.
function aiConfigFromKey(rawKey) {
  const key = typeof rawKey === "string" ? rawKey.trim() : "";
  if (key.length < 20 || key.length > 300 || /\s/.test(key)) return null;
  return { provider: key.startsWith("sk-ant-") ? "claude" : "gemini", key };
}

// The AI settings a job uses: the key pasted for that job, else the server-wide keys.
function aiConfigFor(job) {
  return (job && job.ai) || envAiConfig();
}

// Google retires Gemini model names over time, so the model is looked up from Google's own model
// list (newest plain "-flash" model that can generate content) instead of being hard-coded. Set
// GEMINI_MODEL to force a specific one. If the list can't be read, the fallback below is used.
const GEMINI_FALLBACK_MODEL = "gemini-3.8-flash";
const geminiModelCache = new Map(); // api key -> model name

function geminiBase() {
  return process.env.GEMINI_API_URL || "https://generativelanguage.googleapis.com/v1beta";
}

function pickGeminiModel(models) {
  const usable = (models || []).filter(
    (m) => !m.supportedGenerationMethods || m.supportedGenerationMethods.includes("generateContent")
  );
  for (const pattern of [/^gemini-(\d+(?:\.\d+)?)-flash$/, /^gemini-(\d+(?:\.\d+)?)-flash-lite$/]) {
    const found = usable
      .map((m) => {
        const name = String(m.name || "").replace(/^models\//, "");
        const mm = name.match(pattern);
        return mm ? { name, version: parseFloat(mm[1]) } : null;
      })
      .filter(Boolean)
      .sort((a, b) => b.version - a.version);
    if (found.length) return found[0].name;
  }
  return null;
}

async function resolveGeminiModel(key, forceRefresh = false) {
  if (process.env.GEMINI_MODEL) return process.env.GEMINI_MODEL;
  if (!forceRefresh && geminiModelCache.has(key)) return geminiModelCache.get(key);
  try {
    const resp = await fetch(`${geminiBase()}/models?pageSize=200`, { headers: { "x-goog-api-key": key }, signal: AbortSignal.timeout(15000) });
    if (resp.ok) {
      const best = pickGeminiModel((await resp.json()).models);
      if (best) {
        geminiModelCache.set(key, best);
        return best;
      }
    }
  } catch {
    // fall through to the cached or fallback model
  }
  return geminiModelCache.get(key) || GEMINI_FALLBACK_MODEL;
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const AI_TIMEOUT_MS = Number(process.env.AI_TIMEOUT_MS) || 60000; // one request to the AI service
const ORDERING_BUDGET_MS = Number(process.env.ORDERING_BUDGET_MS) || 150000; // arranging one report's photos
const BATCH_SIZE = 6; // photos per AI request, so a free key's few-requests-a-minute limit is enough

// Posts JSON to an AI service. Every attempt has a time limit. Rate-limit and server errors are retried
// (waiting as the service asks, at most about 90 seconds in total). A daily-limit error is not retried,
// because waiting cannot fix it. Failures become readable messages for the activity log.
async function postJsonWithRetry(name, url, headers, body, attempts, note) {
  let lastErr;
  let waitedMs = 0;
  for (let attempt = 0; attempt < attempts; attempt++) {
    let resp;
    try {
      resp = await fetch(url, {
        method: "POST",
        headers: { "content-type": "application/json", ...headers },
        body: JSON.stringify(body),
        signal: AbortSignal.timeout(AI_TIMEOUT_MS),
      });
    } catch (err) {
      const timedOut = err && (err.name === "AbortError" || err.name === "TimeoutError");
      lastErr = new Error(`could not reach ${name} (${timedOut ? "no answer in time" : err.message})`);
      if (attempt === attempts - 1) throw lastErr;
      await sleep(Math.min(20000, 1000 * 2 ** attempt));
      continue;
    }
    if (resp.ok) return resp.json();
    let detail = "";
    let rawDetails = "";
    let bodyRetrySeconds = NaN;
    try {
      const j = await resp.json();
      detail = (j.error && j.error.message) || "";
      rawDetails = JSON.stringify((j.error && j.error.details) || []);
      const info = ((j.error && j.error.details) || []).find((d) => d && d.retryDelay);
      if (info) bodyRetrySeconds = parseFloat(info.retryDelay);
    } catch {
      // body was not JSON
    }
    if (resp.status === 429 && (/PerDay|per day|daily/i.test(rawDetails) || /per day|daily/i.test(detail))) {
      throw new Error(`${name} daily free limit reached (${detail.slice(0, 100)}). Try again tomorrow or use a different key.`);
    }
    lastErr = new Error(`${name} HTTP ${resp.status}${detail ? ": " + detail.slice(0, 160) : ""}`);
    const retryable = resp.status === 429 || resp.status === 529 || resp.status >= 500;
    if (!retryable || attempt === attempts - 1) throw lastErr;
    const headerSeconds = Number(resp.headers.get("retry-after"));
    const seconds = headerSeconds > 0 ? headerSeconds : bodyRetrySeconds > 0 ? bodyRetrySeconds : 2 ** attempt;
    const waitMs = Math.min(65000, seconds * 1000);
    if (waitedMs + waitMs > 90000) throw lastErr;
    waitedMs += waitMs;
    if (note) note(`${name} is busy or rate-limiting; waiting ${Math.ceil(waitMs / 1000)}s before retrying...`);
    await sleep(waitMs);
  }
  throw lastErr;
}

async function callAnthropic(body, attempts = 4, key = process.env.ANTHROPIC_API_KEY, note) {
  return postJsonWithRetry(
    "Claude API",
    ANTHROPIC_API_URL,
    { "x-api-key": key, "anthropic-version": "2023-06-01" },
    body,
    attempts,
    note
  );
}

// parts: Gemini "parts" array (text and/or inline_data). Returns the reply text.
async function callGemini(parts, attempts = 4, key = process.env.GEMINI_API_KEY, note) {
  async function generate(model) {
    const data = await postJsonWithRetry(
      "Gemini API",
      `${geminiBase()}/models/${model}:generateContent`,
      { "x-goog-api-key": key },
      { contents: [{ parts }], generationConfig: { temperature: 0, maxOutputTokens: 256 } },
      attempts,
      note
    );
    const cand = (data.candidates || [])[0];
    return ((cand && cand.content && cand.content.parts) || []).map((p) => p.text || "").join(" ");
  }
  const model = await resolveGeminiModel(key);
  try {
    return await generate(model);
  } catch (err) {
    // A 404 usually means Google retired this model name: look the model up again and retry once.
    if (/HTTP 404/.test(err.message) && !process.env.GEMINI_MODEL) {
      const fresh = await resolveGeminiModel(key, true);
      if (fresh !== model) return generate(fresh);
    }
    throw err;
  }
}

const VIEW_WORDS = { front: "front", left: "left", right: "right", back: "back", rear: "back", other: "other" };

// Asks the AI about several photos in ONE request. items: [{ buf, ext }]. Returns one view (or null when
// the answer did not cover that photo) per item.
async function classifyBatch(items, workDir, ai, note) {
  const result = new Array(items.length).fill(null);
  const sent = [];
  const images = [];
  for (let i = 0; i < items.length; i++) {
    const img = await prepareForClassification(items[i].buf, items[i].ext, workDir);
    if (img) {
      sent.push(i);
      images.push(img);
    } else {
      result[i] = "other";
    }
  }
  if (images.length === 0) return result;
  const prompt = batchPrompt(images.length);
  let text;
  if (ai.provider === "gemini") {
    const parts = [];
    images.forEach((img, k) => {
      parts.push({ text: `Image ${k + 1}:` });
      parts.push({ inline_data: { mime_type: img.mediaType, data: img.data.toString("base64") } });
    });
    parts.push({ text: prompt });
    text = await callGemini(parts, 4, ai.key, note);
  } else {
    const content = [];
    images.forEach((img, k) => {
      content.push({ type: "text", text: `Image ${k + 1}:` });
      content.push({ type: "image", source: { type: "base64", media_type: img.mediaType, data: img.data.toString("base64") } });
    });
    content.push({ type: "text", text: prompt });
    const data = await callAnthropic(
      { model: orderingModel(), max_tokens: 20 + images.length * 12, messages: [{ role: "user", content }] },
      4,
      ai.key,
      note
    );
    text = ((data.content || []).find((c) => c.type === "text") || {}).text || "";
  }
  const lower = text.toLowerCase();
  for (const m of lower.matchAll(/(\d+)\s*[:.)-]\s*\**\s*(front|left|right|back|rear|other)/g)) {
    const n = Number(m[1]);
    if (n >= 1 && n <= images.length && !result[sent[n - 1]]) result[sent[n - 1]] = VIEW_WORDS[m[2]];
  }
  if (images.length === 1 && !result[sent[0]]) {
    const single = lower.match(/front|left|right|back|rear|other/);
    if (single) result[sent[0]] = VIEW_WORDS[single[0]];
  }
  return result;
}

function withDeadline(promise, ms) {
  let timer;
  const timeout = new Promise((_, reject) => {
    timer = setTimeout(() => reject(new Error("time limit reached")), ms);
  });
  return Promise.race([promise, timeout]).finally(() => clearTimeout(timer));
}

// Sends the photos at `indices` to the AI in batches, with a time limit for the whole report. Photos the
// AI could not place stay null (the caller puts them under "other"). Two failed batches in a row switch
// the AI off for the rest of the job, so a broken or used-up key can never keep a run waiting.
async function classifyInBatches(items, indices, views, ai, workDir, job, label) {
  const batches = [];
  for (let i = 0; i < indices.length; i += BATCH_SIZE) batches.push(indices.slice(i, i + BATCH_SIZE));
  const who = ai.provider === "gemini" ? "Google Gemini" : "Claude";
  log(job, `Arranging ${indices.length} photo(s) in "${label}" with ${who} (${batches.length} request(s)); this can take a minute or two on a free key.`);
  const deadline = Date.now() + ORDERING_BUDGET_MS;
  const note = (msg) => {
    job.waitNotes = (job.waitNotes || 0) + 1;
    if (job.waitNotes <= 6) log(job, msg);
  };
  let next = 0;
  let failedPhotos = 0;
  let firstError = null;
  let timedOut = false;
  let successes = 0;
  async function worker() {
    while (true) {
      const b = next++;
      if (b >= batches.length) return;
      if (job.aiDisabled) continue;
      const remaining = deadline - Date.now();
      if (remaining <= 0) {
        timedOut = true;
        continue;
      }
      try {
        const labels = await withDeadline(classifyBatch(batches[b].map((i) => items[i]), workDir, ai, note), remaining);
        labels.forEach((v, k) => {
          if (v) views[batches[b][k]] = v;
        });
        job.aiFailures = 0;
        successes++;
      } catch (err) {
        failedPhotos += batches[b].length;
        if (err.message === "time limit reached") {
          timedOut = true;
        } else {
          firstError = firstError || err.message;
          job.aiFailures = (job.aiFailures || 0) + 1;
          if (job.aiFailures >= 2) job.aiDisabled = true;
        }
      }
    }
  }
  await Promise.all(Array.from({ length: Math.min(ai.provider === "gemini" ? 1 : 2, batches.length) }, worker));
  if (timedOut && successes === 0) {
    // the service never answered for this whole report; two reports like that in a row switch it off
    job.aiFailures = (job.aiFailures || 0) + 1;
    if (job.aiFailures >= 2) job.aiDisabled = true;
  }
  if (firstError) log(job, `Could not arrange ${failedPhotos} photo(s) in "${label}" (${firstError}); they were placed under other.`);
  if (timedOut) log(job, `Photo ordering for "${label}" stopped after ${Math.round(ORDERING_BUDGET_MS / 1000)} seconds; photos not yet placed went under other.`);
  if (job.aiDisabled && !job.aiDisabledLogged) {
    job.aiDisabledLogged = true;
    log(job, "Photo ordering is switched off for the rest of this run because the AI service kept failing. Check the key with the (test) link.");
  }
}

// items: [{ buf, ext }] in PDF order. Returns them with a .view label, sorted front, left, right,
// back, other; photos with the same view keep their original relative order. preclassify (optional)
// can place a photo without the AI (by file name, or because it is tiny or flat).
async function orderByVehicleView(items, workDir, job, label, preclassify) {
  const ai = aiConfigFor(job);
  const views = new Array(items.length).fill(null);
  const todo = [];
  for (let i = 0; i < items.length; i++) {
    const known = preclassify ? await preclassify(items[i]) : null;
    if (known) views[i] = known;
    else todo.push(i);
  }
  if (todo.length > 0 && !job.aiDisabled) {
    await classifyInBatches(items, todo, views, ai, workDir, job, label);
  }
  return items
    .map((item, idx) => ({ ...item, view: views[idx] || "other", idx }))
    .sort((a, b) => VIEW_ORDER.indexOf(a.view) - VIEW_ORDER.indexOf(b.view) || a.idx - b.idx);
}

// Filename-first guess: if a photo's own file name says what it shows (front, rear, left, interior...),
// use that and skip the AI call. Returns a view, or null when the name says nothing useful.
function viewFromFilename(fileName) {
  const tokens = new Set(
    path.basename(fileName, path.extname(fileName)).toLowerCase().split(/[^a-z]+/).filter(Boolean)
  );
  const has = (...words) => words.some((w) => tokens.has(w));
  if (has("interior", "dashboard", "dash", "engine", "bay", "tyre", "tire", "wheel", "odometer", "meter", "seat", "boot", "trunk", "chassis", "vin", "rc", "document", "doc")) return "other";
  if (has("front", "frnt", "fr")) return "front";
  if (has("rear", "back")) return "back";
  if (has("left", "lhs")) return "left";
  if (has("right", "rhs")) return "right";
  return null;
}

const ORDERED_NAME = /\/img-\d+-(front|left|right|back|other)\./;
const PHOTO_EXTS = new Set([".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff"]);

// Wraps the zip so that any photo a link handler adds gets arranged front, left, right, back, other
// once the handler is done. Everything else (page renders, PDFs, error notes, photos the PDF path
// already arranged) goes straight through. Only used when an AI key is set; otherwise handlers get
// the real zip and behave exactly as before.
function makeOrderingArchive(real, workDir, label, job) {
  const pending = [];
  return {
    append(data, opts) {
      const name = (opts && opts.name) || "";
      const ext = path.extname(name).toLowerCase();
      const isPhoto = Buffer.isBuffer(data) && PHOTO_EXTS.has(ext);
      if (isPhoto && !ORDERED_NAME.test(name) && !/\/(page-|report-page-)/.test(name)) {
        pending.push({ buf: data, ext, name });
        return;
      }
      return real.append(data, opts);
    },
    async flush() {
      if (pending.length === 0) return;
      if (pending.length === 1) {
        real.append(pending[0].buf, { name: pending[0].name });
        return;
      }
      let ordered;
      try {
        ordered = await orderByVehicleView(pending, workDir, job, label, async (item) => {
          const byName = viewFromFilename(path.basename(item.name));
          if (byName) return byName;
          const dims = readImageDimensions(item.buf);
          if (dims && Math.max(dims.width, dims.height) < MIN_EMBEDDED_IMAGE_DIMENSION) return "other";
          const colors = await countDistinctColors(item.buf, item.ext);
          if (colors !== null && colors < MIN_DISTINCT_COLORS) return "other";
          return null;
        });
      } catch (err) {
        log(job, `Could not arrange photos in "${label}": ${err.message}`);
        pending.forEach((it) => real.append(it.buf, { name: it.name }));
        return;
      }
      const width = Math.max(2, String(ordered.length).length);
      ordered.forEach((it, n) => {
        real.append(it.buf, { name: `${label}/img-${String(n + 1).padStart(width, "0")}-${it.view}${it.ext}` });
      });
      const counts = VIEW_ORDER.map((v) => `${ordered.filter((o) => o.view === v).length} ${v}`).join(", ");
      log(job, `Arranged photos in "${label}": ${counts}.`);
    },
  };
}

async function saveFetchedContent(buf, contentType, workDir, label, archive, seenHashes, job) {
  if (contentType.includes("application/zip") || contentType.includes("application/octet-stream")) {
    const isZip = buf.length > 4 && buf[0] === 0x50 && buf[1] === 0x4b; // "PK" zip signature
    if (isZip) {
      const AdmZip = (await import("adm-zip")).default;
      const zip = new AdmZip(buf);
      const zipEntries = zip.getEntries();
      let kept = 0, skippedType = 0, skippedDupe = 0;
      for (const entry of zipEntries) {
        if (entry.isDirectory) continue;
        const ext = path.extname(entry.entryName).toLowerCase();
        if (SKIP_EXTENSIONS.has(ext)) { skippedType++; continue; }
        if (![".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".pdf"].includes(ext)) continue;
        const entryBuf = entry.getData();
        const hash = hashBuffer(entryBuf);
        if (seenHashes.has(hash)) { skippedDupe++; continue; }
        seenHashes.add(hash);
        archive.append(entryBuf, { name: `${label}/${path.basename(entry.entryName)}` });
        kept++;
      }
      if (skippedType) log(job, `Skipped ${skippedType} .svg/.webp file(s) in "${label}".`);
      if (skippedDupe) log(job, `Skipped ${skippedDupe} duplicate file(s) in "${label}".`);
      log(job, `Extracted ${kept} file(s) from shared folder "${label}".`);
      return kept;
    }
  }

  if (contentType.includes("application/pdf") || buf.slice(0, 4).toString() === "%PDF") {
    const pdfPath = path.join(workDir, `${label}-${uuid()}.pdf`);
    fs.writeFileSync(pdfPath, buf);
    const converter = fromPath(pdfPath, { density: 200, saveFilename: label, savePath: workDir, format: "png", width: 1600, height: 2200 });
    const pageCount = await getPdfPageCount(pdfPath);
    let kept = 0;
    for (let p = 1; p <= pageCount; p++) {
      const out = await converter(p, { responseType: "image" });
      const pageBuf = fs.readFileSync(out.path);
      const hash = hashBuffer(pageBuf);
      if (seenHashes.has(hash)) { log(job, `Skipped duplicate page image in "${label}" (page ${p}).`); continue; }
      seenHashes.add(hash);
      archive.append(pageBuf, { name: `${label}/page-${p}.png` });
      kept++;
    }

    // In addition to full-page renders above, also pull out the individual photos embedded
    // inside the PDF itself (e.g. inspection photos), as their own separate image files.
    try {
      const extractPrefix = path.join(workDir, `${label}-embedded-${uuid()}`);
      await execFileAsync("pdfimages", ["-all", pdfPath, extractPrefix]);
      const dir = path.dirname(extractPrefix);
      const prefixName = path.basename(extractPrefix);
      const extractedFiles = fs.readdirSync(dir).filter((f) => f.startsWith(prefixName)).sort();
      let embeddedKept = 0, embeddedSkippedType = 0, embeddedSkippedDupe = 0, embeddedSkippedSmall = 0, embeddedSkippedFlat = 0;
      const keptImages = [];
      for (const file of extractedFiles) {
        const ext = path.extname(file).toLowerCase();
        if (SKIP_EXTENSIONS.has(ext)) { embeddedSkippedType++; continue; }
        const filePath = path.join(dir, file);
        const fileBuf = fs.readFileSync(filePath);

        // Filter out small template chrome (icons, logos, flat-color bars, QR codes) — real
        // content photos are almost always much larger on at least one side.
        const dims = readImageDimensions(fileBuf);
        if (dims && Math.max(dims.width, dims.height) < MIN_EMBEDDED_IMAGE_DIMENSION) {
          embeddedSkippedSmall++;
          continue;
        }

        // Filter out large-but-flat graphics (logos, banners, solid bars) that pass the size
        // check above but still aren't real photo content — real photos have far more
        // distinct colors than a flat-colored logo or icon, however big it's rendered.
        const distinctColors = await countDistinctColors(fileBuf, ext);
        if (distinctColors !== null && distinctColors < MIN_DISTINCT_COLORS) {
          embeddedSkippedFlat++;
          continue;
        }

        const hash = hashBuffer(fileBuf);
        if (seenHashes.has(hash)) { embeddedSkippedDupe++; continue; }
        seenHashes.add(hash);
        keptImages.push({ buf: fileBuf, ext });
        embeddedKept++;
      }

      // Write the kept photos into the zip. With an AI key set, they are arranged
      // front, left, right, back, other; without it they stay in PDF order (original naming).
      if (aiConfigFor(job) && keptImages.length > 1) {
        const ordered = await orderByVehicleView(keptImages, workDir, job, label);
        const width = Math.max(2, String(ordered.length).length);
        ordered.forEach((item, n) => {
          archive.append(item.buf, { name: `${label}/img-${String(n + 1).padStart(width, "0")}-${item.view}${item.ext}` });
        });
        const counts = VIEW_ORDER.map((v) => `${ordered.filter((o) => o.view === v).length} ${v}`).join(", ");
        log(job, `Arranged photos in "${label}": ${counts}.`);
      } else {
        keptImages.forEach((item, n) => {
          archive.append(item.buf, { name: `${label}/img-${n + 1}${item.ext}` });
        });
        if (keptImages.length > 1 && !job.orderNoteLogged) {
          job.orderNoteLogged = true;
          log(job, "Photos kept in PDF order. Add ANTHROPIC_API_KEY (Claude) or GEMINI_API_KEY (Google, has a free tier) on the server to arrange them front, left, right, back.");
        }
      }

      if (embeddedSkippedType) log(job, `Skipped ${embeddedSkippedType} embedded .svg/.webp image(s) in "${label}".`);
      if (embeddedSkippedSmall) log(job, `Skipped ${embeddedSkippedSmall} small embedded image(s) in "${label}" (likely icons/logos, under ${MIN_EMBEDDED_IMAGE_DIMENSION}px).`);
      if (embeddedSkippedFlat) log(job, `Skipped ${embeddedSkippedFlat} flat-color embedded image(s) in "${label}" (likely logos/banners).`);
      if (embeddedSkippedDupe) log(job, `Skipped ${embeddedSkippedDupe} duplicate embedded image(s) in "${label}".`);
      log(job, `Extracted ${embeddedKept} embedded image(s) from PDF "${label}" (in addition to ${kept} page render(s)).`);
      kept += embeddedKept;
    } catch (err) {
      log(job, `Could not extract embedded images from "${label}" PDF: ${err.message}`);
    }

    return kept;
  }

  if (contentType.startsWith("image/")) {
    const ext = "." + (contentType.split("/")[1] || "jpg").split(";")[0];
    if (SKIP_EXTENSIONS.has(ext)) { log(job, `Skipped .svg/.webp file in "${label}".`); return 0; }
    const hash = hashBuffer(buf);
    if (seenHashes.has(hash)) { log(job, `Skipped duplicate file in "${label}".`); return 0; }
    seenHashes.add(hash);
    archive.append(buf, { name: `${label}/image${ext}` });
    return 1;
  }

  throw new Error(`Unrecognized content-type for direct download: ${contentType || "unknown"}`);
}

async function handleNextcloudShare(url, reqHeaders, workDir, label, archive, seenHashes, job) {
  const downloadUrl = url.replace(/\/+$/, "") + "/download";
  const headers = { ...BROWSER_LIKE_HEADERS, ...reqHeaders };
  const resp = await fetch(downloadUrl, { headers, redirect: "follow", signal: AbortSignal.timeout(180000) });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  const contentType = (resp.headers.get("content-type") || "").toLowerCase();
  const buf = Buffer.from(await resp.arrayBuffer());
  return saveFetchedContent(buf, contentType, workDir, label, archive, seenHashes, job);
}

async function handleGenericDownload(url, reqHeaders, workDir, label, archive, seenHashes, job) {
  const headers = { ...BROWSER_LIKE_HEADERS, ...reqHeaders };
  const resp = await fetch(url, { headers, redirect: "follow", signal: AbortSignal.timeout(180000) });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  const contentType = (resp.headers.get("content-type") || "").toLowerCase();
  const buf = Buffer.from(await resp.arrayBuffer());
  return saveFetchedContent(buf, contentType, workDir, label, archive, seenHashes, job);
}

// ---------- report-generation action links (e.g. ?action=generate_report&...) ----------
// These endpoints may return either a raw file (PDF/image/zip) or an HTML report page
// depending on the vendor, so this handler sniffs the actual response and adapts.
function isReportActionLink(url) {
  try {
    const action = (new URL(url).searchParams.get("action") || "").toLowerCase();
    return action.includes("report");
  } catch {
    return false;
  }
}

// Used only by the smart-report handler below: does everything handleHtmlPage does (extract
// embedded <img> tags) PLUS captures a full-page screenshot of the rendered report itself, so
// both the individual images and a "whole report page" image end up in the zip. Kept as a
// separate function so handleHtmlPage (used by the admin-report link) stays untouched.
async function handleHtmlPageWithScreenshot(url, reqHeaders, browser, label, archive, seenHashes, job) {
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

  // Give lazy-loaded images (which only fetch their real src on scroll/intersection) a chance to load.
  await page.evaluate(async () => {
    window.scrollTo(0, document.body.scrollHeight);
    await new Promise((r) => setTimeout(r, 800));
    window.scrollTo(0, 0);
  });
  await new Promise((r) => setTimeout(r, 300));

  // Scan every frame on the page (the main page plus any iframes — Puppeteer can read into
  // iframes even cross-origin ones, unlike a normal in-page script). For each frame, collect
  // both <img> sources (including lazy-load data-* attributes) and CSS background-image URLs,
  // since dashboard-style reports often use one or the other instead of plain <img> tags.
  const candidates = []; // { rawUrl, frameOrigin, frameUrl }
  for (const frame of page.frames()) {
    if (frame.isDetached()) continue;
    let frameUrl;
    try {
      frameUrl = frame.url();
      if (!frameUrl || frameUrl === "about:blank") continue;
    } catch {
      continue;
    }
    let frameOrigin;
    try {
      frameOrigin = new URL(frameUrl).origin;
    } catch {
      continue;
    }

    let imgSrcs = [];
    let bgUrls = [];
    try {
      imgSrcs = await frame.$$eval("img", (imgs) =>
        imgs
          .map((img) => img.currentSrc || img.src || img.dataset.src || img.dataset.lazySrc || img.dataset.original || "")
          .filter((src) => src && !src.startsWith("about:blank"))
      );
    } catch {
      // frame may have navigated away or be inaccessible — skip its <img> tags
    }
    try {
      bgUrls = await frame.evaluate(() => {
        const found = [];
        document.querySelectorAll("*").forEach((el) => {
          const bg = getComputedStyle(el).backgroundImage;
          if (bg && bg !== "none") {
            const matches = bg.match(/url\(["']?([^"')]+)["']?\)/g) || [];
            matches.forEach((m) => {
              const u = m.replace(/^url\(["']?/, "").replace(/["']?\)$/, "");
              if (u) found.push(u);
            });
          }
        });
        return found;
      });
    } catch {
      // frame may not allow style inspection (rare) — skip background-image scan for it
    }

    for (const src of [...imgSrcs, ...bgUrls]) {
      candidates.push({ rawUrl: src, frameOrigin, frameUrl });
    }
  }

  // Resolve to absolute URLs (relative to the frame that referenced them) and drop exact-URL duplicates.
  const seenUrls = new Set();
  const resolved = [];
  for (const c of candidates) {
    let absoluteUrl;
    try {
      absoluteUrl = c.rawUrl.startsWith("data:") ? c.rawUrl : new URL(c.rawUrl, c.frameUrl).href;
    } catch {
      continue;
    }
    if (seenUrls.has(absoluteUrl)) continue;
    seenUrls.add(absoluteUrl);
    resolved.push({ url: absoluteUrl, frameOrigin: c.frameOrigin });
  }

  let kept = 0;
  let skippedType = 0;
  let skippedDupe = 0;
  let failedFetch = 0;
  const failureSamples = [];

  for (let i = 0; i < resolved.length; i++) {
    const { url: imgUrl, frameOrigin } = resolved[i];
    try {
      if (imgUrl.startsWith("data:")) {
        // Inline base64 image — no fetch needed, decode directly.
        const match = imgUrl.match(/^data:image\/([a-zA-Z0-9.+-]+);base64,(.+)$/);
        if (!match) continue;
        const ext = "." + match[1].split("+")[0];
        if (SKIP_EXTENSIONS.has(ext)) { skippedType++; continue; }
        const buf = Buffer.from(match[2], "base64");
        const hash = hashBuffer(buf);
        if (seenHashes.has(hash)) { skippedDupe++; continue; }
        seenHashes.add(hash);
        archive.append(buf, { name: `${label}/img-${i + 1}${ext}` });
        kept++;
        continue;
      }

      const rawExt = path.extname(new URL(imgUrl).pathname).split("?")[0].toLowerCase();
      if (SKIP_EXTENSIONS.has(rawExt)) {
        skippedType++;
        continue;
      }
      const resp = await fetch(imgUrl, {
        headers: { ...BROWSER_LIKE_HEADERS, Referer: frameOrigin, ...reqHeaders },
      });
      if (!resp.ok) {
        failedFetch++;
        if (failureSamples.length < 3) failureSamples.push(`HTTP ${resp.status} on ${imgUrl}`);
        continue;
      }
      const buf = Buffer.from(await resp.arrayBuffer());
      const hash = hashBuffer(buf);
      if (seenHashes.has(hash)) {
        skippedDupe++;
        continue;
      }
      seenHashes.add(hash);
      const ext = rawExt || ".jpg";
      archive.append(buf, { name: `${label}/img-${i + 1}${ext}` });
      kept++;
    } catch (err) {
      failedFetch++;
      if (failureSamples.length < 3) failureSamples.push(`${err.message} on ${imgUrl}`);
    }
  }

  // Capture the rendered report itself as a full-page image, in addition to the embedded images above.
  let pageShots = 0;
  try {
    const shotBuf = await page.screenshot({ fullPage: true, type: "png" });
    const shotHash = hashBuffer(shotBuf);
    if (!seenHashes.has(shotHash)) {
      seenHashes.add(shotHash);
      archive.append(shotBuf, { name: `${label}/report-page-1.png` });
      pageShots = 1;
    }
  } catch (err) {
    log(job, `Could not capture full-page screenshot for "${label}": ${err.message}`);
  }

  if (skippedType) log(job, `Skipped ${skippedType} .svg/.webp image(s) in "${label}".`);
  if (skippedDupe) log(job, `Skipped ${skippedDupe} duplicate image(s) in "${label}".`);
  if (failedFetch) {
    log(job, `${failedFetch} embedded image(s) in "${label}" failed to download. Examples: ${failureSamples.join("; ")}`);
  }
  if (resolved.length === 0) {
    log(job, `No images found on "${label}" (checked <img> tags, lazy-load attributes, CSS backgrounds, and iframes).`);
  }
  log(job, `Saved ${kept} embedded image(s) and ${pageShots} report-page screenshot(s) for "${label}".`);
  await page.close();
  return kept + pageShots;
}

async function handleSmartReport(url, reqHeaders, workDir, label, archive, seenHashes, job, ensureBrowser) {
  const headers = { ...BROWSER_LIKE_HEADERS, ...reqHeaders };
  const resp = await fetch(url, { headers, redirect: "follow", signal: AbortSignal.timeout(180000) });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  const contentType = (resp.headers.get("content-type") || "").toLowerCase();
  const buf = Buffer.from(await resp.arrayBuffer());

  if (contentType.includes("text/html")) {
    log(job, `"${label}" returned an HTML report page — rendering it in a browser to pull images and the report page itself.`);
    const browser = await ensureBrowser();
    return handleHtmlPageWithScreenshot(url, reqHeaders, browser, label, archive, seenHashes, job);
  }

  return saveFetchedContent(buf, contentType, workDir, label, archive, seenHashes, job);
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
  async function ensureBrowser() {
    browser = browser || (await puppeteer.launch({ args: ["--no-sandbox"] }));
    return browser;
  }
  const seenHashes = new Set();
  try {
    for (let i = 0; i < entries.length; i++) {
      const entry = entries[i];
      const label = sanitizeFolderName(entry.folder, `row-${i + 1}`);
      const url = String(entry.url).trim();
      job.current = i + 1;
      log(job, `Fetching "${label}"...`);
      const entryArchive = aiConfigFor(job) ? makeOrderingArchive(archive, workDir, label, job) : archive;
      try {
        if (isNextcloudShare(url)) {
          await handleNextcloudShare(url, reqHeaders, workDir, label, entryArchive, seenHashes, job);
        } else if (url.toLowerCase().includes(".pdf")) {
          await handlePdf(url, reqHeaders, workDir, label, entryArchive, seenHashes, job);
        } else if (isDirectDownloadLink(url)) {
          await handleGenericDownload(url, reqHeaders, workDir, label, entryArchive, seenHashes, job);
        } else if (isReportActionLink(url)) {
          await handleSmartReport(url, reqHeaders, workDir, label, entryArchive, seenHashes, job, ensureBrowser);
        } else {
          const br = await ensureBrowser();
          await handleHtmlPage(url, reqHeaders, br, label, entryArchive, seenHashes, job);
        }
        log(job, `Done: "${label}"`);
      } catch (err) {
        log(job, `Failed "${label}": ${err.message}`);
        archive.append(`Failed to fetch ${url}: ${err.message}`, { name: `${label}-ERROR.txt` });
      }
      if (entryArchive !== archive) {
        try {
          await entryArchive.flush();
        } catch (err) {
          log(job, `Could not arrange photos in "${label}": ${err.message}`);
        }
      }
    }
  } finally {
    if (browser) await browser.close();
    archive.finalize();
    await zipDone;
    fs.rm(workDir, { recursive: true, force: true }, () => {});
  }

  job.status = "done";
  job.ai = null; // the pasted key is not kept after the job
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
  const aiKey = req.body && req.body.aiKey;

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
    .map((r) => ({ folder: r[0], url: String(r[1] || "").trim() }))
    .filter((r) => r.url && /^https?:\/\//i.test(String(r.url)));

  if (entries.length === 0) {
    return res.status(400).json({ error: "No valid rows found. Column A = folder name, Column B = URL." });
  }

  const jobId = startJob(entries);
  jobs.get(jobId).ai = aiConfigFromKey(aiKey);
  runJob(jobId, entries, cookie).catch((err) => {
    const job = jobs.get(jobId);
    if (job) {
      job.status = "error";
      job.ai = null;
      job.error = err.message;
      log(job, `Job failed: ${err.message}`);
    }
  });
  res.json({ jobId });
});

app.post("/jobs/from-urls", async (req, res) => {
  const { urls, cookie, aiKey } = req.body || {};
  if (!Array.isArray(urls) || urls.length === 0) {
    return res.status(400).json({ error: "Provide 'urls' as a non-empty array." });
  }
  const entries = urls.map((u, i) => ({ folder: `item-${i + 1}`, url: String(u || "").trim() }));
  const jobId = startJob(entries);
  jobs.get(jobId).ai = aiConfigFromKey(aiKey);
  runJob(jobId, entries, cookie || "").catch((err) => {
    const job = jobs.get(jobId);
    if (job) {
      job.status = "error";
      job.ai = null;
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
      <div class="field" style="margin-bottom:18px;">
        <label>Photo ordering key (optional)</label>
        <input type="password" id="aiKey" placeholder="Paste a free Gemini key to arrange photos front, left, right, back" autocomplete="off" />
        <p class="hint">Stays in this browser and is sent only when you start a job. <a href="https://aistudio.google.com/apikey" target="_blank" style="color:var(--orange);">Get a free key</a> &middot; <a href="#" id="testKey" style="color:var(--orange);">Test key</a> <span id="keyStatus"></span></p>
      </div>

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
          <p class="hint">First row is auto-detected as a header and skipped if column B isn't a link. <a href="/sample-template.xlsx" style="color:var(--orange);">Download sample file</a> to see the expected format.</p>
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
      <span class="hint" style="margin:0;">Runs entirely on your own Render service. <!--ORDER_STATUS--></span>
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
    var aiKeyValue = getAiKey();
    if (aiKeyValue) formData.append('aiKey', aiKeyValue);
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
        body: JSON.stringify({ urls: urls, cookie: cookie || undefined, aiKey: getAiKey() || undefined })
      });
      const data = await resp.json();
      if (!resp.ok) throw new Error(data.error || 'Failed to start job');
      pollJob(data.jobId);
    } catch (err) {
      addLog('Error: ' + err.message); logEl.style.display='block';
      document.getElementById('goPaste').disabled = false;
    }
  };

  var aiKeyInput = document.getElementById('aiKey');
  try { aiKeyInput.value = localStorage.getItem('kuwy_ai_key') || ''; } catch (e) {}
  aiKeyInput.addEventListener('input', function () {
    try { localStorage.setItem('kuwy_ai_key', aiKeyInput.value.trim()); } catch (e) {}
  });
  function getAiKey() { return aiKeyInput.value.trim(); }
  document.getElementById('testKey').addEventListener('click', function (ev) {
    ev.preventDefault();
    var statusEl = document.getElementById('keyStatus');
    statusEl.textContent = ' checking...';
    fetch('/ordering-check', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ aiKey: getAiKey() }) })
      .then(function (r) { return r.text(); })
      .then(function (t) { statusEl.textContent = ' ' + t; })
      .catch(function (e) { statusEl.textContent = ' Could not reach the server: ' + e.message; });
  });

  resetBtn.onclick = resetUI;
</script>
</body>
</html>`;

app.get("/sample-template.xlsx", (_req, res) => {
  const rows = [
    ["Folder Name", "Image URL"],
    ["CC260113075222233154", "https://reports.kuwycarcheck.com/image/upload/v1771642553/KUWY/CARCHECK/REPORTS/2026/02/21/CC260113075222233154/pdf_image/CC260113075222233154_2026_02_21_08_25.pdf"],
    ["CC260113075222233154-admin", "https://report.kuwycarcheck.com/admin/inspectionReports?id=...&type=..."],
  ];
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  sheet["!cols"] = [{ wch: 28 }, { wch: 90 }];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "Reports");
  const buf = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename="sample-template.xlsx"`);
  res.send(buf);
});

async function checkAi(ai) {
  if (ai.provider === "gemini") {
    await callGemini([{ text: "Reply with the single word: ok" }], 3, ai.key);
    return `Photo ordering is ON and working (Google Gemini, model: ${await resolveGeminiModel(ai.key)}).`;
  }
  await callAnthropic(
    { model: orderingModel(), max_tokens: 5, messages: [{ role: "user", content: "Reply with the single word: ok" }] },
    3,
    ai.key
  );
  return `Photo ordering is ON and working (Claude, model: ${orderingModel()}).`;
}

async function respondToCheck(res, ai, offMessage) {
  res.type("text/plain");
  if (!ai) return res.send(offMessage);
  try {
    res.send(await checkAi(ai));
  } catch (err) {
    res.status(502).send(`Photo ordering key found (${ai.provider === "gemini" ? "Google Gemini" : "Claude"}) but the API call failed: ${err.message}`);
  }
}

app.get("/ordering-check", (_req, res) =>
  respondToCheck(
    res,
    envAiConfig(),
    "Photo ordering is OFF on the server: neither ANTHROPIC_API_KEY nor GEMINI_API_KEY is set.\n\nEasiest fix: paste a free Gemini key into the box at the top of the tool page (get one at aistudio.google.com). Or set GEMINI_API_KEY under your service's Environment tab in Render."
  )
);

app.post("/ordering-check", (req, res) =>
  respondToCheck(
    res,
    aiConfigFromKey(req.body && req.body.aiKey) || envAiConfig(),
    "No key to test: paste a key into the box first (a free Gemini key from aistudio.google.com works)."
  )
);

app.get("/", (_req, res) => {
  const orderingOn = Boolean(orderingProvider());
  const status = orderingOn
    ? 'Photo ordering: ON <a href="/ordering-check" target="_blank" style="color:var(--orange);">(test)</a>'
    : 'Photo ordering: OFF on the server (paste a key above, or set GEMINI_API_KEY in Render) <a href="/ordering-check" target="_blank" style="color:var(--orange);">(test)</a>';
  res.send(HOMEPAGE_HTML.replace("<!--ORDER_STATUS-->", status));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Listening on ${PORT}`));

