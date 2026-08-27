# Bulk Image Downloader (Render-ready)

Turns a list of report URLs into a single `images.zip`:

- **Direct `.pdf` links** → each page is rasterized to a PNG (plus the original PDF is kept in the zip).
- **HTML report pages** → headless Chromium loads the page and every `<img>` it finds is downloaded.

## Deploy on Render

1. Push this folder to a GitHub repo.
2. In Render: **New → Blueprint**, point it at the repo (it will read `render.yaml`), or
   **New → Web Service → Docker** and point it at the same repo.
3. Deploy. Render will build the Docker image (installs GraphicsMagick/Ghostscript for PDF
   rasterizing and Chromium for Puppeteer) and start the service.

## Usage

```bash
curl -X POST https://YOUR-SERVICE.onrender.com/bulk-download \
  -H "Content-Type: application/json" \
  -o images.zip \
  -d '{
    "urls": [
      "https://reports.kuwycarcheck.com/image/upload/.../pdf_image/CC260113075222233154_2026_02_21_08_25.pdf",
      "https://report.kuwycarcheck.com/admin/inspectionReports?id=...&type=..."
    ],
    "cookie": "session=PASTE_YOUR_LOGGED_IN_SESSION_COOKIE_HERE"
  }'
```

### About the `cookie` field

The `/admin/inspectionReports` style URL is an authenticated admin route — it will not load
for an anonymous request. If it's a report you're entitled to view, log into the site in your
browser, open DevTools → Network, copy the `Cookie` header from a request to that page, and
pass it in the `cookie` field above (or use `headers.Authorization` if the site uses a bearer
token instead). Without valid credentials the tool will just save an `-ERROR.txt` note for
that URL inside the zip rather than guess at a login.

## Local test

```bash
npm install
npm start
# in another terminal
curl -X POST http://localhost:3000/bulk-download -H "Content-Type: application/json" \
  -o images.zip -d '{"urls":["https://example.com/report.pdf"]}'
```

## Notes / limits

- Batches are processed sequentially and streamed straight into the zip, so memory stays low
  even for many URLs — but very large batches will take a while on Render's free/starter tier.
- If a report site blocks headless browsers or requires solving a captcha, this approach won't
  bypass that — it only automates access you already legitimately have.
- Only point this at reports you're authorized to access; some sites' terms of service
  prohibit automated scraping even of your own purchased reports, so it's worth checking
  kuwycarcheck's terms before running this at scale.
