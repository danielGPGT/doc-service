# GPGT Document Generation Service

Stateless API that converts Canva-exported PPTX templates into personalised PDFs.

## What It Does

1. Receives PPTX page files + booking data (JSON)
2. Preprocesses each PPTX (fixes Canva's text fragmentation)
3. Expands list-of-link placeholders (see below) into hyperlinked paragraphs
4. Replaces `{d.xxx}` placeholder tags with booking data
5. Converts each page to PDF via LibreOffice
6. Merges all page PDFs into one final document
7. Returns the PDF

## Hyperlink lists (`{d.ticket_urls_links}`)

Plain `{d.xxx}` substitution is text-only — it cannot create a hyperlink or
repeat a paragraph, so a single tag can never become several clickable URLs.
The `{d.ticket_urls_links}` marker is handled specially: it is expanded into one
label paragraph + one **clickable** URL paragraph per item in
`booking_data["ticket_urls"]` (a list of `{"label": "...", "url": "..."}`).

The URL is set as an explicit run hyperlink, so it stays clickable even when it
wraps across lines (LibreOffice's auto-detection only links the first line), and
the visible URL text remains copy-pasteable. Authoring rules:

- Put the marker on **its own paragraph** in a text box (it replaces that paragraph).
- Link colour / underline / font size are inherited from the marker paragraph —
  style that paragraph to control how the links look.
- An empty/absent list just removes the marker (renders nothing).

## Endpoints

### `GET /health`
Health check. Returns LibreOffice availability.

### `POST /preprocess`
Upload a PPTX, get a preprocessed version + placeholder report.

```bash
curl -X POST https://your-service.railway.app/preprocess \
  -F "template=@cover.pptx"
```

Report only (no file download):
```bash
curl -X POST "https://your-service.railway.app/preprocess?report_only=true" \
  -F "template=@cover.pptx"
```

### `POST /render`
Render a single page PPTX with booking data, get a PDF.

```bash
curl -X POST https://your-service.railway.app/render \
  -F "template=@flights.pptx" \
  -F 'data={"lead_name": "James Mitchell", "airline": "British Airways"}' \
  -o flights.pdf
```

### `POST /render-batch`
Render multiple pages (uploaded as files), get a merged PDF.

```bash
curl -X POST https://your-service.railway.app/render-batch \
  -F "pages=@cover.pptx" \
  -F "pages=@client_info.pptx" \
  -F "pages=@flights.pptx" \
  -F "pages=@accommodation.pptx" \
  -F 'data={"lead_name": "James Mitchell", "hotel_name": "Hotel Milano"}' \
  -o final_document.pdf
```

### `POST /render-batch-from-urls`
Render multiple pages (fetched from URLs), get a merged PDF.
This is the endpoint called by the Supabase Edge Function.

```bash
curl -X POST https://your-service.railway.app/render-batch-from-urls \
  -H "Content-Type: application/json" \
  -d '{
    "pages": [
      {"url": "https://xxx.supabase.co/storage/v1/object/public/page-templates/cover.pptx"},
      {"url": "https://xxx.supabase.co/storage/v1/object/public/page-templates/flights.pptx"}
    ],
    "booking_data": {
      "lead_name": "James Mitchell",
      "hotel_name": "Hotel Milano"
    }
  }' \
  -o final_document.pdf
```

## Deploy to Railway

### Option 1: One-click from GitHub

1. Push this folder to a GitHub repo
2. Go to [railway.app](https://railway.app)
3. New Project → Deploy from GitHub Repo
4. Select the repo
5. Railway auto-detects the Dockerfile and deploys
6. Add environment variables if needed (see below)

### Option 2: Railway CLI

```bash
npm install -g @railway/cli
railway login
railway init
railway up
```

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `PORT` | 8080 | Server port (Railway sets this automatically) |
| `SOFFICE_PATH` | /usr/bin/soffice | Path to LibreOffice binary |
| `GPGT_API_KEY` | (none) | Optional API key for authentication |
| `MAX_FILE_SIZE` | 52428800 | Max upload size in bytes (default 50MB) |

### Setting the API Key

If you set `GPGT_API_KEY`, all requests (except /health) must include:
- Header: `X-API-Key: your-key-here`
- Or query param: `?api_key=your-key-here`

Recommended for production to prevent unauthorised use.

## Local Development

```bash
# Install dependencies
pip install -r requirements.txt

# Install LibreOffice
# Mac: brew install --cask libreoffice
# Ubuntu: sudo apt-get install libreoffice

# Run
python app.py

# Or with gunicorn
gunicorn --bind 0.0.0.0:8080 --timeout 300 app:app
```

## Cost

Free on Railway's free tier for your volume. The service is stateless — it only runs when processing requests.
