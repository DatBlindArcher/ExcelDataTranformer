# Excel AI Transform

An Office.js Task Pane Add-in for Microsoft Excel that uses AI (Claude or OpenAI) to automatically transform spreadsheet data. Provide input data, an example of desired output, and the AI generates transformed data plus a reusable VBA/Office Script.

## Architecture

- **Task Pane frontend** — static HTML/JS hosted on GitHub Pages
- **Proxy server** — lightweight Node.js/Express proxy on a VPS (bypasses CORS)
- **User provides their own API key** — nothing stored server-side

## Quick Start

### 1. Deploy the Proxy Server

```bash
cd proxy
npm install
npm start
```

The proxy runs on port 3100 by default. For production, use PM2:

```bash
npm install -g pm2
pm2 start ecosystem.config.js
```

Set up nginx as a reverse proxy with HTTPS (Let's Encrypt). See the `MVP_SPEC.md` for the full nginx config. Key setting: `proxy_read_timeout 120s` since AI calls can be slow.

**Environment variables:**
- `PORT` — server port (default: `3100`)
- `ALLOWED_ORIGINS` — comma-separated allowed CORS origins (e.g. `https://yourusername.github.io`)

### 2. Deploy the Frontend

1. Push this repo to GitHub.
2. Enable GitHub Pages (Settings → Pages → Source: root or `/taskpane`).
3. Update `manifest.xml` — replace all `<user>` placeholders with your GitHub username.

### 3. Sideload the Add-in

1. Open Excel (desktop or web).
2. Go to **Insert → My Add-ins → Upload My Add-in**.
3. Browse to `manifest.xml` and upload.
4. The "AI Transform" button appears in the Home tab ribbon.

### 4. Configure the Add-in

On first launch, the Settings view opens:

1. **AI Provider** — Choose Claude or OpenAI.
2. **API Key** — Enter your API key (stored in browser localStorage only).
3. **Proxy URL** — Enter your proxy server URL (e.g. `https://proxy.yourdomain.com`).
4. **Model** — Optionally override the default model.
5. **Script Language** — VBA or Office Scripts (TypeScript).
6. Click **Test Connection** to verify everything works.

## Usage

1. **Capture Input Data** — Select a range or sheet containing your source data.
2. **Capture Output Example** — Select a range or sheet showing what the transformed data should look like (even a few rows is enough).
3. **(Optional)** Expand "Additional Context" to add transformation rules or a previous script.
4. Click **Transform**.
5. Review the transformed data preview and generated script.
6. Click **Write to New Sheet** to output the results.
7. Copy the generated script for future reuse.

## Project Structure

```
├── manifest.xml              # Office Add-in manifest
├── taskpane/
│   ├── index.html            # Task Pane UI
│   ├── app.js                # Application logic (vanilla JS)
│   ├── styles.css            # Styling
│   └── assets/               # Icon PNGs
├── proxy/
│   ├── package.json
│   ├── server.js             # Express proxy server
│   └── ecosystem.config.js   # PM2 config
└── MVP_SPEC.md               # Full specification
```

## Security

- API keys are stored in browser localStorage only — never sent to proxy logs.
- The proxy is stateless with no database.
- Rate limiting: 30 requests/minute per IP.
- CORS restricts which origins can call the proxy.
- HTTPS required on both frontend and proxy.
