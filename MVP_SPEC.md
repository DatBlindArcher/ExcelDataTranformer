# Excel AI Transform — Project Specification

## Overview

An Office.js Task Pane Add-in for Microsoft Excel that uses AI (Claude or OpenAI) to automatically transform spreadsheet data. The user provides input data, an example of desired output, and optionally a ruleset or previous script. The tool returns transformed data written back into Excel and a reusable transformation script (VBA or Office Scripts).

The add-in frontend is static HTML/JS hosted on GitHub Pages. A lightweight proxy server is hosted separately on a VPS with nginx to forward AI API requests (bypassing CORS restrictions). The user provides their own AI API key — no secrets are stored server-side.

---

## Architecture

```
GitHub Pages (static hosting)              VPS (nginx + Node.js proxy)
┌──────────────────────────────┐          ┌─────────────────────────────┐
│  Excel Task Pane Add-in      │          │  /api/transform             │
│  ┌────────────────────────┐  │          │                             │
│  │  taskpane.html         │  │  HTTPS   │  - Receives:                │
│  │  taskpane.js (bundled) │──┼─────────▶│    • user's API key         │
│  │  taskpane.css          │  │          │    • AI provider choice     │
│  │                        │◀─┼──────────│    • prompt payload         │
│  │  Settings panel:       │  │          │  - Forwards to Claude/OpenAI│
│  │   • API key            │  │          │  - Returns AI response      │
│  │   • Provider select    │  │          │  - Stateless, no DB         │
│  │   • Proxy URL          │  │          │                             │
│  └────────────────────────┘  │          └─────────────────────────────┘
│                              │
│  manifest.xml                │
└──────────────────────────────┘
          ↕ Office.js API
     Excel Workbook
      • Read input range/tab
      • Read output example range/tab
      • Write transformed data
```

---

## Repository Structure

```
excel-ai-transform/
├── README.md                     # Setup & usage instructions
├── SPEC.md                       # This file
├── manifest.xml                  # Office Add-in manifest
├── .gitignore
│
├── taskpane/                     # Static frontend (deployed to GitHub Pages)
│   ├── index.html                # Main Task Pane HTML
│   ├── app.js                    # Application logic (vanilla JS, no build step)
│   ├── styles.css                # Styling
│   ├── libs/                     # Vendored libraries (no npm/build step needed)
│   │   └── office.js             # Office.js loader (or loaded via CDN)
│   └── assets/
│       ├── icon-16.png
│       ├── icon-32.png
│       ├── icon-80.png
│       └── icon-128.png
│
├── proxy/                        # Proxy server (deployed to VPS)
│   ├── package.json
│   ├── server.js                 # Express.js proxy server
│   └── ecosystem.config.js       # PM2 config for running on VPS
│
└── docs/                         # Optional: user-facing documentation
    └── usage-guide.md
```

### GitHub Pages Deployment

GitHub Pages should be configured to serve from the repository root **or** from `/taskpane` depending on preference. The simplest approach: configure GitHub Pages to serve from the repo root and place an `index.html` redirect at root, or serve directly from `/taskpane`.

**Recommended:** Configure GitHub Pages to serve from the root. The `manifest.xml` references URLs under `https://<user>.github.io/excel-ai-transform/taskpane/`.

---

## manifest.xml

The Office Add-in manifest tells Excel where to load the Task Pane from.

### Requirements

- `Id`: A unique GUID (generate one, e.g. `a1b2c3d4-e5f6-7890-abcd-ef1234567890`).
- `ProviderName`: Your name or company.
- `DisplayName`: `Excel AI Transform`
- `Description`: `Transform Excel data using AI with input/output examples.`
- `DefaultValue` of `SourceLocation`: `https://<user>.github.io/excel-ai-transform/taskpane/index.html`
- `SupportUrl` and `AppDomains`: Include your GitHub Pages domain and the proxy VPS domain.
- Icons: 16x16, 32x32, 80x80, 128x128 PNG files hosted on GitHub Pages.
- The manifest should define a **Ribbon button** in the Home tab that opens the Task Pane.
- Target: `Workbook` (Excel).
- API requirement set: `ExcelApi 1.7` minimum (supports range read/write, worksheet creation).
- The manifest schema version should be `1.1` using the `OfficeApp` XML namespace.

### Sideloading

The client loads this manifest via: Excel → Insert → My Add-ins → Upload My Add-in → browse to `manifest.xml`.

---

## Task Pane Frontend

### Technology

- **Vanilla JavaScript** — no build step, no React, no bundler. This keeps it simple to host on GitHub Pages without CI/CD.
- **Office.js** — loaded via CDN (`https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js`).
- **Single-page app** in `taskpane/index.html` with `taskpane/app.js` and `taskpane/styles.css`.

### UI Layout

The Task Pane is a vertical panel (~350px wide) inside Excel. Design it with a clean, modern look using a neutral color palette (whites, light grays, blue accents) consistent with Microsoft's Fluent UI style. Use system fonts (`Segoe UI, -apple-system, sans-serif`).

The UI has two views, toggled by tabs or a simple nav at the top:

#### 1. Settings View

Shown on first launch or when the user clicks a gear icon.

| Field | Type | Storage | Notes |
|---|---|---|---|
| AI Provider | Dropdown: `Claude`, `OpenAI` | `localStorage` | Default: `Claude` |
| API Key | Password input + show/hide toggle | `localStorage` | Stored in browser only, never sent to proxy logs |
| AI Model | Text input with smart defaults | `localStorage` | Default: `claude-sonnet-4-20250514` for Claude, `gpt-4o` for OpenAI. The user can override. |
| Proxy URL | Text input | `localStorage` | Default: empty. Placeholder: `https://your-proxy.example.com`. Required before use. |
| Script Language | Dropdown: `VBA`, `Office Scripts (TypeScript)` | `localStorage` | Default: `VBA` |

Include a **"Test Connection"** button that sends a minimal test request to the proxy to verify the API key and proxy URL work. Display a green checkmark or red error.

#### 2. Transform View (Main)

This is the primary working view.

**Section: Input Data**
- Label: "Input Data"
- A button: **"Use Current Selection"** — reads the currently selected range in Excel via `Office.js` and displays a preview (first 5 rows, truncated columns) in a small table below the button.
- A button: **"Use Entire Sheet"** — reads the active sheet.
- A dropdown: **"Select Sheet"** — lists all sheet names in the workbook; selecting one reads that sheet's used range.
- Display: A compact preview table showing row/column count and first few rows of captured data.
- Internal state: Store the captured data as a 2D array (array of arrays). The first row should be treated as headers.

**Section: Output Example**
- Same UI pattern as Input Data (selection, sheet, dropdown).
- Label: "Output Example"
- This is the example of what the transformed data should look like.

**Section: Additional Context (collapsible, default collapsed)**
- **Ruleset / Process Description**: A `<textarea>` where the user can describe transformation rules in plain English. Placeholder: `e.g. "Split the full name column into first name and last name, convert dates from MM/DD/YYYY to YYYY-MM-DD, remove rows where status is 'inactive'"`
- **Previous Script**: A `<textarea>` where the user can paste a previous VBA or Office Script that should be used as a starting point or reference. Placeholder: `Paste a previous VBA macro or Office Script here...`

**Section: Action**
- A large primary button: **"Transform"** — triggers the AI call.
- Below it, a status indicator: idle / "Analyzing data..." / "Generating transformation..." / "Writing results..." / error message.

**Section: Results (shown after successful transform)**
- **Transformed Data Preview**: A compact table showing first 5 rows of the result.
- **Write to Excel** button with a dropdown:
  - "New Sheet" (default) — creates a new sheet named `AI_Transform_<timestamp>` and writes the data there.
  - "Current Selection" — writes starting at the currently selected cell.
- **Transformation Script** panel:
  - A `<pre><code>` block with the generated VBA or Office Script.
  - A **"Copy Script"** button (copies to clipboard).
  - A **"Insert as VBA Module"** button (if VBA is selected — note: Office.js cannot directly insert VBA modules, so this button copies to clipboard and shows instructions: "Press Alt+F11, Insert → Module, paste the code").

### Office.js API Usage

#### Reading Data

```javascript
// Read from a selection
await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("values, address, rowCount, columnCount");
    await context.sync();
    // range.values is a 2D array
    return range.values;
});

// Read from a specific sheet
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("SheetName");
    const usedRange = sheet.getUsedRange();
    usedRange.load("values, rowCount, columnCount");
    await context.sync();
    return usedRange.values;
});

// List all sheet names
await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    return sheets.items.map(s => s.name);
});
```

#### Writing Data

```javascript
// Write to a new sheet
await Excel.run(async (context) => {
    const newSheet = context.workbook.worksheets.add("AI_Transform_Result");
    const startCell = newSheet.getRange("A1");
    const outputRange = startCell.getResizedRange(
        data.length - 1,
        data[0].length - 1
    );
    outputRange.values = data; // 2D array
    newSheet.activate();
    await context.sync();
});
```

### Constructing the AI Prompt

The `app.js` must construct a well-structured prompt to send to the AI. This is the core logic.

#### Prompt Template

```
You are an Excel data transformation assistant. You will be given:
1. INPUT DATA: The raw source data from an Excel spreadsheet (as CSV).
2. OUTPUT EXAMPLE: An example of what the transformed data should look like (as CSV).
3. (Optional) TRANSFORMATION RULES: Additional rules or descriptions of the transformation.
4. (Optional) PREVIOUS SCRIPT: A previously used script for reference.

Your task:
A) Analyze the input and output example to infer the transformation logic.
B) Apply that transformation to the FULL input data.
C) Generate a reusable script in {scriptLanguage} that performs this transformation.

Respond with EXACTLY this JSON structure (no markdown fences, no extra text):
{
  "transformedData": [
    ["Header1", "Header2", ...],
    ["row1col1", "row1col2", ...],
    ...
  ],
  "script": "...the full VBA or Office Script code as a string...",
  "explanation": "Brief explanation of the transformation logic applied."
}

--- INPUT DATA (CSV) ---
{inputDataAsCsv}

--- OUTPUT EXAMPLE (CSV) ---
{outputExampleAsCsv}

--- TRANSFORMATION RULES ---
{rulesOrEmpty}

--- PREVIOUS SCRIPT ---
{previousScriptOrEmpty}
```

#### Data Serialization

Convert the 2D arrays to CSV strings for the prompt. Use a simple CSV serializer:
- Wrap values containing commas, quotes, or newlines in double quotes.
- Escape double quotes by doubling them.
- Join columns with commas, rows with newlines.

**Important:** If the data is very large (>500 rows), include only the first 50 rows in the prompt as a representative sample, along with a note: `"(Showing first 50 of {totalRows} rows. Apply the transformation to all rows.)"`. Then, when receiving the AI response, if the AI only returns 50 transformed rows, apply the generated script logic client-side or re-run with the full dataset. For an MVP, it is acceptable to send up to ~200 rows and note this limitation to the user.

#### Parsing the AI Response

1. Strip any markdown code fences if present (` ```json ... ``` `).
2. `JSON.parse()` the response.
3. Validate that `transformedData` is a 2D array and `script` is a string.
4. If parsing fails, show the raw response to the user with an error message and a retry button.

### Error Handling

- **No API key configured**: Show a clear message directing to Settings.
- **No proxy URL configured**: Same.
- **Proxy unreachable**: Show "Cannot reach proxy server. Check your proxy URL and that the server is running."
- **AI API error (401)**: "Invalid API key. Check your key in Settings."
- **AI API error (429)**: "Rate limited. Please wait a moment and try again."
- **AI API error (500+)**: "AI service error. Try again or switch providers."
- **Response parse failure**: Show raw response, offer retry.
- **Empty selection**: "No data selected. Please select a range or choose a sheet."
- **Office.js not ready**: Wrap all Office calls in `Office.onReady()`.

### Styling Guidelines

- Font: `Segoe UI, -apple-system, BlinkMacSystemFont, sans-serif`
- Background: `#ffffff`
- Section backgrounds: `#f5f5f5` with `8px` border-radius
- Primary button: `#0078d4` (Microsoft blue), white text, `8px` border-radius
- Secondary buttons: `#f0f0f0` background, `#333` text
- Error text: `#d13438`
- Success text: `#107c10`
- Spacing: `12px` padding between sections
- Preview tables: bordered, compact, max-height `200px` with overflow scroll
- Code block: `monospace` font, `#1e1e1e` background, `#d4d4d4` text (VS Code-like)
- The entire pane must be scrollable vertically.

---

## Proxy Server

### Technology

- **Node.js** with **Express.js**
- Minimal dependencies: `express`, `cors`, `node-fetch` (or native fetch in Node 18+)
- No database, no authentication, no state. It is a pure pass-through proxy.

### Endpoints

#### `POST /api/transform`

**Request body (JSON):**

```json
{
  "provider": "claude" | "openai",
  "apiKey": "sk-...",
  "model": "claude-sonnet-4-20250514",
  "prompt": "...the full prompt string..."
}
```

**Behavior:**

1. Validate that `provider`, `apiKey`, and `prompt` are present. Return `400` if not.
2. Based on `provider`, forward to the appropriate API:

   **Claude (Anthropic):**
   ```
   POST https://api.anthropic.com/v1/messages
   Headers:
     x-api-key: {apiKey}
     anthropic-version: 2023-06-01
     content-type: application/json
   Body:
     {
       "model": "{model}",
       "max_tokens": 16384,
       "messages": [
         { "role": "user", "content": "{prompt}" }
       ]
     }
   ```

   **OpenAI:**
   ```
   POST https://api.openai.com/v1/chat/completions
   Headers:
     Authorization: Bearer {apiKey}
     content-type: application/json
   Body:
     {
       "model": "{model}",
       "max_tokens": 16384,
       "messages": [
         { "role": "user", "content": "{prompt}" }
       ]
     }
   ```

3. Normalize the response:

   **Response body (JSON):**
   ```json
   {
     "success": true,
     "content": "...the text content from the AI response..."
   }
   ```
   
   For Claude, extract: `response.content[0].text`
   For OpenAI, extract: `response.choices[0].message.content`

4. On error from the upstream API, forward the status code and return:
   ```json
   {
     "success": false,
     "error": "Error message",
     "status": 401
   }
   ```

#### `GET /api/health`

Returns `{ "status": "ok" }`. Used by the Task Pane "Test Connection" feature.

#### `POST /api/test`

**Request body:**
```json
{
  "provider": "claude" | "openai",
  "apiKey": "sk-...",
  "model": "claude-sonnet-4-20250514"
}
```

Sends a minimal request (e.g. `"Respond with: ok"`) to the AI API to validate the key works. Returns `{ "success": true }` or `{ "success": false, "error": "..." }`.

### CORS Configuration

The proxy must set CORS headers to allow requests from GitHub Pages:

```javascript
app.use(cors({
    origin: [
        'https://<user>.github.io',
        'http://localhost:3000',       // for local dev
        'https://localhost:3000'
    ],
    methods: ['POST', 'GET', 'OPTIONS'],
    allowedHeaders: ['Content-Type']
}));
```

**Note:** The implementer should replace `<user>` with the actual GitHub username. This should be configurable via environment variable.

### Rate Limiting

Add basic rate limiting to prevent abuse (since the proxy URL could be discovered in the public repo):

- Use `express-rate-limit` middleware.
- Limit: 30 requests per minute per IP.
- Return `429` with message when exceeded.

### Logging

- Log each request: timestamp, provider, model, IP (no API key, no prompt content).
- Use `console.log` — keep it simple. Nginx access logs will capture the rest.

### Deployment on VPS

#### PM2 Process Manager

Use PM2 to keep the server running. Include `ecosystem.config.js`:

```javascript
module.exports = {
    apps: [{
        name: 'excel-ai-proxy',
        script: 'server.js',
        instances: 1,
        env: {
            NODE_ENV: 'production',
            PORT: 3100,
            ALLOWED_ORIGINS: 'https://<user>.github.io'
        }
    }]
};
```

#### Nginx Configuration

The VPS should have nginx configured as a reverse proxy with HTTPS (via Let's Encrypt):

```nginx
server {
    listen 443 ssl;
    server_name proxy.yourdomain.com;

    ssl_certificate /etc/letsencrypt/live/proxy.yourdomain.com/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/proxy.yourdomain.com/privkey.pem;

    location /api/ {
        proxy_pass http://127.0.0.1:3100;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_read_timeout 120s;  # AI responses can be slow
        proxy_send_timeout 120s;
    }
}
```

**Important:** `proxy_read_timeout` must be at least `120s` because AI API calls for large transformations may take 30–60 seconds.

---

## .gitignore

```
node_modules/
.env
*.log
.DS_Store
Thumbs.db
```

No secrets exist in the repo. The `ALLOWED_ORIGINS` env var in `ecosystem.config.js` is not secret (it's just the public GitHub Pages URL). API keys are only ever held in the user's browser `localStorage` and sent per-request.

---

## Security Considerations

1. **API keys are never stored server-side.** They pass through the proxy in-memory only and are not logged.
2. **The proxy is stateless.** No database, no sessions, no cookies.
3. **Rate limiting** protects against abuse of the public proxy endpoint.
4. **CORS** restricts which origins can call the proxy.
5. **HTTPS is required** on both GitHub Pages (automatic) and the proxy (nginx + Let's Encrypt).
6. **The proxy should never log request bodies** to avoid accidentally logging API keys or sensitive data.

---

## Implementation Sequence

An AI implementing this project should follow this order:

1. **`proxy/package.json`** — Initialize with dependencies: `express`, `cors`, `express-rate-limit`.
2. **`proxy/server.js`** — Implement the proxy server with all three endpoints (`/api/health`, `/api/test`, `/api/transform`), CORS, rate limiting, and response normalization.
3. **`proxy/ecosystem.config.js`** — PM2 config.
4. **`taskpane/styles.css`** — Full stylesheet following the styling guidelines above.
5. **`taskpane/app.js`** — All application logic:
   - `Office.onReady()` initialization
   - Settings management (load/save to `localStorage`)
   - Excel data reading (selection, sheet, all sheets listing)
   - CSV serialization
   - Prompt construction
   - Proxy API communication
   - Response parsing
   - Excel data writing (new sheet or selection)
   - UI state management (show/hide sections, loading states, errors)
6. **`taskpane/index.html`** — The Task Pane HTML structure, loading Office.js via CDN, referencing `app.js` and `styles.css`.
7. **`manifest.xml`** — The Office Add-in manifest with a placeholder GitHub Pages URL (the implementer replaces `<user>` with their GitHub username).
8. **Icon assets** — Generate simple placeholder PNG icons at 16x16, 32x32, 80x80, 128x128.
9. **`README.md`** — Setup instructions covering: GitHub Pages deployment, VPS proxy setup, sideloading the manifest, first-time configuration in the add-in.
10. **`.gitignore`**

---

## Testing Checklist

Before considering the implementation complete, verify:

- [ ] `manifest.xml` is valid XML and passes the [Office Add-in manifest validator](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-manifest).
- [ ] Task Pane loads in Excel (via sideloading) and displays the Settings view.
- [ ] Settings (API key, provider, proxy URL, model, script language) persist across Task Pane reopens.
- [ ] "Test Connection" button works — green check on success, red error on failure.
- [ ] Sheet names populate in the dropdown.
- [ ] "Use Current Selection" captures selected range data and shows preview.
- [ ] "Use Entire Sheet" captures the used range and shows preview.
- [ ] Transform button sends correct prompt to proxy and receives response.
- [ ] Transformed data preview displays correctly.
- [ ] "Write to New Sheet" creates a new sheet with correct data.
- [ ] Generated script displays in code block with copy button working.
- [ ] Error states display appropriately (no key, no proxy, API error, parse error).
- [ ] Proxy server responds to `/api/health`.
- [ ] Proxy correctly forwards to both Claude and OpenAI APIs.
- [ ] Proxy rate limiting works (returns 429 after 30 requests/minute).
- [ ] CORS blocks requests from non-allowed origins.