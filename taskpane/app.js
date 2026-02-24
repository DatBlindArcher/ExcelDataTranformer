/* ============================================================
   Excel AI Transform — Task Pane Application Logic
   ============================================================ */

(function () {
    'use strict';

    // ── State ──────────────────────────────────────────────────
    const state = {
        inputData: null,   // 2D array
        outputData: null,  // 2D array
        result: null       // { transformedData, script, explanation }
    };

    // ── Configuration ─────────────────────────────────────────
    const PROXY_URL = 'https://proxy.excel.archtech.be';

    function getScriptLanguage() {
        return localStorage.getItem('scriptLanguage') || 'VBA';
    }

    function saveScriptLanguage(value) {
        localStorage.setItem('scriptLanguage', value);
    }

    // ── CSV helpers ───────────────────────────────────────────
    function arrayToCsv(data) {
        return data.map(row =>
            row.map(cell => {
                const str = cell == null ? '' : String(cell);
                if (str.includes(',') || str.includes('"') || str.includes('\n')) {
                    return '"' + str.replace(/"/g, '""') + '"';
                }
                return str;
            }).join(',')
        ).join('\n');
    }

    // ── Prompt construction ───────────────────────────────────
    function buildPrompt(inputData, outputExample, rules, previousScript, scriptLanguage) {
        let inputCsv = arrayToCsv(inputData);
        let note = '';

        if (inputData.length > 200) {
            const sample = inputData.slice(0, 51); // header + 50 rows
            inputCsv = arrayToCsv(sample);
            note = `(Showing first 50 of ${inputData.length} rows. Apply the transformation to all rows.)`;
        }

        let outputCsv = arrayToCsv(outputExample);

        return `You are an Excel data transformation assistant. You will be given:
1. INPUT DATA: The raw source data from an Excel spreadsheet (as CSV).
2. OUTPUT EXAMPLE: An example of what the transformed data should look like (as CSV).
3. (Optional) TRANSFORMATION RULES: Additional rules or descriptions of the transformation.
4. (Optional) PREVIOUS SCRIPT: A previously used script for reference.

Your task:
A) Analyze the input and output example to infer the transformation logic.
B) Apply that transformation to the FULL input data.
C) Generate a reusable script in ${scriptLanguage} that performs this transformation.

Respond with EXACTLY this JSON structure (no markdown fences, no extra text):
{
  "transformedData": [
    ["Header1", "Header2", ...],
    ["row1col1", "row1col2", ...],
    ...
  ],
  "script": "...the full ${scriptLanguage} code as a string...",
  "explanation": "Brief explanation of the transformation logic applied."
}

--- INPUT DATA (CSV) ---
${inputCsv}
${note}

--- OUTPUT EXAMPLE (CSV) ---
${outputCsv}

--- TRANSFORMATION RULES ---
${rules || '(none provided)'}

--- PREVIOUS SCRIPT ---
${previousScript || '(none provided)'}`;
    }

    // ── AI response parsing ───────────────────────────────────
    function parseAIResponse(raw) {
        // Strip markdown code fences if present
        let cleaned = raw.trim();
        cleaned = cleaned.replace(/^```(?:json)?\s*\n?/i, '');
        cleaned = cleaned.replace(/\n?```\s*$/i, '');
        cleaned = cleaned.trim();

        const parsed = JSON.parse(cleaned);

        if (!Array.isArray(parsed.transformedData) || typeof parsed.script !== 'string') {
            throw new Error('Invalid response structure: missing transformedData array or script string.');
        }

        return parsed;
    }

    // ── Safe response parsing ──────────────────────────────────
    async function safeJsonParse(response) {
        const text = await response.text();
        try {
            return JSON.parse(text);
        } catch {
            // Likely an HTML error page from nginx or a misconfigured proxy URL
            if (text.trimStart().startsWith('<')) {
                throw new Error(
                    `Proxy returned an error page (HTTP ${response.status}). ` +
                    'Check that the Proxy URL is correct and the proxy service is running.'
                );
            }
            throw new Error(`Unexpected response from proxy (HTTP ${response.status}).`);
        }
    }

    // ── API communication ─────────────────────────────────────
    async function apiCall(endpoint, body) {
        const url = PROXY_URL + endpoint;
        const response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });

        const data = await safeJsonParse(response);
        if (!data.success) {
            const status = data.status || response.status;
            if (status === 429) throw new Error('Rate limited. Please wait a moment and try again.');
            if (status >= 500) throw new Error('AI service error. Try again later.');
            throw new Error(data.error || 'Unknown error from proxy.');
        }
        return data;
    }

    // ── Excel helpers ─────────────────────────────────────────
    async function getSheetNames() {
        return Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load('items/name');
            await context.sync();
            return sheets.items.map(s => s.name);
        });
    }

    async function readSelection() {
        return Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load('values, rowCount, columnCount');
            await context.sync();
            return { values: range.values, rows: range.rowCount, cols: range.columnCount };
        });
    }

    async function readSheet(sheetName) {
        return Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            const usedRange = sheet.getUsedRange();
            usedRange.load('values, rowCount, columnCount');
            await context.sync();
            return { values: usedRange.values, rows: usedRange.rowCount, cols: usedRange.columnCount };
        });
    }

    async function readActiveSheet() {
        return Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = sheet.getUsedRange();
            usedRange.load('values, rowCount, columnCount');
            await context.sync();
            return { values: usedRange.values, rows: usedRange.rowCount, cols: usedRange.columnCount };
        });
    }

    async function writeToNewSheet(data) {
        return Excel.run(async (context) => {
            const now = new Date();
            const ts = now.getFullYear().toString().slice(2)
                + String(now.getMonth() + 1).padStart(2, '0')
                + String(now.getDate()).padStart(2, '0')
                + '_'
                + String(now.getHours()).padStart(2, '0')
                + String(now.getMinutes()).padStart(2, '0')
                + String(now.getSeconds()).padStart(2, '0');
            const name = 'AI_Transform_' + ts;
            const newSheet = context.workbook.worksheets.add(name);
            const startCell = newSheet.getRange('A1');
            const outputRange = startCell.getResizedRange(data.length - 1, data[0].length - 1);
            outputRange.values = data;
            newSheet.activate();
            await context.sync();
            return name;
        });
    }

    async function writeToSelection(data) {
        return Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load('address');
            await context.sync();
            const startCell = context.workbook.worksheets.getActiveWorksheet().getRange(range.address.split(':')[0]);
            const outputRange = startCell.getResizedRange(data.length - 1, data[0].length - 1);
            outputRange.values = data;
            await context.sync();
        });
    }

    // ── UI helpers ────────────────────────────────────────────
    function $(selector) { return document.querySelector(selector); }

    function renderPreview(container, data) {
        if (!data || data.length === 0) {
            container.innerHTML = '<div class="empty-state">No data captured</div>';
            return;
        }
        const maxRows = Math.min(data.length, 6); // header + 5 data rows
        const maxCols = Math.min(data[0].length, 8);
        let html = '<table><thead><tr>';
        // Header row
        for (let c = 0; c < maxCols; c++) {
            html += `<th>${escapeHtml(String(data[0][c] ?? ''))}</th>`;
        }
        if (data[0].length > maxCols) html += '<th>...</th>';
        html += '</tr></thead><tbody>';
        // Data rows
        for (let r = 1; r < maxRows; r++) {
            html += '<tr>';
            for (let c = 0; c < maxCols; c++) {
                html += `<td>${escapeHtml(String(data[r][c] ?? ''))}</td>`;
            }
            if (data[0].length > maxCols) html += '<td>...</td>';
            html += '</tr>';
        }
        html += '</tbody></table>';
        container.innerHTML = html;
    }

    function escapeHtml(str) {
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }

    function showStatus(el, type, message) {
        el.className = `status visible status-${type}`;
        if (type === 'loading') {
            el.innerHTML = `<span class="spinner"></span>${escapeHtml(message)}`;
        } else {
            el.textContent = message;
        }
    }

    function hideStatus(el) {
        el.className = 'status';
        el.textContent = '';
    }

    function populateSheetDropdown(selectEl, sheetNames) {
        selectEl.innerHTML = '<option value="">-- Select Sheet --</option>';
        sheetNames.forEach(name => {
            const opt = document.createElement('option');
            opt.value = name;
            opt.textContent = name;
            selectEl.appendChild(opt);
        });
    }

    // ── Initialization ────────────────────────────────────────
    Office.onReady(function (info) {
        if (info.host === Office.HostType.Excel) {
            init();
        }
    });

    function init() {
        // Load script language preference
        $('#script-language').value = getScriptLanguage();

        // Load sheet names
        refreshSheetLists();

        bindEvents();
    }

    async function refreshSheetLists() {
        try {
            const names = await getSheetNames();
            populateSheetDropdown($('#input-sheet-select'), names);
            populateSheetDropdown($('#output-sheet-select'), names);
        } catch (e) {
            console.warn('Could not load sheet names:', e);
        }
    }

    // ── Event binding ─────────────────────────────────────────
    function bindEvents() {
        // Script language preference
        $('#script-language').addEventListener('change', (e) => {
            saveScriptLanguage(e.target.value);
        });

        // Input data buttons
        $('#btn-input-selection').addEventListener('click', () => captureData('input', 'selection'));
        $('#btn-input-sheet').addEventListener('click', () => captureData('input', 'activeSheet'));
        $('#input-sheet-select').addEventListener('change', (e) => {
            if (e.target.value) captureData('input', 'sheet', e.target.value);
        });

        // Output data buttons
        $('#btn-output-selection').addEventListener('click', () => captureData('output', 'selection'));
        $('#btn-output-sheet').addEventListener('click', () => captureData('output', 'activeSheet'));
        $('#output-sheet-select').addEventListener('change', (e) => {
            if (e.target.value) captureData('output', 'sheet', e.target.value);
        });

        // Collapsible
        $('.collapsible-header').addEventListener('click', () => {
            const chevron = $('.collapsible-header .chevron');
            const body = $('.collapsible-body');
            chevron.classList.toggle('open');
            body.classList.toggle('open');
        });

        // Transform
        $('#btn-transform').addEventListener('click', runTransform);

        // Write results
        $('#btn-write-new-sheet').addEventListener('click', () => writeResults('newSheet'));
        $('#btn-write-selection').addEventListener('click', () => writeResults('selection'));

        // Copy script
        $('#btn-copy-script').addEventListener('click', copyScript);

        // Refresh sheets on focus
        document.addEventListener('visibilitychange', () => {
            if (!document.hidden) refreshSheetLists();
        });
    }

    // ── Actions ───────────────────────────────────────────────
    async function captureData(target, mode, sheetName) {
        const statusEl = target === 'input' ? $('#input-status') : $('#output-status');
        const previewEl = target === 'input' ? $('#input-preview') : $('#output-preview');
        const infoEl = target === 'input' ? $('#input-info') : $('#output-info');

        showStatus(statusEl, 'loading', 'Reading data...');

        try {
            let result;
            if (mode === 'selection') {
                result = await readSelection();
            } else if (mode === 'activeSheet') {
                result = await readActiveSheet();
            } else if (mode === 'sheet') {
                result = await readSheet(sheetName);
            }

            if (!result.values || result.values.length === 0) {
                throw new Error('No data found. Select a range with data.');
            }

            if (target === 'input') {
                state.inputData = result.values;
            } else {
                state.outputData = result.values;
            }

            renderPreview(previewEl, result.values);
            infoEl.textContent = `${result.rows} rows × ${result.cols} columns`;
            hideStatus(statusEl);
        } catch (e) {
            showStatus(statusEl, 'error', e.message || 'Failed to read data.');
        }
    }

    async function runTransform() {
        const statusEl = $('#transform-status');
        const resultsEl = $('#results-section');
        resultsEl.classList.remove('visible');

        // Validate
        if (!state.inputData) {
            showStatus(statusEl, 'error', 'No input data captured. Use the buttons above to select data.');
            return;
        }
        if (!state.outputData) {
            showStatus(statusEl, 'error', 'No output example captured. Use the buttons above to select example data.');
            return;
        }

        const rules = $('#rules-textarea').value.trim();
        const previousScript = $('#prev-script-textarea').value.trim();
        const scriptLangValue = $('#script-language').value;
        const scriptLang = scriptLangValue === 'VBA' ? 'VBA' : 'Office Scripts (TypeScript)';

        const prompt = buildPrompt(state.inputData, state.outputData, rules, previousScript, scriptLang);

        showStatus(statusEl, 'loading', 'Analyzing data and generating transformation...');
        $('#btn-transform').disabled = true;

        try {
            const response = await apiCall('/api/transform', { prompt });

            showStatus(statusEl, 'loading', 'Parsing response...');

            let parsed;
            try {
                parsed = parseAIResponse(response.content);
            } catch (parseErr) {
                // Show raw response on parse failure
                showStatus(statusEl, 'error', 'Failed to parse AI response. Raw response shown below.');
                $('#result-script-code').textContent = response.content;
                resultsEl.classList.add('visible');
                $('#result-preview-container').innerHTML = '';
                $('#result-explanation').textContent = '';
                return;
            }

            state.result = parsed;

            // Show results
            renderPreview($('#result-preview-container'), parsed.transformedData);
            $('#result-preview-info').textContent = parsed.transformedData
                ? `${parsed.transformedData.length} rows × ${(parsed.transformedData[0] || []).length} columns`
                : '';
            $('#result-explanation').textContent = parsed.explanation || '';
            $('#result-script-code').textContent = parsed.script || '';
            resultsEl.classList.add('visible');
            showStatus(statusEl, 'success', 'Transformation complete!');
        } catch (e) {
            showStatus(statusEl, 'error', e.message || 'Transform failed.');
        } finally {
            $('#btn-transform').disabled = false;
        }
    }

    async function writeResults(mode) {
        if (!state.result || !state.result.transformedData) return;

        const statusEl = $('#transform-status');
        showStatus(statusEl, 'loading', 'Writing results to Excel...');

        try {
            if (mode === 'newSheet') {
                const name = await writeToNewSheet(state.result.transformedData);
                showStatus(statusEl, 'success', `Data written to new sheet: ${name}`);
            } else {
                await writeToSelection(state.result.transformedData);
                showStatus(statusEl, 'success', 'Data written to current selection.');
            }
        } catch (e) {
            showStatus(statusEl, 'error', e.message || 'Failed to write data.');
        }
    }

    function copyScript() {
        const code = $('#result-script-code').textContent;
        if (!code) return;
        navigator.clipboard.writeText(code).then(() => {
            const btn = $('#btn-copy-script');
            const orig = btn.textContent;
            btn.textContent = 'Copied!';
            setTimeout(() => { btn.textContent = orig; }, 1500);
        }).catch(() => {
            // Fallback
            const textarea = document.createElement('textarea');
            textarea.value = code;
            document.body.appendChild(textarea);
            textarea.select();
            document.execCommand('copy');
            document.body.removeChild(textarea);
        });
    }

})();
