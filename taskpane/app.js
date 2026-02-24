/* ============================================================
   Excel AI Transform — Task Pane Application Logic
   ============================================================ */

(function () {
    'use strict';

    // ── State ──────────────────────────────────────────────────
    const state = {
        inputData: null,       // 2D array
        outputData: null,      // 2D array
        result: null,          // { transformedData, jsTransform, script, explanation }
        lastJsTransform: null, // last generated JS function source (for retry)
        lastExecError: null,   // last local execution error message (for retry)
        abortController: null  // AbortController for in-flight API request
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
    const SAMPLE_SIZE = 50;

    function buildTransformPrompt(inputData, outputExample, rules, previousScript, scriptLanguage) {
        const totalRows = inputData.length - 1; // exclude header
        const sampleRows = inputData.slice(0, Math.min(inputData.length, SAMPLE_SIZE + 1));
        const sampleCsv = arrayToCsv(sampleRows);
        const showingCount = Math.min(totalRows, SAMPLE_SIZE);
        const outputCsv = arrayToCsv(outputExample);

        return `You are an Excel data transformation assistant. You will be given:
1. INPUT DATA: A sample of the raw source data (header + up to ${SAMPLE_SIZE} rows, as CSV).
2. OUTPUT EXAMPLE: An example of what the transformed data should look like (as CSV).
3. (Optional) TRANSFORMATION RULES: Additional rules or descriptions.
4. (Optional) PREVIOUS SCRIPT: A previously used script for reference.

Your task:
A) Analyze the input sample and output example to infer the transformation logic.
B) Write a JavaScript function that implements this transformation.
C) Generate a reusable script in ${scriptLanguage} for the user.

The JavaScript function MUST follow this exact signature and contract:
- Signature: function transform(header, rows)
- header: a 1D array of strings (the first row / column names from the input)
- rows: a 2D array of the remaining data rows (each row is an array of values)
- Return value: a 2D array INCLUDING the new header as the first row
- The function must be pure (no external dependencies, no DOM access, no fetch)
- The function must handle edge cases: empty cells (null or ""), missing columns
- Use only standard JavaScript (ES2017) — no import/require, no Node.js APIs
- Cell values can be strings, numbers, booleans, or null (empty cells). Handle all types appropriately.

Respond with EXACTLY this JSON structure (no markdown fences, no extra text):
{
  "jsTransform": "function transform(header, rows) { ... }",
  "script": "...the full ${scriptLanguage} code as a string...",
  "explanation": "Brief explanation of the transformation logic."
}

--- INPUT DATA (CSV, ${totalRows} total data rows, showing first ${showingCount}) ---
${sampleCsv}

--- OUTPUT EXAMPLE (CSV) ---
${outputCsv}

--- TRANSFORMATION RULES ---
${rules || '(none provided)'}

--- PREVIOUS SCRIPT ---
${previousScript || '(none provided)'}`;
    }

    function buildFixPrompt(inputData, outputExample, failedFunction, errorMessage, scriptLanguage) {
        const sampleCsv = arrayToCsv(inputData.slice(0, Math.min(inputData.length, SAMPLE_SIZE + 1)));
        const outputCsv = arrayToCsv(outputExample);

        return `You previously generated a JavaScript transform function that failed with an error when executed locally.

--- ORIGINAL INPUT DATA (CSV, sample) ---
${sampleCsv}

--- EXPECTED OUTPUT (CSV) ---
${outputCsv}

--- FAILED FUNCTION ---
${failedFunction}

--- ERROR MESSAGE ---
${errorMessage}

Please fix the function. Same requirements as before:
- Signature: function transform(header, rows)
- header: a 1D array of strings (the first row / column names)
- rows: a 2D array of the remaining data rows
- Return a 2D array INCLUDING the new header as the first row
- Pure JavaScript (ES2017), no external dependencies
- Cell values can be strings, numbers, booleans, or null

Respond with EXACTLY this JSON (no markdown fences, no extra text):
{
  "jsTransform": "function transform(header, rows) { ... }",
  "script": "...the full ${scriptLanguage} code...",
  "explanation": "Explanation of what was fixed."
}`;
    }

    // ── AI response parsing ───────────────────────────────────
    function parseTransformResponse(raw) {
        // Strip markdown code fences if present
        let cleaned = raw.trim();
        cleaned = cleaned.replace(/^```(?:json)?\s*\n?/i, '');
        cleaned = cleaned.replace(/\n?```\s*$/i, '');
        cleaned = cleaned.trim();

        const parsed = JSON.parse(cleaned);

        // New format: jsTransform function
        if (typeof parsed.jsTransform === 'string' && typeof parsed.script === 'string') {
            return parsed;
        }

        // Backward-compatible: AI returned old format with transformedData
        if (Array.isArray(parsed.transformedData) && typeof parsed.script === 'string') {
            return {
                jsTransform: null,
                transformedData: parsed.transformedData,
                script: parsed.script,
                explanation: parsed.explanation || ''
            };
        }

        throw new Error('Invalid response structure: missing jsTransform function or script string.');
    }

    // ── Local JS execution engine ────────────────────────────
    async function executeJsTransform(jsTransformSource, fullData) {
        // Shallow copy to prevent AI-generated code from mutating state
        var header = fullData[0].slice();
        var rows = fullData.slice(1).map(function(row) { return row.slice(); });

        // Create the transform function from AI-generated source
        var factory = new Function('"use strict"; return (' + jsTransformSource + ')');
        var transformFn = factory();

        if (typeof transformFn !== 'function') {
            throw new Error('AI did not return a valid function.');
        }

        // Yield to event loop so UI can update before heavy computation
        await new Promise(function(resolve) { setTimeout(resolve, 0); });

        var result = transformFn(header, rows);

        // Validate result shape
        if (!Array.isArray(result) || result.length === 0) {
            throw new Error('Transform function returned empty or non-array result.');
        }
        if (!Array.isArray(result[0])) {
            throw new Error('Transform function did not return a 2D array.');
        }

        return result;
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
    var CLIENT_TIMEOUT_MS = 90000; // 90s — under nginx's 120s for clean errors

    async function apiCall(endpoint, body, externalSignal) {
        var url = PROXY_URL + endpoint;

        // Set up timeout + optional external abort
        var controller = new AbortController();
        var wasTimeout = false;
        var timeoutId = setTimeout(function() {
            wasTimeout = true;
            controller.abort();
        }, CLIENT_TIMEOUT_MS);

        // If caller provided an external signal, forward its abort
        if (externalSignal) {
            if (externalSignal.aborted) {
                clearTimeout(timeoutId);
                throw new Error('Request was cancelled.');
            }
            externalSignal.addEventListener('abort', function() {
                clearTimeout(timeoutId);
                controller.abort();
            });
        }

        try {
            var response = await fetch(url, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(body),
                signal: controller.signal
            });

            clearTimeout(timeoutId);

            var data = await safeJsonParse(response);
            if (!data.success) {
                var status = data.status || response.status;
                if (status === 429) throw new Error('Rate limited. Please wait a moment and try again.');
                if (status >= 500) throw new Error('AI service error. Try again later.');
                throw new Error(data.error || 'Unknown error from proxy.');
            }
            return data;
        } catch (e) {
            clearTimeout(timeoutId);
            if (e.name === 'AbortError') {
                if (wasTimeout) {
                    throw new Error('Request timed out. The AI took too long to respond.');
                }
                throw new Error('Request was cancelled.');
            }
            throw e;
        }
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
        $('#btn-transform').addEventListener('click', function() { runTransform(false); });

        // Cancel in-flight transform
        $('#btn-cancel-transform').addEventListener('click', function() {
            if (state.abortController) {
                state.abortController.abort();
                state.abortController = null;
            }
        });

        // Retry with AI fix after local execution failure
        $('#btn-retry-transform').addEventListener('click', function() { runTransform(true); });

        // JS transform section collapsible
        $('#js-transform-header').addEventListener('click', function() {
            var chevron = $('#js-transform-header .chevron');
            var body = $('#js-transform-body');
            chevron.classList.toggle('open');
            body.classList.toggle('open');
        });

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

    async function runTransform(retryMode) {
        var statusEl = $('#transform-status');
        var resultsEl = $('#results-section');
        resultsEl.classList.remove('visible');
        $('#btn-retry-transform').style.display = 'none';

        // Validate
        if (!state.inputData) {
            showStatus(statusEl, 'error', 'No input data captured. Use the buttons above to select data.');
            return;
        }
        if (!state.outputData) {
            showStatus(statusEl, 'error', 'No output example captured. Use the buttons above to select example data.');
            return;
        }

        var rules = $('#rules-textarea').value.trim();
        var previousScript = $('#prev-script-textarea').value.trim();
        var scriptLangValue = $('#script-language').value;
        var scriptLang = scriptLangValue === 'VBA' ? 'VBA' : 'Office Scripts (TypeScript)';

        // Build prompt: new transform or retry fix
        var prompt;
        if (retryMode && state.lastJsTransform && state.lastExecError) {
            prompt = buildFixPrompt(
                state.inputData, state.outputData,
                state.lastJsTransform, state.lastExecError, scriptLang
            );
        } else {
            prompt = buildTransformPrompt(
                state.inputData, state.outputData, rules, previousScript, scriptLang
            );
        }

        // ── Phase 1: AI call ────────────────────────────────
        showStatus(statusEl, 'loading', retryMode
            ? 'Asking AI to fix the transform function...'
            : 'Analyzing pattern from sample data...');
        $('#btn-transform').disabled = true;
        $('#btn-cancel-transform').style.display = '';

        // Cancel any previous in-flight request
        if (state.abortController) state.abortController.abort();
        state.abortController = new AbortController();

        try {
            var response = await apiCall('/api/transform', { prompt }, state.abortController.signal);

            showStatus(statusEl, 'loading', 'Parsing AI response...');

            var parsed;
            try {
                parsed = parseTransformResponse(response.content);
            } catch (parseErr) {
                // Show raw response on parse failure
                showStatus(statusEl, 'error', 'Failed to parse AI response. Raw response shown below.');
                $('#result-script-code').textContent = response.content;
                resultsEl.classList.add('visible');
                $('#result-preview-container').innerHTML = '';
                $('#result-explanation').textContent = '';
                return;
            }

            // ── Phase 2: Local execution ────────────────────
            var transformedData;

            if (parsed.jsTransform) {
                // New format: execute JS locally
                var rowCount = state.inputData.length - 1;
                showStatus(statusEl, 'loading',
                    'Applying transformation to ' + rowCount + ' rows...');

                try {
                    transformedData = await executeJsTransform(parsed.jsTransform, state.inputData);
                } catch (execErr) {
                    // Store for retry
                    state.lastJsTransform = parsed.jsTransform;
                    state.lastExecError = execErr.message;

                    showStatus(statusEl, 'error',
                        'Local execution failed: ' + execErr.message);
                    $('#result-explanation').textContent = parsed.explanation || '';
                    $('#result-script-code').textContent = parsed.script || '';
                    $('#result-js-transform-code').textContent = parsed.jsTransform;
                    $('#js-transform-section').style.display = '';
                    $('#result-preview-container').innerHTML =
                        '<div class="empty-state">Transform function failed — see generated code below</div>';
                    $('#result-preview-info').textContent = '';
                    resultsEl.classList.add('visible');
                    $('#btn-retry-transform').style.display = '';
                    return;
                }
            } else if (parsed.transformedData) {
                // Backward-compatible: AI returned data directly
                transformedData = parsed.transformedData;
            } else {
                showStatus(statusEl, 'error', 'AI did not return a transform function or data.');
                return;
            }

            // ── Success ─────────────────────────────────────
            state.result = {
                transformedData: transformedData,
                jsTransform: parsed.jsTransform || null,
                script: parsed.script,
                explanation: parsed.explanation
            };
            state.lastJsTransform = parsed.jsTransform || null;
            state.lastExecError = null;

            renderPreview($('#result-preview-container'), transformedData);
            $('#result-preview-info').textContent = transformedData
                ? transformedData.length + ' rows \u00d7 ' + (transformedData[0] || []).length + ' columns'
                : '';
            $('#result-explanation').textContent = parsed.explanation || '';
            $('#result-script-code').textContent = parsed.script || '';
            if (parsed.jsTransform) {
                $('#result-js-transform-code').textContent = parsed.jsTransform;
                $('#js-transform-section').style.display = '';
            } else {
                $('#js-transform-section').style.display = 'none';
            }
            resultsEl.classList.add('visible');
            showStatus(statusEl, 'success', 'Transformation complete!');
        } catch (e) {
            showStatus(statusEl, 'error', e.message || 'Transform failed.');
        } finally {
            $('#btn-transform').disabled = false;
            $('#btn-cancel-transform').style.display = 'none';
            state.abortController = null;
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
