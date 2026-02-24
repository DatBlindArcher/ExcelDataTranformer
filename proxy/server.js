const express = require('express');
const cors = require('cors');
const rateLimit = require('express-rate-limit');

const app = express();
const PORT = process.env.PORT || 3100;

// AI configuration from environment
const AI_PROVIDER = process.env.AI_PROVIDER || 'claude';
const AI_API_KEY  = process.env.AI_API_KEY  || '';
const AI_MODEL    = process.env.AI_MODEL    || '';

// Parse allowed origins from environment variable or use defaults
const allowedOrigins = process.env.ALLOWED_ORIGINS
    ? process.env.ALLOWED_ORIGINS.split(',').map(o => o.trim())
    : ['http://localhost:3000', 'https://localhost:3000'];

// CORS configuration — browsers will reject responses to non-allowed origins
app.use(cors({
    origin: allowedOrigins,
    methods: ['POST', 'GET', 'OPTIONS'],
    allowedHeaders: ['Content-Type']
}));

// JSON body parsing
app.use(express.json({ limit: '5mb' }));

// Origin enforcement — blocks non-browser requests to POST endpoints.
// CORS only instructs the browser; this middleware rejects requests server-side
// when the Origin header is missing or doesn't match the allow-list.
app.use('/api/', (req, res, next) => {
    // Allow GET (health check) and preflight OPTIONS through
    if (req.method === 'GET' || req.method === 'OPTIONS') return next();

    const origin = req.headers['origin'];
    if (!origin || !allowedOrigins.includes(origin)) {
        return res.status(403).json({
            success: false,
            error: 'Forbidden: request origin not allowed.',
            status: 403
        });
    }
    next();
});

// Rate limiting: 30 requests per minute per IP
const limiter = rateLimit({
    windowMs: 60 * 1000,
    max: 30,
    standardHeaders: true,
    legacyHeaders: false,
    message: { success: false, error: 'Rate limited. Please wait a moment and try again.', status: 429 }
});
app.use('/api/', limiter);

// Request logging (no API keys or prompt content)
app.use('/api/', (req, res, next) => {
    const timestamp = new Date().toISOString();
    const ip = req.headers['x-real-ip'] || req.headers['x-forwarded-for'] || req.ip;
    console.log(`[${timestamp}] ${req.method} ${req.path} | IP: ${ip} | Provider: ${AI_PROVIDER} | Model: ${AI_MODEL || '(default)'}`);
    next();
});

// Health check
app.get('/api/health', (req, res) => {
    res.json({ status: 'ok' });
});

// Test connection — sends a minimal request to the AI API using server-side config
app.post('/api/test', async (req, res) => {
    if (!AI_API_KEY) {
        return res.status(500).json({ success: false, error: 'AI_API_KEY is not configured on the server.' });
    }

    try {
        const testPrompt = 'Respond with exactly: ok';
        const result = await callAI(AI_PROVIDER, AI_API_KEY, AI_MODEL, testPrompt);
        res.json({ success: true, content: result });
    } catch (err) {
        const status = err.status || 500;
        console.error(`[${new Date().toISOString()}] TEST FAILED | Provider: ${AI_PROVIDER} | Model: ${AI_MODEL || 'default'} | Status: ${status} | ${err.message}`);
        res.status(status).json({ success: false, error: err.message, status });
    }
});

// Transform — forwards the full prompt to the AI API using server-side config
app.post('/api/transform', async (req, res) => {
    const { prompt } = req.body;

    if (!prompt) {
        return res.status(400).json({ success: false, error: 'Missing required field: prompt' });
    }

    if (!AI_API_KEY) {
        return res.status(500).json({ success: false, error: 'AI_API_KEY is not configured on the server.' });
    }

    try {
        const content = await callAI(AI_PROVIDER, AI_API_KEY, AI_MODEL, prompt);
        res.json({ success: true, content });
    } catch (err) {
        const status = err.status || 500;
        console.error(`[${new Date().toISOString()}] TRANSFORM FAILED | Provider: ${AI_PROVIDER} | Model: ${AI_MODEL || 'default'} | Status: ${status} | ${err.message}`);
        res.status(status).json({ success: false, error: err.message, status });
    }
});

// Call AI provider API
async function callAI(provider, apiKey, model, prompt) {
    let url, headers, body;

    if (provider === 'claude') {
        url = 'https://api.anthropic.com/v1/messages';
        headers = {
            'x-api-key': apiKey,
            'anthropic-version': '2023-06-01',
            'content-type': 'application/json'
        };
        body = JSON.stringify({
            model: model || 'claude-sonnet-4-20250514',
            max_tokens: 16384,
            messages: [
                { role: 'user', content: prompt },
                { role: 'assistant', content: '{' }   // Prefill forces JSON output
            ]
        });
    } else if (provider === 'openai') {
        url = 'https://api.openai.com/v1/chat/completions';
        headers = {
            'Authorization': `Bearer ${apiKey}`,
            'content-type': 'application/json'
        };
        body = JSON.stringify({
            model: model || 'gpt-4o',
            max_completion_tokens: 16384,
            response_format: { type: 'json_object' }, // Enforce JSON output
            messages: [{ role: 'user', content: prompt }]
        });
    } else {
        const err = new Error(`Unsupported provider: ${provider}`);
        err.status = 400;
        throw err;
    }

    const response = await fetch(url, { method: 'POST', headers, body });

    if (!response.ok) {
        const errorBody = await response.text();
        let errorMessage;
        try {
            const parsed = JSON.parse(errorBody);
            errorMessage = parsed.error?.message || parsed.error || errorBody;
        } catch {
            errorMessage = errorBody;
        }
        const err = new Error(errorMessage);
        err.status = response.status;
        throw err;
    }

    const data = await response.json();

    // Normalize response: extract the text content
    if (provider === 'claude') {
        // Prepend '{' to match the assistant prefill we used to force JSON
        return '{' + data.content[0].text;
    } else {
        return data.choices[0].message.content;
    }
}

app.listen(PORT, () => {
    console.log(`Excel AI Proxy running on port ${PORT}`);
    console.log(`Allowed origins: ${allowedOrigins.join(', ')}`);
    console.log(`AI Provider: ${AI_PROVIDER} | Model: ${AI_MODEL || '(default)'}`);
    if (!AI_API_KEY) {
        console.warn('WARNING: AI_API_KEY is not set. API calls will fail.');
    }
});
