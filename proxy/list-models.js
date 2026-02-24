#!/usr/bin/env node

// List available models for the configured AI provider.
// Usage: node list-models.js
// Reads AI_PROVIDER and AI_API_KEY from ../.env or environment variables.

const fs = require('fs');
const path = require('path');

// Load .env file from project root
const envPath = path.join(__dirname, '..', '.env');
if (fs.existsSync(envPath)) {
    for (const line of fs.readFileSync(envPath, 'utf8').split('\n')) {
        const trimmed = line.trim();
        if (!trimmed || trimmed.startsWith('#')) continue;
        const eq = trimmed.indexOf('=');
        if (eq === -1) continue;
        const key = trimmed.slice(0, eq).trim();
        const val = trimmed.slice(eq + 1).trim();
        if (!process.env[key]) process.env[key] = val;
    }
}

const provider = process.env.AI_PROVIDER || 'claude';
const apiKey = process.env.AI_API_KEY || '';

if (!apiKey) {
    console.error('Error: AI_API_KEY is not set. Add it to ../.env or set it as an environment variable.');
    process.exit(1);
}

async function listModels() {
    let url, headers;

    if (provider === 'claude') {
        url = 'https://api.anthropic.com/v1/models?limit=100';
        headers = {
            'x-api-key': apiKey,
            'anthropic-version': '2023-06-01'
        };
    } else if (provider === 'openai') {
        url = 'https://api.openai.com/v1/models';
        headers = {
            'Authorization': `Bearer ${apiKey}`
        };
    } else {
        console.error(`Unsupported provider: ${provider}`);
        process.exit(1);
    }

    console.log(`Provider: ${provider}`);
    console.log(`Fetching models...\n`);

    const response = await fetch(url, { headers });

    if (!response.ok) {
        const body = await response.text();
        console.error(`API error (${response.status}): ${body}`);
        process.exit(1);
    }

    const data = await response.json();

    let models;
    if (provider === 'claude') {
        models = data.data.map(m => ({ id: m.id, name: m.display_name || m.id, created: m.created_at }));
        models.sort((a, b) => b.created.localeCompare(a.created));
    } else {
        models = data.data.map(m => ({ id: m.id, created: new Date(m.created * 1000).toISOString().slice(0, 10) }));
        models.sort((a, b) => b.created.localeCompare(a.created));
    }

    console.log(`Found ${models.length} models:\n`);

    for (const m of models) {
        if (provider === 'claude') {
            console.log(`  ${m.id}  (${m.name})`);
        } else {
            console.log(`  ${m.id}  (${m.created})`);
        }
    }
}

listModels().catch(err => {
    console.error('Failed:', err.message);
    process.exit(1);
});
