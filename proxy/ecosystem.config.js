module.exports = {
    apps: [{
        name: 'excel-ai-proxy',
        script: 'server.js',
        instances: 1,
        env: {
            NODE_ENV: 'production',
            PORT: 3100,
            ALLOWED_ORIGINS: 'https://excel.archtech.be'
        }
    }]
};
