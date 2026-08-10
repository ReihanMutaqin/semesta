const express = require('express');
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const app = express();
const PORT = process.env.PORT || 3000;

const DB_PATH = path.join(__dirname, 'semesta.db');
const SUMMARY_PATH = path.join(__dirname, 'summary.json');
const EXCEL_PATH = path.join(__dirname, 'export.xlsx');
const CONVERT_SCRIPT = path.join(__dirname, 'convert_to_sqlite.py');
const QUERY_SCRIPT = path.join(__dirname, 'query_orders.py');
const STATIC_DIR = path.join(__dirname, 'public');

// Body parser for up to 200MB file upload
app.use(express.raw({ limit: '200mb', type: '*/*' }));
app.use(express.static(STATIC_DIR));

// Serve index.html explicitly at root URL
app.get('/', (req, res) => {
    res.sendFile(path.join(STATIC_DIR, 'index.html'));
});

// Endpoint 1: Get Summary
app.get('/api/summary', (req, res) => {
    try {
        if (fs.existsSync(SUMMARY_PATH)) {
            const data = fs.readFileSync(SUMMARY_PATH, 'utf-8');
            res.setHeader('Content-Type', 'application/json');
            res.send(data);
        } else {
            res.status(404).json({ error: 'Summary file not found' });
        }
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// Endpoint 2: Upload new Excel
app.post('/api/upload', (req, res) => {
    try {
        if (!req.body || req.body.length === 0) {
            return res.status(400).json({ error: 'File payload is empty' });
        }

        console.log(`Received uploaded file: ${(req.body.length / (1024 * 1024)).toFixed(2)} MB. Saving to ${EXCEL_PATH}...`);
        fs.writeFileSync(EXCEL_PATH, req.body);

        console.log('Running convert_to_sqlite.py to rebuild database and summary stats...');
        execSync(`python "${CONVERT_SCRIPT}"`, { stdio: 'inherit' });

        console.log('Re-indexing completed!');

        if (fs.existsSync(SUMMARY_PATH)) {
            const updatedSummary = JSON.parse(fs.readFileSync(SUMMARY_PATH, 'utf-8'));
            return res.json({
                success: true,
                message: 'File excel berhasil diupload dan diproses!',
                summary: updatedSummary
            });
        }

        res.json({ success: true, message: 'File excel berhasil diproses!' });

    } catch (err) {
        console.error('Upload process error:', err);
        res.status(500).json({ error: `Gagal memproses file upload: ${err.message}` });
    }
});

// Endpoint 3: Query Orders from SQLite via query_orders.py
app.get('/api/orders', (req, res) => {
    try {
        const queryParams = new URLSearchParams(req.query).toString();
        const cmd = `python "${QUERY_SCRIPT}" "${queryParams.replace(/"/g, '\\"')}"`;
        const stdout = execSync(cmd, { encoding: 'utf-8', maxBuffer: 20 * 1024 * 1024 });
        res.setHeader('Content-Type', 'application/json');
        res.send(stdout);
    } catch (err) {
        console.error('API Orders error:', err);
        res.status(500).json({ error: err.message });
    }
});

app.listen(PORT, () => {
    console.log(`Node Express Server running at http://localhost:${PORT}`);
});

module.exports = app;
