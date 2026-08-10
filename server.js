const express = require('express');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

const SUMMARY_PATH = path.join(__dirname, 'summary.json');
const STATIC_DIR = path.join(__dirname, 'public');

// Body parser
app.use(express.raw({ limit: '200mb', type: '*/*' }));
app.use(express.static(STATIC_DIR));

// Root route
app.get('/', (req, res) => {
    res.sendFile(path.join(STATIC_DIR, 'index.html'));
});

// Endpoint 1: Get Summary
app.get('/api/summary', (req, res) => {
    try {
        if (fs.existsSync(SUMMARY_PATH)) {
            const data = fs.readFileSync(SUMMARY_PATH, 'utf-8');
            res.setHeader('Content-Type', 'application/json');
            return res.send(data);
        }
        res.json({
            max_date: 'Belum Ada Data',
            total_order_semesta: 0,
            total_ps: 0,
            ps_percentage: 0,
            total_ps_last_month: 0,
            type_summary: [],
            segment_summary: [],
            daily_trend: []
        });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// Endpoint 2: Get Orders
app.get('/api/orders', (req, res) => {
    res.json({
        page: 1,
        limit: 20,
        total_records: 0,
        total_pages: 1,
        orders: []
    });
});

app.listen(PORT, () => {
    console.log(`Node Express Server running at http://localhost:${PORT}`);
});

module.exports = app;
