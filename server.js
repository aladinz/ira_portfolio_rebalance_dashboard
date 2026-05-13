/* ================================================================
   IRA Portfolio Rebalancing Dashboard — Local Server
   ================================================================
   Serves the static front-end and exposes three JSON endpoints so
   the browser can load and save portfolio data from a local file
   (./data/portfolio.json) instead of using localStorage or cloud sync.

   Start:  npm start          (or: node server.js)
   Dev:    npm run dev        (auto-restarts on file changes, Node ≥18)
   Open:   http://localhost:3000
   ================================================================ */

'use strict';

const express = require('express');
const path    = require('path');
const store   = require('./data/dataStore');

const app  = express();
const PORT = process.env.PORT || 3000;

/* ── Middleware ─────────────────────────────────────────────── */
app.use(express.json({ limit: '2mb' }));

/* Serve all static front-end files from the project root */
app.use(express.static(path.join(__dirname)));

/* ── GET /api/data ─────────────────────────────────────────── */
/* Load portfolio data from disk and return it as JSON.          */
app.get('/api/data', async (_req, res) => {
  try {
    const data = await store.loadData();
    res.json(data);
  } catch (err) {
    console.error('[GET /api/data]', err.message);
    res.status(500).json({ error: 'Failed to read portfolio data.' });
  }
});

/* ── POST /api/data ────────────────────────────────────────── */
/* Validate and atomically write portfolio data to disk.         */
app.post('/api/data', async (req, res) => {
  try {
    await store.saveData(req.body);
    res.json({ ok: true });
  } catch (err) {
    console.error('[POST /api/data]', err.message);
    res.status(400).json({ error: err.message });
  }
});

/* ── POST /api/data/reset ──────────────────────────────────── */
/* Wipe portfolio data to empty defaults (backs up first).       */
app.post('/api/data/reset', async (_req, res) => {
  try {
    await store.resetData();
    const data = await store.loadData();
    res.json(data);
  } catch (err) {
    console.error('[POST /api/data/reset]', err.message);
    res.status(500).json({ error: 'Failed to reset portfolio data.' });
  }
});

/* ── Start ──────────────────────────────────────────────────── */
app.listen(PORT, () => {
  console.log(`\n  IRA Portfolio Dashboard`);
  console.log(`  → http://localhost:${PORT}\n`);
});
