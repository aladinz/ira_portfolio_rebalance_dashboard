/* ================================================================
   IRA Portfolio Rebalancing Dashboard — Data Store
   ================================================================
   Node.js module (server-side only).
   Reads and writes portfolio data as JSON to ./data/portfolio.json.

   All writes are atomic: data is first written to a temp file
   and then renamed over the final file so a crash mid-write can
   never corrupt the live data.

   On first run:    portfolio.json is created with empty defaults.
   On corruption:   portfolio.json is backed up to portfolio.bak.json
                    and a clean defaults file is generated.
   ================================================================ */

'use strict';

const fs   = require('fs/promises');
const path = require('path');

/* ── File paths ─────────────────────────────────────────────── */
const DATA_FILE = path.join(__dirname, 'portfolio.json');
const TEMP_FILE = path.join(__dirname, 'portfolio.tmp.json');
const BAK_FILE  = path.join(__dirname, 'portfolio.bak.json');

/* ── Default structure written on first run ─────────────────── */
const DEFAULT_DATA = {
  portfolioSeq: 0,
  savedAt     : new Date().toISOString(),
  settings    : { finnhubKey: '' },
  portfolios  : [],
};

/* ================================================================
   VALIDATE
   ================================================================ */

/**
 * Confirm the object has the required shape.
 * Throws a descriptive Error if anything is missing or wrong type.
 * Silently normalises settings to a safe default if absent.
 *
 * @param   {unknown} data
 * @returns {object}  The same data object, mutated if needed.
 */
function validate(data) {
  if (!data || typeof data !== 'object' || Array.isArray(data)) {
    throw new Error('Portfolio data must be a JSON object.');
  }
  if (!Array.isArray(data.portfolios)) {
    throw new Error('Missing or invalid "portfolios" array.');
  }
  if (typeof data.portfolioSeq !== 'number') {
    throw new Error('Missing or invalid "portfolioSeq" number.');
  }
  /* Normalise missing settings block */
  if (!data.settings || typeof data.settings !== 'object' || Array.isArray(data.settings)) {
    data.settings = { finnhubKey: '' };
  }
  return data;
}

/* ================================================================
   LOAD
   ================================================================ */

/**
 * Read portfolio data from disk.
 *
 * Behaviour:
 *   - File does not exist  → create it with empty defaults (first run).
 *   - File is corrupted    → back it up to portfolio.bak.json and
 *                            regenerate a clean defaults file.
 *   - File is valid        → parse, validate, and return.
 *
 * @returns {Promise<object>}
 */
async function loadData() {
  try {
    const raw  = await fs.readFile(DATA_FILE, 'utf8');
    const data = JSON.parse(raw);
    return validate(data);
  } catch (err) {
    if (err.code === 'ENOENT') {
      /* First run — file does not exist; create it with defaults */
      const fresh = { ...DEFAULT_DATA, savedAt: new Date().toISOString() };
      await saveData(fresh);
      return { ...fresh };
    }

    /* File exists but could not be parsed or failed validation — back it up */
    console.warn('[dataStore] portfolio.json is corrupted; backing up to', BAK_FILE);
    console.warn('[dataStore] Error:', err.message);
    try { await fs.copyFile(DATA_FILE, BAK_FILE); } catch (_) { /* ignore */ }

    const clean = { ...DEFAULT_DATA, savedAt: new Date().toISOString() };
    await saveData(clean);
    return { ...clean };
  }
}

/* ================================================================
   SAVE  (atomic write)
   ================================================================ */

/**
 * Persist data to disk using an atomic two-step write:
 *   1. Write JSON to portfolio.tmp.json
 *   2. Rename temp → portfolio.json  (atomic on the same filesystem)
 *
 * This guarantees the live file is never in a half-written state.
 *
 * @param   {object} data
 * @returns {Promise<void>}
 */
async function saveData(data) {
  const safe   = validate({ ...data });
  safe.savedAt = new Date().toISOString();
  const json   = JSON.stringify(safe, null, 2);
  await fs.writeFile(TEMP_FILE, json, 'utf8');
  await fs.rename(TEMP_FILE, DATA_FILE);
}

/* ================================================================
   RESET
   ================================================================ */

/**
 * Wipe data to empty defaults.
 * Backs up the current file first so nothing is permanently lost.
 *
 * @returns {Promise<void>}
 */
async function resetData() {
  try { await fs.copyFile(DATA_FILE, BAK_FILE); } catch (_) { /* no file to back up */ }
  const clean = { ...DEFAULT_DATA, savedAt: new Date().toISOString() };
  await saveData(clean);
}

/* ── Exports ────────────────────────────────────────────────── */
module.exports = { loadData, saveData, resetData };
