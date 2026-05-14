# IRA Portfolio Rebalancing Dashboard

A client-side portfolio rebalancing tool for multiple IRA accounts. Tracks holdings, calculates drift from target allocations, fetches live prices, generates rebalance suggestions, and performs tax-aware analysis — all in the browser with no backend required for the hosted version.

**Live site:** https://aladinz.github.io/ira_portfolio_rebalance_dashboard/

---

## Features

| Feature | Description |
|---|---|
| **Multi-portfolio support** | Manage any number of IRA accounts (Traditional, Roth, Rollover, etc.) on a single dashboard |
| **Drift analysis** | Real-time current % vs target % with color-coded drift alerts |
| **Suggested trades** | Auto-calculated buy/sell quantities to return each holding to target weight |
| **Live price fetching** | Fetches current market prices via Finnhub API (free key required) |
| **Fidelity fund proxies** | FZROX, FZILX, FXAIX etc. are automatically mapped to equivalent ETFs for price lookup |
| **Rebalance Suggestion** | Generates a formatted text report per portfolio, copyable to clipboard |
| **Tax-Aware Analysis** | Stage 8 tax layer scores each trade by tax impact (gain/loss, holding period) |
| **Metrics dashboard** | Dedicated analytics page with allocation visuals, concentration diagnostics, and simulation toolkit |
| **Export CSV** | Downloads all holdings for a portfolio as a spreadsheet |
| **Cloud sync** | JSONBin.io integration — persists data across browsers and devices |
| **localStorage fallback** | All data is cached locally so the page survives a refresh without the cloud key |

---

## Interface Enhancements

Recent UI updates introduced a cohesive visual language across both `index.html` and `dashboard.html`:

- Refined typography with clearer hierarchy for dense financial data
- Atmospheric gradient backdrop and elevated card surfaces for readability
- Consistent control styling (buttons, links, pills, badges) across pages
- Polished table states, focus rings, and hover feedback for fast scanning
- Subtle entrance motion with `prefers-reduced-motion` accessibility fallback

Use the **Metrics** button in the main header to open the advanced analytics page.

---

## Data Persistence

The dashboard has three layers of persistence, used in priority order:

```
1. JSONBin.io (cloud)  ← primary, works across browsers/devices
2. localStorage        ← automatic local cache, survives refresh
3. Local server        ← used only when running locally with Node.js
```

### How data is saved

Every edit (shares, price, target %, ticker, portfolio name, etc.) automatically triggers a debounced save within 500 ms. The **✓ Saved** indicator in the header confirms the write completed.

---

## Cloud Sync Setup (JSONBin.io)

This is the recommended one-time setup to make your data permanent and portable.

### Step 1 — Create a free JSONBin account

Go to https://jsonbin.io and sign up for a free account.

### Step 2 — Get your X-Master-Key

After logging in:
1. Click your profile → **API Keys**
2. Copy the **X-Master-Key** (starts with `$2b$10$…`)

### Step 3 — Enter the key in the dashboard

1. Open the dashboard at https://aladinz.github.io/ira_portfolio_rebalance_dashboard/
2. Click **⚙ Settings** in the top-right corner
3. Under **Cloud Save — JSONBin.io**, paste your X-Master-Key
4. Click **Save Key**

A private storage bin is **created automatically** on first save. The Bin ID appears below the field (e.g. `69e552f036566621a8ce7f7e`). Keep note of it — it's informational only, you only ever need the key to reconnect.

### Step 4 — Using the dashboard on a different browser or device

The X-Master-Key is stored in each browser's localStorage. On any new browser:

1. Open https://aladinz.github.io/ira_portfolio_rebalance_dashboard/
2. The page loads with demo data (the key isn't there yet)
3. Click **⚙ Settings** → paste your X-Master-Key → click **Save Key**
4. Your portfolios are pulled from JSONBin and restored immediately

> **Think of the X-Master-Key as your password.** Keep it somewhere safe (a password manager). Anyone with this key can read and overwrite your bin.

---

## Live Prices — Finnhub API

### Step 1 — Get a free Finnhub key

Sign up at https://finnhub.io. Your free API key is shown on the dashboard immediately after login.

### Step 2 — Enter the key

1. Click **⚙ Settings**
2. Under **Live Prices — Finnhub API Key**, paste the key
3. Click **Save Key**

### Step 3 — Fetch prices

Click **Fetch Prices** on any portfolio card. All tickers are fetched in parallel. Cells with proxied prices (Fidelity funds) are tinted amber with a tooltip showing the proxy ETF used.

---

## Running Locally (Node.js)

Use this mode if you want data saved to a local file instead of the cloud.

### Prerequisites

- Node.js 18 or later
- npm

### Install and start

```bash
git clone https://github.com/aladinz/ira_portfolio_rebalance_dashboard.git
cd ira_portfolio_rebalance_dashboard
npm install
npm start
```

Open http://localhost:3000 in your browser.

Portfolio data is saved to `data/portfolio.json` automatically. To back up your data, copy that file. To restore it, replace the file and restart the server.

For auto-restart on file changes during development:

```bash
npm run dev
```

---

## Project Structure

```
index.html              — Main rebalancing dashboard UI and modals
dashboard.html          — Analytics and metrics dashboard (charts + toolkit)
styles.css              — Shared styling system
script.js               — Main dashboard logic and persistence
toolkit.js              — Shared investor toolkit components
stage8_tax_layer.js     — Tax-aware rebalancing analysis layer
server.js               — Local Node.js/Express server
package.json            — npm config
data/
  dataStore.js          — File-based persistence (local server only)
  portfolio.json        — Saved portfolio data (local server only)
```

---

## Resetting to Demo Data

Click **↺ Reset Demo** in the header to wipe all portfolios and restore the original demo data. This also overwrites your JSONBin bin and localStorage, so use with caution.

---

## Browser Compatibility

Any modern browser (Chrome, Edge, Firefox, Safari). JavaScript must be enabled. No frameworks or external libraries are used — the entire front-end is vanilla HTML, CSS, and JavaScript.
