/* ================================================================
   IRA Portfolio Rebalancing Dashboard — Script
   Vanilla JS · No frameworks · No dependencies
   ================================================================ */

'use strict';

/* ── Portfolio Data ─────────────────────────────────────────── */
/*
  Demo portfolios — prices reflect a plausible 2026 snapshot.
  Targets are intentionally skewed from current weights so the
  initial render shows interesting drift values and alert rows.
*/
/* ── Default portfolio data factory ────────────────────────── */
/*
  Wrapped in a function so we always get a fresh deep copy when
  resetting to demo data — avoids accidental mutation.
*/
function getDefaultPortfolios() {
  return [
    {
      id: 'trad-ira',
      name: 'Traditional IRA',
      subtitle: 'Aggressive Growth — Tax-Deferred',
      holdings: [
        { ticker: 'VTI',  shares: 150, price: 238.42, targetPct: 45, mktPrice: 0 },
        { ticker: 'QQQ',  shares:  45, price: 495.80, targetPct: 35, mktPrice: 0 },
        { ticker: 'VXUS', shares: 100, price:  62.15, targetPct: 10, mktPrice: 0 },
        { ticker: 'BND',  shares:  50, price:  74.22, targetPct:  5, mktPrice: 0 },
        { ticker: 'GLD',  shares:  25, price: 189.50, targetPct:  5, mktPrice: 0 },
      ],
    },
    {
      id: 'roth-ira',
      name: 'Roth IRA',
      subtitle: 'Income & Stability — Tax-Free Growth',
      holdings: [
        { ticker: 'SCHD', shares: 200, price:  79.88, targetPct: 20, mktPrice: 0 },
        { ticker: 'VYM',  shares: 120, price: 124.33, targetPct: 29, mktPrice: 0 },
        { ticker: 'VGIT', shares: 200, price:  58.90, targetPct: 19, mktPrice: 0 },
        { ticker: 'VTIP', shares: 100, price: 106.75, targetPct: 21, mktPrice: 0 },
        { ticker: 'VNQ',  shares:  80, price:  88.50, targetPct: 11, mktPrice: 0 },
      ],
    },
  ];
}

let PORTFOLIOS = getDefaultPortfolios();


/* ── Demo template used when adding a new portfolio ────────── */
const DEMO_PORTFOLIO_TEMPLATE = [
  { ticker: 'VTI',  shares: 100, price: 238.42, targetPct: 40, mktPrice: 0 },
  { ticker: 'VXUS', shares:  80, price:  62.15, targetPct: 20, mktPrice: 0 },
  { ticker: 'BND',  shares:  60, price:  74.22, targetPct: 20, mktPrice: 0 },
  { ticker: 'VNQ',  shares:  40, price:  88.50, targetPct: 10, mktPrice: 0 },
  { ticker: 'GLD',  shares:  20, price: 189.50, targetPct: 10, mktPrice: 0 },
];

/* ── Counters ───────────────────────────────────────────────── */
let _rowId = 0;
const nextRowId = () => `row-${++_rowId}`;

/* Tracks how many portfolios have ever been created (for unique IDs/names) */
let _portfolioSeq = PORTFOLIOS.length;

/**
 * SAFETY GUARD — Gist auto-push is disabled until real user data has been
 * confirmed loaded (from localStorage or Gist). This prevents demo data
 * from ever being auto-pushed when the app opens fresh on a new browser/origin.
 * Only set to true by loadState() returning true, a successful Gist pull,
 * or an explicit user action (manual Push button).
 */
let _gistPushEnabled = false;

/* ── localStorage persistence ───────────────────────────────── */
const STORAGE_KEY = 'ira-dashboard-v1';

/* ── GitHub Gist cloud sync ──────────────────────────── */
const GIST_PAT_KEY  = 'ira-gist-pat';
const GIST_ID_KEY   = 'ira-gist-id';
const GIST_FILENAME = 'ira-dashboard.json';
const GH_API        = 'https://api.github.com';

const gistSync = {
  get pat()       { return localStorage.getItem(GIST_PAT_KEY) || ''; },
  get gistId()    { return localStorage.getItem(GIST_ID_KEY)  || ''; },
  get connected() { return !!this.pat; },

  _headers(extra = {}) {
    return {
      Authorization: `Bearer ${this.pat}`,
      Accept: 'application/vnd.github+json',
      'Content-Type': 'application/json',
      ...extra,
    };
  },

  /** Validate PAT by hitting /user; store it and auto-discover existing Gist. */
  async connect(pat) {
    const res = await fetch(`${GH_API}/user`, {
      headers: { Authorization: `Bearer ${pat}`, Accept: 'application/vnd.github+json' },
    });
    if (!res.ok) throw new Error(`GitHub auth failed (​${res.status})`);
    localStorage.setItem(GIST_PAT_KEY, pat);
    // Try to find an existing Gist so we can pull immediately on a new device
    await this.findGist();
    updateSyncStatus();
  },

  disconnect() {
    localStorage.removeItem(GIST_PAT_KEY);
    localStorage.removeItem(GIST_ID_KEY);
    updateSyncStatus();
  },

  /** Search the user’s Gist list for one containing our filename. */
  async findGist() {
    const res = await fetch(`${GH_API}/gists?per_page=100`, { headers: this._headers() });
    if (!res.ok) return null;
    const list  = await res.json();
    const found = list.find(g => g.files?.[GIST_FILENAME]);
    if (found) {
      localStorage.setItem(GIST_ID_KEY, found.id);
      return found.id;
    }
    return null;
  },

  /** Push current state JSON to Gist (create on first push, update thereafter). */
  async push(state) {
    if (!this.connected) return;
    const body = JSON.stringify({
      description: 'IRA Portfolio Rebalancing Dashboard — auto-save',
      public: false,
      files: { [GIST_FILENAME]: { content: JSON.stringify(state, null, 2) } },
    });
    let res;
    if (this.gistId) {
      res = await fetch(`${GH_API}/gists/${this.gistId}`,
        { method: 'PATCH', headers: this._headers(), body });
    } else {
      res = await fetch(`${GH_API}/gists`,
        { method: 'POST', headers: this._headers(), body });
      if (res.ok) {
        const data = await res.json();
        localStorage.setItem(GIST_ID_KEY, data.id);
      }
    }
    if (!res.ok) throw new Error(`Gist push failed (​${res.status})`);
    setSyncTimestamp();
  },

  /** Pull state JSON from Gist; returns parsed object or null. */
  async pull() {
    if (!this.connected) return null;
    // If we have no stored ID yet, try to discover it first
    if (!this.gistId) await this.findGist();
    if (!this.gistId) return null;
    const res = await fetch(`${GH_API}/gists/${this.gistId}`, { headers: this._headers() });
    if (!res.ok) throw new Error(`Gist pull failed (​${res.status})`);
    const data    = await res.json();
    const content = data.files?.[GIST_FILENAME]?.content;
    if (!content) return null;
    return JSON.parse(content);
  },
};

/** Update the cloud-sync button appearance in the site header. */
function updateSyncStatus() {
  const btn = document.getElementById('btn-cloud-sync');
  if (!btn) return;
  if (gistSync.connected) {
    btn.classList.add('connected');
    btn.title = 'Cloud Sync — Connected to GitHub Gist (click to manage)';
  } else {
    btn.classList.remove('connected');
    btn.title = 'Cloud Sync — Not connected (click to set up)';
  }
}

/** Write the “last synced” timestamp to the sync modal status line. */
function setSyncTimestamp() {
  const el = document.getElementById('sync-last-saved');
  if (el) {
    el.textContent = 'Last synced: ' + new Date().toLocaleTimeString();
    el.className = 'sync-status-msg sync-ok';
  }
}

/** Show a message in the sync modal status line. */
function setSyncMsg(msg, type = 'info') {
  const el = document.getElementById('sync-last-saved');
  if (el) {
    el.textContent = msg;
    el.className   = `sync-status-msg sync-${type}`;
  }
}

/**
 * Returns true if the given portfolios array is identical to the built-in
 * demo defaults (matched by portfolio count, holding count, and all tickers).
 * Used as a second-layer guard to prevent auto-pushing demo data to Gist.
 */
function _looksLikeDemo(portfolios) {
  const defaults = getDefaultPortfolios();
  if (!Array.isArray(portfolios) || portfolios.length !== defaults.length) return false;
  return portfolios.every((p, pi) => {
    const d = defaults[pi];
    if (!p.holdings || p.holdings.length !== d.holdings.length) return false;
    return p.holdings.every((h, hi) => h.ticker === d.holdings[hi].ticker);
  });
}

/**
 * Read every editable field + fetched mkt prices from the live DOM
 * and persist to localStorage as JSON.
 */
function saveState() {
  try {
    const cards = Array.from(document.querySelectorAll('.portfolio-card'));
    const state = {
      portfolioSeq: _portfolioSeq,
      savedAt: new Date().toISOString(),
      portfolios: cards.map(card => {
        const id       = card.id.replace('card-', '');
        const name     = card.querySelector('.card-title')?.textContent     || '';
        const subtitle = card.querySelector('.card-subtitle')?.textContent  || '';
        const rows     = Array.from(card.querySelectorAll('tbody tr[data-row]'));
        const holdings = rows.map(row => ({
          ticker   : row.querySelector('[data-ticker]')?.value              || '',
          shares   : toNum(row.querySelector('[data-shares]')?.value),
          price    : toNum(row.querySelector('[data-cost-basis]')?.value),
          targetPct: toNum(row.querySelector('[data-target-pct]')?.value),
          mktPrice : toNum(row.querySelector('[data-mkt-price]')?.dataset.raw),
        }));
        return { id, name, subtitle, holdings };
      }),
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    flashSaveIndicator();
    // Auto-push to Gist only when real user data is confirmed loaded.
    // _gistPushEnabled stays false on a fresh origin until a pull/load succeeds,
    // which prevents demo data from silently overwriting the real Gist.
    if (gistSync.connected && _gistPushEnabled && !_looksLikeDemo(state.portfolios)) {
      gistSync.push(state).catch(e => console.warn('Gist push failed:', e));
    }
  } catch (e) {
    console.warn('saveState failed:', e);
  }
}

/**
 * Load persisted state from localStorage into PORTFOLIOS.
 * Returns true if valid saved data was found, false otherwise.
 */
function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return false;
    const state = JSON.parse(raw);
    if (!Array.isArray(state?.portfolios) || state.portfolios.length === 0) return false;
    PORTFOLIOS = state.portfolios;
    _portfolioSeq = state.portfolioSeq ?? state.portfolios.length;
    _gistPushEnabled = true;  // real data confirmed — safe to auto-push
    return true;
  } catch (e) {
    console.warn('loadState failed:', e);
    return false;
  }
}

/**
 * Wipe localStorage and reload the original demo portfolios.
 * Called by the "Reset to Demo" button in the site header.
 */
function resetToDemo() {
  if (!confirm('Reset all portfolios to demo data? All your changes will be lost.')) return;
  localStorage.removeItem(STORAGE_KEY);
  PORTFOLIOS = getDefaultPortfolios();
  _portfolioSeq = PORTFOLIOS.length;
  _gistPushEnabled = false;  // don't auto-push demo data
  renderDashboard();
}

/** Flash a brief "Saved" indicator in the site header. */
function flashSaveIndicator() {
  const el = document.getElementById('save-indicator');
  if (!el) return;
  el.classList.add('visible');
  clearTimeout(el._timer);
  el._timer = setTimeout(() => el.classList.remove('visible'), 1800);
}

/* ================================================================
   INITIALISATION
   ================================================================ */
document.addEventListener('DOMContentLoaded', () => {
  // 1. Try to pull from Gist first (if connected); fall back to localStorage.
  if (gistSync.connected && gistSync.gistId) {
    gistSync.pull()
      .then(state => {
        if (state?.portfolios?.length) {
          PORTFOLIOS     = state.portfolios;
          _portfolioSeq  = state.portfolioSeq ?? state.portfolios.length;
          // Also refresh localStorage so offline fallback stays current
          localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
          _gistPushEnabled = true;  // real data loaded — safe to auto-push
          setSyncTimestamp();
        } else {
          loadState();
        }
      })
      .catch(e => {
        console.warn('Gist pull on load failed, using localStorage:', e);
        loadState();
      })
      .finally(() => {
        setHeaderDate();
        renderDashboard();
        initModal();
        initSyncModal();
        updateSyncStatus();
      });
  } else {
    loadState();   // sets _gistPushEnabled=true if real data found
    setHeaderDate();
    renderDashboard();
    initModal();
    initSyncModal();
    updateSyncStatus();
  }
});

function setHeaderDate() {
  const el = document.getElementById('current-date');
  if (!el) return;
  el.textContent = new Date().toLocaleDateString('en-US', {
    weekday: 'short', month: 'short', day: 'numeric', year: 'numeric',
  });
}

/* ================================================================
   DASHBOARD RENDER
   ================================================================ */
function renderDashboard() {
  const dashboard = document.getElementById('dashboard');
  if (!dashboard) return;
  dashboard.innerHTML = '';

  PORTFOLIOS.forEach((portfolio, idx) => {
    const card = buildPortfolioCard(portfolio, idx + 1);
    dashboard.appendChild(card);
    recalculate(portfolio.id);   // initial calculation
  });

  // "Add Portfolio" bar — always anchored at the bottom of the dashboard
  const addBar = document.createElement('div');
  addBar.id = 'add-portfolio-bar';
  addBar.className = 'add-portfolio-bar';
  addBar.innerHTML = `
    <button class="btn btn-add-portfolio" onclick="addPortfolio()"
            title="Add a new demo portfolio">
      <span aria-hidden="true">＋</span> Add Portfolio
    </button>
  `;
  dashboard.appendChild(addBar);

  updateAggregateStats();
}

/* ================================================================
   BUILD PORTFOLIO CARD (DOM construction)
   ================================================================ */
function buildPortfolioCard(portfolio, cardNumber) {
  const section = document.createElement('section');
  section.className = 'portfolio-card';
  section.id = `card-${portfolio.id}`;

  section.innerHTML = `
    <div class="card-header">
      <div class="card-title-group">
        <div class="card-index-badge">${cardNumber}</div>
        <div>
          <div class="card-title"
               contenteditable="true"
               spellcheck="false"
               data-original="${escAttr(portfolio.name)}"
               onkeydown="handleTitleKeydown(event, this)"
               onblur="commitTitleEdit(this, '${escAttr(portfolio.id)}')"
               title="Click to rename"
               >${escHTML(portfolio.name)}</div>
          <div class="card-subtitle"
               contenteditable="true"
               spellcheck="false"
               data-original="${escAttr(portfolio.subtitle)}"
               onkeydown="handleTitleKeydown(event, this)"
               onblur="commitTitleEdit(this, '${escAttr(portfolio.id)}')"
               title="Click to edit subtitle"
               >${escHTML(portfolio.subtitle)}</div>
        </div>
      </div>
      <div class="card-summary">
        <div class="sum-item">
          <span class="sum-label">Total Value</span>
          <span class="sum-value" data-sum-total-value>—</span>
        </div>
        <div class="sum-item">
          <span class="sum-label">Target &Sigma;</span>
          <span class="sum-value" data-sum-target-total>—</span>
        </div>
        <div class="sum-item">
          <span class="sum-label">Holdings</span>
          <span class="sum-value" data-sum-holding-count>—</span>
        </div>
        <div class="sum-item">
          <span class="sum-label">Drift Alerts</span>
          <span class="sum-value" data-sum-alert-count>—</span>
        </div>
      </div>
      <button class="btn-del-portfolio"
              onclick="deletePortfolio('${escAttr(portfolio.id)}')"
              title="Remove this portfolio"
              aria-label="Remove portfolio">
        ⊗ Remove
      </button>
    </div>

    <div class="target-warning" data-target-warning></div>

    <div class="table-wrapper">
      <table class="holdings-table" aria-label="${escAttr(portfolio.name)} holdings">
        <thead>
          <tr>
            <th class="th-left col-ticker">Ticker</th>
            <th class="col-shares">Shares</th>
            <th class="col-costbasis">Avg&nbsp;Cost</th>
            <th class="col-mktprice">Mkt&nbsp;Price</th>
            <th class="col-curval">Current Value</th>
            <th class="col-gainloss">Gain / Loss</th>
            <th class="col-target">Target&nbsp;%</th>
            <th class="col-curpct">Current&nbsp;%</th>
            <th class="col-drift">Drift&nbsp;%</th>
            <th class="th-left col-trade">Suggested Trade</th>
            <th class="col-action" aria-label="Row actions"></th>
          </tr>
        </thead>
        <tbody data-tbody="${escAttr(portfolio.id)}"></tbody>
      </table>
    </div>

    <div class="card-footer">
      <div class="footer-actions">
        <button class="btn btn-recalc"
                onclick="recalculate('${escAttr(portfolio.id)}')"
                title="Recalculate all derived values">
          <span aria-hidden="true">⟳</span> Recalculate
        </button>
        <button class="btn btn-fetchprices"
                data-fetch-btn="${escAttr(portfolio.id)}"
                onclick="fetchPrices('${escAttr(portfolio.id)}')"
                title="Fetch live market prices from Yahoo Finance">
          <span aria-hidden="true">↺</span> Fetch Prices
        </button>
        <button class="btn btn-rebalance"
                onclick="generateRebalanceSuggestion('${escAttr(portfolio.id)}')"
                title="Generate rebalance suggestion text">
          <span aria-hidden="true">◎</span> Rebalance Suggestion
        </button>
        <button class="btn btn-export"
                onclick="exportToCSV('${escAttr(portfolio.id)}')"
                title="Download this portfolio as a CSV file">
          <span aria-hidden="true">&#8595;</span> Export CSV
        </button>
        <div class="btn-divider" role="separator"></div>
        <button class="btn btn-add-row"
                onclick="addRow('${escAttr(portfolio.id)}')"
                title="Add a new holding row">
          <span aria-hidden="true">＋</span> Add Row
        </button>
      </div>
      <span class="footer-meta" data-last-calc></span>
    </div>
  `;

  // Populate tbody with initial holdings
  const tbody = section.querySelector(`[data-tbody="${portfolio.id}"]`);
  portfolio.holdings.forEach(h => tbody.appendChild(buildRow(portfolio.id, h)));

  return section;
}

/* ================================================================
   BUILD A TABLE ROW
   ================================================================ */
function buildRow(portfolioId, holding = {}) {
  const id          = nextRowId();
  const ticker      = holding.ticker    ?? '';
  const shares      = holding.shares    ?? '';
  const costBasis   = holding.price     ?? '';  // purchase / avg cost price
  const targetPct   = holding.targetPct ?? '';
  const mktPriceRaw = holding.mktPrice  ?? 0;

  // Pre-compute mkt-price display so the template stays readable
  const mktPriceDisplay = mktPriceRaw > 0
    ? '$' + Number(mktPriceRaw).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
    : '—';
  const mktPriceCls = mktPriceRaw > 0 ? 'cell-ro mkt-price-live' : 'cell-ro';

  const tr = document.createElement('tr');
  tr.dataset.row = id;

  tr.innerHTML = `
    <td class="td-left col-ticker">
      <input type="text"
             class="tbl-input inp-ticker"
             data-ticker
             value="${escAttr(ticker)}"
             placeholder="TICK"
             maxlength="12"
             spellcheck="false"
             onblur="recalculate('${escAttr(portfolioId)}')"
             aria-label="Ticker symbol" />
    </td>
    <td class="col-shares">
      <input type="number"
             class="tbl-input inp-shares"
             data-shares
             value="${escAttr(String(shares))}"
             min="0" step="any" placeholder="0"
             onblur="recalculate('${escAttr(portfolioId)}')"
             aria-label="Number of shares" />
    </td>
    <td class="col-costbasis">
      <input type="number"
             class="tbl-input inp-price"
             data-cost-basis
             value="${escAttr(String(costBasis))}"
             min="0" step="any" placeholder="0.00"
             onblur="recalculate('${escAttr(portfolioId)}')"
             aria-label="Average cost / purchase price" />
    </td>
    <td class="col-mktprice">
      <span class="${escAttr(mktPriceCls)}" data-mkt-price data-raw="${escAttr(String(mktPriceRaw))}" aria-live="polite">${mktPriceDisplay}</span>
    </td>
    <td class="col-curval">
      <span class="cell-ro" data-current-value aria-live="polite">—</span>
    </td>
    <td class="col-gainloss">
      <span class="cell-ro gain-zero" data-gain-loss aria-live="polite">—</span>
    </td>
    <td class="col-target">
      <input type="number"
             class="tbl-input inp-target"
             data-target-pct
             value="${escAttr(String(targetPct))}"
             min="0" max="100" step="any" placeholder="0.00"
             onblur="recalculate('${escAttr(portfolioId)}')"
             aria-label="Target allocation percent" />
    </td>
    <td class="col-curpct">
      <span class="cell-ro" data-current-pct aria-live="polite">—</span>
    </td>
    <td class="col-drift">
      <span class="cell-ro drift-neutral" data-drift-pct aria-live="polite">—</span>
    </td>
    <td class="td-left col-trade">
      <span class="cell-ro trade-hold" data-suggested-trade aria-live="polite">—</span>
    </td>
    <td class="col-action">
      <button class="btn-del-row"
              onclick="deleteRow(this, '${escAttr(portfolioId)}')"
              title="Remove this holding"
              aria-label="Delete row">
        ✕
      </button>
    </td>
  `;

  return tr;
}

/* ================================================================
   ADD / DELETE / RENUMBER PORTFOLIOS
   ================================================================ */

/**
 * Append a new portfolio card pre-filled with the demo template.
 * The "Add Portfolio" bar always stays at the bottom.
 */
function addPortfolio() {
  _portfolioSeq++;
  const id = `portfolio-${_portfolioSeq}`;
  const newPortfolio = {
    id,
    name: `IRA Portfolio ${_portfolioSeq}`,
    subtitle: 'New Portfolio — Edit name and holdings as needed',
    holdings: DEMO_PORTFOLIO_TEMPLATE.map(h => ({ ...h })),
  };
  PORTFOLIOS.push(newPortfolio);

  const dashboard = document.getElementById('dashboard');
  const addBar    = document.getElementById('add-portfolio-bar');
  const card      = buildPortfolioCard(newPortfolio, PORTFOLIOS.length);
  dashboard.insertBefore(card, addBar);
  recalculate(id);
  updateAggregateStats();
  card.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

/**
 * Remove a portfolio card and its entry in PORTFOLIOS.
 * Prompts for confirmation when the portfolio has any filled holdings.
 */
function deletePortfolio(portfolioId) {
  const idx  = PORTFOLIOS.findIndex(p => p.id === portfolioId);
  if (idx === -1) return;

  const card = document.getElementById(`card-${portfolioId}`);
  if (!card) return;

  // Only confirm when the portfolio actually has rows with a ticker entered
  const hasTickers = Array.from(card.querySelectorAll('[data-ticker]'))
    .some(inp => inp.value.trim() !== '');

  if (hasTickers && !confirm(`Remove "${PORTFOLIOS[idx].name}" and all its holdings?`)) return;

  PORTFOLIOS.splice(idx, 1);
  card.remove();
  renumberCards();
  updateAggregateStats();
  saveState();
}

/** Re-sync the visible card-index badge numbers after any add/remove. */
function renumberCards() {
  Array.from(document.querySelectorAll('.portfolio-card')).forEach((card, idx) => {
    const badge = card.querySelector('.card-index-badge');
    if (badge) badge.textContent = idx + 1;
  });
}

/* ================================================================
   ADD / DELETE ROWS
   ================================================================ */
function addRow(portfolioId) {
  const tbody = document.querySelector(`[data-tbody="${portfolioId}"]`);
  if (!tbody) return;
  const newRow = buildRow(portfolioId);
  tbody.appendChild(newRow);
  recalculate(portfolioId);
  // Focus ticker field so user can type immediately
  newRow.querySelector('[data-ticker]')?.focus();
}

function deleteRow(btnEl, portfolioId) {
  const tr = btnEl.closest('tr');
  if (!tr) return;
  tr.remove();
  recalculate(portfolioId);
}

/* ================================================================
   RECALCULATE  —  core engine
   ================================================================ */
function recalculate(portfolioId) {
  const card = document.getElementById(`card-${portfolioId}`);
  if (!card) return;

  const rows = Array.from(card.querySelectorAll('tbody tr[data-row]'));

  /* ── Pass 1: collect raw values, sum total portfolio value ── */
  let totalValue = 0;
  const rowData = rows.map(row => {
    const shares     = toNum(row.querySelector('[data-shares]')?.value);
    const costBasis  = toNum(row.querySelector('[data-cost-basis]')?.value);
    const mktPrice   = toNum(row.querySelector('[data-mkt-price]')?.dataset.raw);
    // Use live market price when available, otherwise fall back to cost basis
    const price      = mktPrice > 0 ? mktPrice : costBasis;
    const targetPct  = toNum(row.querySelector('[data-target-pct]')?.value);
    const currentValue = shares * price;
    totalValue += currentValue;
    return { row, shares, costBasis, price, targetPct, currentValue };
  });

  /* ── Pass 2: derive % values, classify drift, update cells ── */
  let totalTargetPct = 0;
  let alertCount = 0;

  rowData.forEach(({ row, shares, costBasis, price, targetPct, currentValue }) => {
    totalTargetPct += targetPct;

    const currentPct = totalValue > 0 ? (currentValue / totalValue) * 100 : 0;
    const driftPct   = currentPct - targetPct;

    /* Store computed values as data attributes for the suggestion generator */
    row.dataset.computedCurrentValue = currentValue;
    row.dataset.computedCurrentPct   = currentPct;
    row.dataset.computedDriftPct     = driftPct;

    /* Current Value cell */
    const cvEl = row.querySelector('[data-current-value]');
    if (cvEl) cvEl.textContent = fmtCurrency(currentValue);

    /* Current % cell */
    const cpEl = row.querySelector('[data-current-pct]');
    if (cpEl) {
      cpEl.textContent = fmtPct(currentPct);
      cpEl.className   = 'cell-ro';
    }

    /* Drift % cell — coloured, signed, classified */
    const driftEl = row.querySelector('[data-drift-pct]');
    if (driftEl) {
      driftEl.textContent = fmtDrift(driftPct);
      driftEl.className   = `cell-ro ${driftClass(driftPct)}`;
    }

    /* Gain / Loss cell: (mktPrice − costBasis) × shares */
    const glEl = row.querySelector('[data-gain-loss]');
    if (glEl) {
      if (price > 0 && costBasis > 0) {
        const gainLoss = (price - costBasis) * shares;
        const pct      = (((price - costBasis) / costBasis) * 100).toFixed(2);
        const sign     = gainLoss >= 0 ? '+' : '';
        glEl.textContent = `${sign}${fmtCurrency(gainLoss)} (${sign}${pct}%)`;
        glEl.className   = `cell-ro ${gainLoss > 0 ? 'gain-pos' : gainLoss < 0 ? 'gain-neg' : 'gain-zero'}`;
      } else {
        glEl.textContent = '—';
        glEl.className   = 'cell-ro gain-zero';
      }
    }

    /* Suggested Trade cell — auto-generated, read-only */
    const tradeEl = row.querySelector('[data-suggested-trade]');
    if (tradeEl) {
      const { text, cls } = calcSuggestedTrade(driftPct, targetPct, currentValue, totalValue, price);
      tradeEl.textContent = text;
      tradeEl.className   = `cell-ro ${cls}`;
    }

    /* Alert row — abs(drift) > 3 % */
    if (Math.abs(driftPct) > 3) {
      row.classList.add('row-alert');
      alertCount++;
    } else {
      row.classList.remove('row-alert');
    }
  });

  /* ── Update summary strip in card header ─────────────────── */
  const totalValueEl  = card.querySelector('[data-sum-total-value]');
  const targetTotalEl = card.querySelector('[data-sum-target-total]');
  const holdingCntEl  = card.querySelector('[data-sum-holding-count]');
  const alertCntEl    = card.querySelector('[data-sum-alert-count]');
  const warningEl     = card.querySelector('[data-target-warning]');
  const lastCalcEl    = card.querySelector('[data-last-calc]');

  if (totalValueEl)  totalValueEl.textContent  = fmtCurrency(totalValue);
  if (holdingCntEl)  holdingCntEl.textContent  = rows.length;

  if (alertCntEl) {
    alertCntEl.textContent = alertCount;
    alertCntEl.className   = `sum-value ${alertCount > 0 ? 'v-danger' : 'v-ok'}`;
  }

  /* Target % validation */
  const targetDiff = Math.abs(totalTargetPct - 100);

  if (targetTotalEl) {
    targetTotalEl.textContent = fmtPct(totalTargetPct);
    targetTotalEl.className   = `sum-value ${targetDiff < 0.01 ? 'v-ok' : 'v-danger'}`;
  }

  if (warningEl) {
    if (targetDiff >= 0.01 && rows.length > 0) {
      warningEl.textContent = `⚠  Target allocations sum to ${fmtPct(totalTargetPct)} — must equal exactly 100.00%`;
      warningEl.classList.add('visible');
    } else {
      warningEl.textContent = '';
      warningEl.classList.remove('visible');
    }
  }

  if (lastCalcEl) {
    lastCalcEl.textContent = `Last calculated: ${new Date().toLocaleTimeString('en-US', {
      hour: '2-digit', minute: '2-digit', second: '2-digit',
    })}`;
  }

  /* Refresh global aggregate stats in site header */
  updateAggregateStats();
  saveState();
}

/* ================================================================
   FETCH LIVE PRICES  (Yahoo Finance v8 — multi-proxy cascade)
   ================================================================ */

/**
 * Fetches live prices for every ticker.
 *
 * Strategy 1 — Yahoo Finance v7 BATCH quote: all tickers in ONE request,
 *   tried through 3 CORS proxies.  A single HTTP call avoids rate-limiting
 *   that kills per-ticker parallel requests.
 *
 * Strategy 2 — Individual fallback: for any ticker not covered by the batch,
 *   try Yahoo v8 chart then Stooq CSV, each through the same 3 proxies.
 *
 * Proxy A: allorigins.win/raw   Proxy B: corsproxy.io   Proxy C: codetabs.com
 */
async function fetchPrices(portfolioId) {
  const card = document.getElementById(`card-${portfolioId}`);
  if (!card) return;

  const rows = Array.from(card.querySelectorAll('tbody tr[data-row]'));

  /* ticker → rows map (handles duplicate tickers) */
  const tickerRowMap = {};
  rows.forEach(row => {
    const t = (row.querySelector('[data-ticker]')?.value || '').trim().toUpperCase();
    if (t) (tickerRowMap[t] = tickerRowMap[t] || []).push(row);
  });

  const uniqueTickers = Object.keys(tickerRowMap);
  if (uniqueTickers.length === 0) return;

  /* ── Loading state ─────────────────────────────────────── */
  const btn = card.querySelector(`[data-fetch-btn="${portfolioId}"]`);
  const originalHTML = btn ? btn.innerHTML : '';
  if (btn) {
    btn.disabled = true;
    btn.innerHTML = `<span aria-hidden="true" class="spin-icon">⟳</span> Fetching…`;
  }

  /* Three CORS proxies — tried in order for every target URL */
  const proxyWrap = [
    url => `https://api.allorigins.win/raw?url=${encodeURIComponent(url)}`,
    url => `https://corsproxy.io/?${encodeURIComponent(url)}`,
    url => `https://api.codetabs.com/v1/proxy?quest=${encodeURIComponent(url)}`,
  ];

  /* ticker → price  (filled in below) */
  const priceMap = {};

  /* ── Strategy 1: single batch request via Yahoo v7 quote ── */
  /* One proxy call returns all tickers — avoids rate-limit spam */
  const batchUrl =
    `https://query1.finance.yahoo.com/v7/finance/quote` +
    `?symbols=${uniqueTickers.map(encodeURIComponent).join(',')}` +
    `&fields=regularMarketPrice`;

  for (const makeProxy of proxyWrap) {
    const ctrl  = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), 12000);
    try {
      const res = await fetch(makeProxy(batchUrl), { signal: ctrl.signal });
      clearTimeout(timer);
      if (!res.ok) continue;
      const data    = await res.json();
      const results = data?.quoteResponse?.result;
      if (Array.isArray(results) && results.length > 0) {
        results.forEach(q => {
          if (q?.regularMarketPrice != null)
            priceMap[q.symbol.toUpperCase()] = q.regularMarketPrice;
        });
        break;   /* batch succeeded — stop trying other proxies */
      }
    } catch (e) {
      clearTimeout(timer);
    }
  }

  /* ── Strategy 2: individual fallback for any still-missing tickers ── */
  async function fetchOne(ticker) {
    let lastErr;

    /* Yahoo v8 chart (per-ticker) */
    const yahooUrl =
      `https://query2.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(ticker)}` +
      `?interval=1d&range=1d&includePrePost=false`;

    for (const makeProxy of proxyWrap) {
      const ctrl  = new AbortController();
      const timer = setTimeout(() => ctrl.abort(), 8000);
      try {
        const res = await fetch(makeProxy(yahooUrl), { signal: ctrl.signal });
        clearTimeout(timer);
        if (!res.ok) { lastErr = new Error(`HTTP ${res.status}`); continue; }
        const data = await res.json();
        if (data?.chart?.error) {
          lastErr = new Error(data.chart.error.description || 'Yahoo error');
          continue;
        }
        const price = data?.chart?.result?.[0]?.meta?.regularMarketPrice;
        if (price != null) return price;
        lastErr = new Error('No price in Yahoo v8 response');
      } catch (e) {
        clearTimeout(timer);
        lastErr = e;
      }
    }

    /* Stooq CSV — plain ticker then ticker.US */
    for (const sym of [ticker, `${ticker}.US`]) {
      const stooqUrl =
        `https://stooq.com/q/l/?s=${encodeURIComponent(sym)}&f=sd2t2ohlcv&e=csv`;
      for (const makeProxy of proxyWrap) {
        const ctrl  = new AbortController();
        const timer = setTimeout(() => ctrl.abort(), 8000);
        try {
          const res = await fetch(makeProxy(stooqUrl), { signal: ctrl.signal });
          clearTimeout(timer);
          if (!res.ok) { lastErr = new Error(`HTTP ${res.status}`); continue; }
          const text   = await res.text();
          /* Stooq CSV: Symbol,Date,Time,Open,High,Low,Close,Volume */
          const fields = text.trim().split('\n').pop().split(',');
          const close  = parseFloat(fields[6]);
          if (!isNaN(close) && close > 0) return close;
          lastErr = new Error('Stooq: no valid price');
        } catch (e) {
          clearTimeout(timer);
          lastErr = e;
        }
      }
    }

    throw lastErr || new Error('All sources and proxies failed');
  }

  const missing = uniqueTickers.filter(t => priceMap[t] == null);
  if (missing.length > 0) {
    const fallbackResults = await Promise.allSettled(missing.map(t => fetchOne(t)));
    missing.forEach((ticker, idx) => {
      if (fallbackResults[idx].status === 'fulfilled')
        priceMap[ticker] = fallbackResults[idx].value;
      else
        console.warn(`fetchPrices — ${ticker}:`, fallbackResults[idx].reason?.message);
    });
  }

  /* ── Apply prices to DOM ───────────────────────────────── */
  let fetched = 0;
  uniqueTickers.forEach(ticker => {
    const price = priceMap[ticker];
    (tickerRowMap[ticker] || []).forEach(row => {
      const mktEl = row.querySelector('[data-mkt-price]');
      if (!mktEl) return;
      if (price != null) {
        mktEl.dataset.raw = price;
        mktEl.textContent = '$' + Number(price).toLocaleString('en-US', {
          minimumFractionDigits: 2, maximumFractionDigits: 2,
        });
        mktEl.className = 'cell-ro mkt-price-live';
        fetched++;
      } else {
        mktEl.dataset.raw = '0';
        mktEl.textContent = 'N/A';
        mktEl.className   = 'cell-ro mkt-price-na';
      }
    });
  });

  recalculate(portfolioId);

  /* ── Button feedback ───────────────────────────────────── */
  if (btn) {
    btn.innerHTML = fetched === 0
      ? `<span aria-hidden="true">⚠</span> No prices returned`
      : `<span aria-hidden="true">✓</span> Updated (${fetched}/${uniqueTickers.length})`;
    setTimeout(() => {
      btn.innerHTML = originalHTML;
      btn.disabled  = false;
    }, 3200);
  }
}

/* ================================================================
   REBALANCE SUGGESTION GENERATOR
   ================================================================ */
function generateRebalanceSuggestion(portfolioId) {
  /* Ensure values are fresh before building the report */
  recalculate(portfolioId);

  const card      = document.getElementById(`card-${portfolioId}`);
  const portfolio = PORTFOLIOS.find(p => p.id === portfolioId);
  if (!card || !portfolio) return;

  const rows = Array.from(card.querySelectorAll('tbody tr[data-row]'));
  if (rows.length === 0) {
    showModal(`Rebalance Suggestion — ${portfolio.name}`, '  No holdings to analyse.');
    return;
  }

  /* Total value from stored computed attributes */
  const totalValue = rows.reduce(
    (s, r) => s + (parseFloat(r.dataset.computedCurrentValue) || 0), 0
  );

  const dateStr = new Date().toLocaleString('en-US', {
    month: 'long', day: 'numeric', year: 'numeric',
    hour: '2-digit', minute: '2-digit',
  });

  /* ── Column widths (characters) ─────────────────────────── */
  const W = { ticker: 8, value: 16, target: 10, current: 11, drift: 10 };

  /* ── Table header row ───────────────────────────────────── */
  const colHeader = [
    'Ticker'.padEnd(W.ticker),
    'Current Value'.padStart(W.value),
    'Target %'.padStart(W.target),
    'Current %'.padStart(W.current),
    'Drift %'.padStart(W.drift),
    '  Status',
  ].join('');

  const rowSep = '─'.repeat(colHeader.length);

  /* ── Build data lines ───────────────────────────────────── */
  const alertTickers = [];
  const dataLines = rows.map(row => {
    const ticker     = (row.querySelector('[data-ticker]')?.value || '').toUpperCase().trim() || '—';
    const curVal     = parseFloat(row.dataset.computedCurrentValue) || 0;
    const curPct     = parseFloat(row.dataset.computedCurrentPct)   || 0;
    const driftPct   = parseFloat(row.dataset.computedDriftPct)     || 0;
    const targetPct  = toNum(row.querySelector('[data-target-pct]')?.value);
    const isAlert    = Math.abs(driftPct) > 3;

    if (isAlert) alertTickers.push(ticker);

    const sign     = driftPct >= 0 ? '+' : '';
    const driftStr = `${sign}${driftPct.toFixed(2)}%`;
    const status   = isAlert ? '  ◄ REBALANCE NEEDED' : '  ✓';

    return [
      ticker.padEnd(W.ticker),
      fmtCurrency(curVal).padStart(W.value),
      (targetPct.toFixed(2) + '%').padStart(W.target),
      (curPct.toFixed(2)    + '%').padStart(W.current),
      driftStr.padStart(W.drift),
      status,
    ].join('');
  });

  /* ── Total row ──────────────────────────────────────────── */
  const totalRow = [
    'TOTAL'.padEnd(W.ticker),
    fmtCurrency(totalValue).padStart(W.value),
    ''.padStart(W.target),
    ''.padStart(W.current),
    ''.padStart(W.drift),
    '',
  ].join('');

  /* ── Footer note ────────────────────────────────────────── */
  const footerNote = alertTickers.length > 0
    ? `⚠   ${alertTickers.length} position(s) require rebalancing (|Drift %| > 3%): ${alertTickers.join(', ')}`
    : '✓   All positions are within tolerance — no rebalancing action required.';

  /* ── Assemble full text ─────────────────────────────────── */
  const BORDER = '═'.repeat(colHeader.length);
  const lines = [
    BORDER,
    `  REBALANCE SUGGESTION — ${portfolio.name.toUpperCase()}`,
    `  ${portfolio.subtitle}`,
    `  Generated: ${dateStr}`,
    BORDER,
    '',
    `  Total Portfolio Value: ${fmtCurrency(totalValue)}`,
    '',
    `  ${colHeader}`,
    `  ${rowSep}`,
    ...dataLines.map(l => `  ${l}`),
    `  ${rowSep}`,
    `  ${totalRow}`,
    '',
    `  ${footerNote}`,
    '',
    BORDER,
  ];

  const text = lines.join('\n');

  /* Show modal (always) */
  showModal(`Rebalance Suggestion — ${portfolio.name}`, text);

  /* Also attempt to auto-copy */
  copyText(text)
    .then(() => showCopyFeedback('Auto-copied to clipboard'))
    .catch(() => { /* silent — user can still copy manually */ });
}

/* ================================================================
   AGGREGATE STATS (site header)
   ================================================================ */
function updateAggregateStats() {
  const statsEl = document.getElementById('aggregate-stats');
  if (!statsEl) return;

  let grandTotal  = 0;
  let alertTotal  = 0;
  let portCount   = 0;

  PORTFOLIOS.forEach(p => {
    const card = document.getElementById(`card-${p.id}`);
    if (!card) return;
    portCount++;
    Array.from(card.querySelectorAll('tbody tr[data-row]')).forEach(row => {
      grandTotal += parseFloat(row.dataset.computedCurrentValue) || 0;
      if (Math.abs(parseFloat(row.dataset.computedDriftPct) || 0) > 3) alertTotal++;
    });
  });

  statsEl.innerHTML = `
    <div class="stat-item">
      <span class="stat-label">Total AUM</span>
      <span class="stat-value">${fmtCurrency(grandTotal)}</span>
    </div>
    <div class="stat-item">
      <span class="stat-label">Portfolios</span>
      <span class="stat-value">${portCount}</span>
    </div>
    <div class="stat-item">
      <span class="stat-label">Drift Alerts</span>
      <span class="stat-value ${alertTotal > 0 ? 's-warn' : 's-ok'}">${alertTotal}</span>
    </div>
  `;
}

/* ================================================================
   MODAL
   ================================================================ */
function initModal() {
  const backdrop = document.getElementById('modal-backdrop');
  const closeBtn  = document.getElementById('modal-close');
  const copyBtn   = document.getElementById('btn-copy-suggestion');

  closeBtn?.addEventListener('click', closeModal);

  /* Close on backdrop click */
  backdrop?.addEventListener('click', e => {
    if (e.target === backdrop) closeModal();
  });

  /* Close on Escape */
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') closeModal();
  });

  /* Manual copy button inside modal */
  copyBtn?.addEventListener('click', () => {
    const text = document.getElementById('suggestion-textarea')?.value ?? '';
    copyText(text)
      .then(() => showCopyFeedback('Copied!'))
      .catch(() => {
        /* Fallback: select all text so user can Ctrl+C manually */
        const ta = document.getElementById('suggestion-textarea');
        ta?.select();
      });
  });
}

function showModal(title, text) {
  const titleEl    = document.getElementById('modal-title');
  const textareaEl = document.getElementById('suggestion-textarea');
  const backdrop   = document.getElementById('modal-backdrop');

  if (titleEl)    titleEl.textContent    = title;
  if (textareaEl) textareaEl.value       = text;
  if (backdrop)   backdrop.classList.add('active');
}

function closeModal() {
  document.getElementById('modal-backdrop')?.classList.remove('active');
}

function showCopyFeedback(msg = 'Copied!') {
  const el = document.getElementById('copy-confirm');
  if (!el) return;
  el.textContent = `✓ ${msg}`;
  el.classList.add('visible');
  setTimeout(() => el.classList.remove('visible'), 2600);
}

/* Clipboard — async API with legacy textarea fallback */
async function copyText(text) {
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(text);
  } else {
    /* Legacy execCommand fallback */
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.cssText = 'position:fixed;opacity:0;top:0;left:0;';
    document.body.appendChild(ta);
    ta.focus();
    ta.select();
    document.execCommand('copy');
    document.body.removeChild(ta);
  }
}

/* ================================================================
   SUGGESTED TRADE CALCULATOR
   ================================================================ */

/**
 * Returns { text, cls } for the Suggested Trade cell.
 *
 * Logic:
 *   targetValue  = (targetPct / 100) * totalPortfolioValue
 *   dollarDiff   = targetValue - currentValue   (+ = need to buy, - = need to sell)
 *   sharesToTrade = |dollarDiff| / price
 *
 * Thresholds:
 *   |driftPct| < 0.1 %  → Hold (within noise)
 *   otherwise           → Buy / Sell
 */
function calcSuggestedTrade(driftPct, targetPct, currentValue, totalValue, price) {
  if (totalValue <= 0 || price <= 0) return { text: '—', cls: 'trade-hold' };

  const absDrift = Math.abs(driftPct);
  if (absDrift < 0.1) return { text: 'Hold', cls: 'trade-hold' };

  const targetValue  = (targetPct / 100) * totalValue;
  const dollarDiff   = targetValue - currentValue;
  const sharesNeeded = Math.abs(dollarDiff) / price;

  // Display whole shares when ≥ 1, otherwise 3 decimal places
  const sharesStr = sharesNeeded >= 1
    ? Math.round(sharesNeeded).toLocaleString('en-US')
    : sharesNeeded.toFixed(3);

  if (dollarDiff > 0) {
    return {
      text: `Buy ${sharesStr} sh  (${fmtCurrency(dollarDiff)})`,
      cls:  'trade-buy',
    };
  }
  return {
    text: `Sell ${sharesStr} sh  (${fmtCurrency(Math.abs(dollarDiff))})`,
    cls:  'trade-sell',
  };
}

/* ================================================================
   EDITABLE PORTFOLIO NAME / SUBTITLE
   ================================================================ */

/**
 * Enter  → commit (blur).
 * Escape → revert to the value stored in data-original, then blur.
 * Prevent newlines — these are single-line labels.
 */
function handleTitleKeydown(e, el) {
  if (e.key === 'Enter') {
    e.preventDefault();
    el.blur();
  } else if (e.key === 'Escape') {
    e.preventDefault();
    el.textContent = el.dataset.original || '';
    el.blur();
  }
}

/**
 * On blur: strip any pasted HTML (read as plain text), trim, enforce
 * a non-empty fallback, then persist via saveState().
 */
function commitTitleEdit(el, portfolioId) {   // portfolioId kept for future hooks
  void portfolioId;  // not needed currently; saveState reads DOM directly
  // Read as plain text to discard any pasted HTML
  const clean = el.textContent.replace(/\s+/g, ' ').trim();
  el.textContent = clean || el.dataset.original || 'Untitled';
  // Update the stored original so subsequent Escape reverts to the new value
  el.dataset.original = el.textContent;
  saveState();
}

/* ================================================================
   EXPORT TO CSV
   ================================================================ */

/**
 * Build a CSV file from the live DOM state of a portfolio card and
 * trigger a browser download.  Column order mirrors the table:
 *   Ticker, Shares, Avg Cost, Mkt Price, Current Value,
 *   Gain/Loss $, Gain/Loss %, Target %, Current %, Drift %, Suggested Trade
 */
function exportToCSV(portfolioId) {
  const card = document.getElementById(`card-${portfolioId}`);
  if (!card) return;

  const portfolioName = card.querySelector('.card-title')?.textContent?.trim() || portfolioId;
  const rows          = Array.from(card.querySelectorAll('tbody tr[data-row]'));

  const headers = [
    'Ticker', 'Shares', 'Avg Cost ($)', 'Mkt Price ($)',
    'Current Value ($)', 'Gain/Loss', 'Target %',
    'Current %', 'Drift %', 'Suggested Trade',
  ];

  const dataRows = rows.map(row => {
    const ticker      = row.querySelector('[data-ticker]')?.value?.trim()   || '';
    const shares      = row.querySelector('[data-shares]')?.value?.trim()   || '';
    const avgCost     = row.querySelector('[data-cost-basis]')?.value?.trim() || '';
    const mktRaw      = toNum(row.querySelector('[data-mkt-price]')?.dataset.raw);
    const mktPrice    = mktRaw > 0 ? mktRaw.toFixed(2) : '';
    const curVal      = row.querySelector('[data-current-value]')?.textContent?.replace(/[$,]/g, '').trim() || '';
    const gainLoss    = row.querySelector('[data-gain-loss]')?.textContent?.trim() || '';
    const targetPct   = row.querySelector('[data-target-pct]')?.value?.trim()  || '';
    const curPct      = row.querySelector('[data-current-pct]')?.textContent?.trim() || '';
    const driftPct    = row.querySelector('[data-drift-pct]')?.textContent?.trim()   || '';
    const trade       = row.querySelector('[data-suggested-trade]')?.textContent?.trim() || '';
    return [ticker, shares, avgCost, mktPrice, curVal, gainLoss, targetPct, curPct, driftPct, trade];
  });

  // Wrap any field that contains a comma, quote, or newline in double-quotes;
  // escape embedded double-quotes by doubling them (RFC 4180).
  const escape = v => {
    const s = String(v ?? '');
    return /[,"\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };

  const csvLines = [
    headers.map(escape).join(','),
    ...dataRows.map(r => r.map(escape).join(',')),
  ];

  // Append a summary footer row with the portfolio total value
  const totalEl = card.querySelector('[data-total-value]');
  if (totalEl) {
    const totalRaw = totalEl.textContent.replace(/[$,]/g, '').trim();
    csvLines.push('');
    csvLines.push(`Portfolio Total,,,,,${escape(fmtCurrency(parseFloat(totalRaw) || 0))}`);
  }

  // Stamp with export date
  const now       = new Date();
  const dateStamp = now.toISOString().slice(0, 10);           // YYYY-MM-DD
  const timeStamp = now.toTimeString().slice(0, 5).replace(':', ''); // HHMM
  csvLines.push('');
  csvLines.push(`Exported,${dateStamp} ${timeStamp}`);

  const blob     = new Blob([csvLines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
  const url      = URL.createObjectURL(blob);
  const filename = `${portfolioName.replace(/[^a-z0-9]/gi, '_')}_${dateStamp}.csv`;

  const link     = document.createElement('a');
  link.href      = url;
  link.download  = filename;
  link.style.display = 'none';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  // Release the object URL after a tick so the download has time to start
  setTimeout(() => URL.revokeObjectURL(url), 250);
}

/* ================================================================
   SYNC MODAL  (GitHub Gist settings)
   ================================================================ */

function initSyncModal() {
  document.getElementById('sync-modal-close')
    ?.addEventListener('click', closeSyncModal);
  document.getElementById('sync-modal-backdrop')
    ?.addEventListener('click', e => { if (e.target.id === 'sync-modal-backdrop') closeSyncModal(); });
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && document.getElementById('sync-modal-backdrop')?.classList.contains('active'))
      closeSyncModal();
  });
}

function openSyncModal() {
  const backdrop = document.getElementById('sync-modal-backdrop');
  if (!backdrop) return;
  // Pre-fill PAT field if already stored (masked)
  const patInput = document.getElementById('sync-pat-input');
  if (patInput) patInput.value = gistSync.connected ? '••••••••••••••••••••' : '';
  _refreshSyncModalUI();
  backdrop.classList.add('active');
  if (!gistSync.connected) patInput?.focus();
}

function closeSyncModal() {
  document.getElementById('sync-modal-backdrop')?.classList.remove('active');
  setSyncMsg('');
}

/** Toggle connected/disconnected views inside the modal. */
function _refreshSyncModalUI() {
  const connectedView    = document.getElementById('sync-view-connected');
  const disconnectedView = document.getElementById('sync-view-disconnected');
  const gistIdEl         = document.getElementById('sync-gist-id-display');
  if (gistSync.connected) {
    connectedView?.style.setProperty('display', 'flex');
    disconnectedView?.style.setProperty('display', 'none');
    if (gistIdEl) gistIdEl.textContent = gistSync.gistId || 'Will be created on next save';
  } else {
    connectedView?.style.setProperty('display', 'none');
    disconnectedView?.style.setProperty('display', 'flex');
  }
}

/** Called by the "Connect" button in the sync modal. */
async function connectGist() {
  const patInput = document.getElementById('sync-pat-input');
  const pat      = patInput?.value?.trim();
  if (!pat || pat.startsWith('•')) {
    setSyncMsg('Please paste your Personal Access Token.', 'warn');
    patInput?.focus();
    return;
  }
  const btn = document.getElementById('btn-sync-connect');
  btn.disabled    = true;
  btn.textContent = 'Connecting…';
  setSyncMsg('Validating token…', 'info');
  try {
    await gistSync.connect(pat);
    patInput.value = '••••••••••••••••••••';
    setSyncMsg(gistSync.gistId
      ? 'Connected! Found existing Gist — pulling your data…'
      : 'Connected! A new Gist will be created on your first save.', 'ok');
    _refreshSyncModalUI();
    // If an existing Gist was found, pull immediately and reload the dashboard
    if (gistSync.gistId) {
      const state = await gistSync.pull();
      if (state?.portfolios?.length) {
        PORTFOLIOS    = state.portfolios;
        _portfolioSeq = state.portfolioSeq ?? state.portfolios.length;
        localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
        _gistPushEnabled = true;  // real data loaded — safe to auto-push
        setSyncTimestamp();
        renderDashboard();
        setSyncMsg('Data pulled from Gist and dashboard updated.', 'ok');
      }
    }
  } catch (e) {
    setSyncMsg(`Error: ${e.message}`, 'error');
  } finally {
    btn.disabled    = false;
    btn.textContent = 'Connect';
    updateSyncStatus();
  }
}

/** Called by the "Pull now" button — fetch latest Gist data and reload. */
async function pullFromGist() {
  setSyncMsg('Pulling from Gist…', 'info');
  try {
    const state = await gistSync.pull();
    if (!state?.portfolios?.length) { setSyncMsg('No data found in Gist.', 'warn'); return; }
    PORTFOLIOS    = state.portfolios;
    _portfolioSeq = state.portfolioSeq ?? state.portfolios.length;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    _gistPushEnabled = true;  // real data loaded — safe to auto-push
    setSyncTimestamp();
    renderDashboard();
    setSyncMsg('Dashboard updated from Gist.', 'ok');
  } catch (e) {
    setSyncMsg(`Pull failed: ${e.message}`, 'error');
  }
}

/** Called by the "Push now" button — force-save current state to Gist. */
async function pushToGist() {
  // Build state from DOM (same as saveState does)
  const cards = Array.from(document.querySelectorAll('.portfolio-card'));
  const state = {
    portfolioSeq: _portfolioSeq,
    savedAt: new Date().toISOString(),
    portfolios: cards.map(card => {
      const id       = card.id.replace('card-', '');
      const name     = card.querySelector('.card-title')?.textContent     || '';
      const subtitle = card.querySelector('.card-subtitle')?.textContent  || '';
      const rows     = Array.from(card.querySelectorAll('tbody tr[data-row]'));
      const holdings = rows.map(row => ({
        ticker   : row.querySelector('[data-ticker]')?.value              || '',
        shares   : toNum(row.querySelector('[data-shares]')?.value),
        price    : toNum(row.querySelector('[data-cost-basis]')?.value),
        targetPct: toNum(row.querySelector('[data-target-pct]')?.value),
        mktPrice : toNum(row.querySelector('[data-mkt-price]')?.dataset.raw),
      }));
      return { id, name, subtitle, holdings };
    }),
  };

  // Warn the user if the data currently on screen looks like demo defaults
  if (_looksLikeDemo(state.portfolios)) {
    const ok = confirm(
      'Warning: the dashboard is currently showing demo data.\n\n' +
      'Pushing now will OVERWRITE your real Gist data with demo placeholders.\n\n' +
      'Are you sure you want to continue?'
    );
    if (!ok) { setSyncMsg('Push cancelled.', 'warn'); return; }
  }

  setSyncMsg('Pushing to Gist…', 'info');
  try {
    await gistSync.push(state);
    _gistPushEnabled = true;
    const gistIdEl = document.getElementById('sync-gist-id-display');
    if (gistIdEl) gistIdEl.textContent = gistSync.gistId;
    setSyncMsg('Pushed successfully.', 'ok');
  } catch (e) {
    setSyncMsg(`Push failed: ${e.message}`, 'error');
  }
}

/** Disconnect and clear stored credentials. */
function disconnectGist() {
  if (!confirm('Disconnect cloud sync? Your GitHub token will be removed from this browser. Your Gist data on GitHub will not be deleted.')) return;
  gistSync.disconnect();
  _refreshSyncModalUI();
  setSyncMsg('Disconnected. Data stays in localStorage only.', 'info');
}

/* ================================================================
   FORMATTING UTILITIES
   ================================================================ */

/** Format a number as USD currency: $1,234.56 */
function fmtCurrency(value) {
  return '$' + Number(value).toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

/** Format a number as a percentage with 2 decimal places */
function fmtPct(value) {
  return `${Number(value).toFixed(2)}%`;
}

/** Format a drift value with explicit +/- sign */
function fmtDrift(value) {
  const n    = Number(value);
  const sign = n >= 0 ? '+' : '';
  return `${sign}${n.toFixed(2)}%`;
}

/**
 * Return the CSS class for a drift value:
 *   neutral    — |drift| < 0.5 %
 *   mod-pos    — 0.5 – 3.0 % overweight  (amber warning)
 *   mod-neg    — 0.5 – 3.0 % underweight (blue)
 *   alert-pos  — > 3.0 % overweight      (red — row flagged)
 *   alert-neg  — > 3.0 % underweight     (green — row flagged)
 */
function driftClass(driftPct) {
  const abs = Math.abs(driftPct);
  if (abs < 0.5) return 'drift-neutral';
  if (abs > 3)   return driftPct > 0 ? 'drift-alert-pos' : 'drift-alert-neg';
  return driftPct > 0 ? 'drift-mod-pos' : 'drift-mod-neg';
}

/** Parse a string/number to float, returning 0 for NaN */
function toNum(val) {
  const n = parseFloat(val);
  return isNaN(n) ? 0 : n;
}

/** Escape HTML special characters — used for injecting text into innerHTML */
function escHTML(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Escape for HTML attribute values (double-quote context).
 * Portfolios IDs are safe ASCII but applied for defence-in-depth.
 */
function escAttr(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}
