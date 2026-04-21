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

/* ── In-memory settings (populated from server on startup) ──── */
let _settings = { finnhubKey: '', jsonbinKey: '', jsonbinId: '' };
const getFinnhubKey = () => _settings.finnhubKey || '';

/* ── Price proxy map — Fidelity-only funds with no public API ── */
/*
  FZROX / FZILX / FXAIX etc. are Fidelity Zero-fee mutual funds that
  are NOT listed on any public exchange and have no ticker on Yahoo
  Finance, Finnhub, Polygon, Tiingo, or any CORS-accessible API.
  Fidelity offers no public market-data endpoint for them.

  For each such ticker we map to a **proxy ETF** that tracks the same
  underlying index so Fetch Prices can return a useful price.
  The fetched cell is tinted amber and titled "Price proxied from …".
  Users can also type their own price directly into the Mkt Price cell.

    FZROX → VTI   (both track CRSP US Total Market)
    FZILX → VXUS  (Fidelity Zero International → Vanguard Total Intl)
    FXAIX → SPY   (Fidelity 500 Index → S&P 500 ETF)
*/
const PRICE_PROXY_MAP = {
  FZROX: 'VTI',
  FZILX: 'VXUS',
  FXAIX: 'SPY',
};


/**
 * Collect DOM state and persist to the local server (debounced 500 ms).
 * The save indicator flashes immediately for responsive UI feedback.
 */
let _saveTimer = null;
function saveState() {
  flashSaveIndicator();
  clearTimeout(_saveTimer);
  _saveTimer = setTimeout(_persistState, 500);
}

const _LS_KEY = 'ira-dashboard-state';
const _IS_LOCAL = ['localhost', '127.0.0.1'].includes(location.hostname);

async function _persistState() {
  const cards = Array.from(document.querySelectorAll('.portfolio-card'));
  const state = {
    portfolioSeq: _portfolioSeq,
    settings    : _settings,
    portfolios  : cards.map(card => {
      const id       = card.id.replace('card-', '');
      const name     = card.querySelector('.card-title')?.textContent     || '';
      const subtitle = card.querySelector('.card-subtitle')?.textContent  || '';
      const rows     = Array.from(card.querySelectorAll('tbody tr[data-row]'));
      const holdings = rows.map(row => ({
        ticker   : row.querySelector('[data-ticker]')?.value              || '',
        shares   : toNum(row.querySelector('[data-shares]')?.value),
        price    : toNum(row.querySelector('[data-cost-basis]')?.value),
        targetPct: toNum(row.querySelector('[data-target-pct]')?.value),
        mktPrice : toNum(row.querySelector('[data-mkt-price]')?.value),
      }));
      return { id, name, subtitle, holdings };
    }),
  };

  /* Always persist to localStorage so GitHub Pages retains data. */
  try {
    localStorage.setItem(_LS_KEY, JSON.stringify(state));
  } catch (e) {
    console.warn('localStorage save failed:', e);
  }

  /* Sync to JSONBin when an API key is configured. */
  if (_settings.jsonbinKey) {
    try {
      if (!_settings.jsonbinId) {
        /* First save — auto-create a private bin. */
        const res = await fetch('https://api.jsonbin.io/v3/b', {
          method : 'POST',
          headers: {
            'Content-Type' : 'application/json',
            'X-Master-Key' : _settings.jsonbinKey,
            'X-Bin-Name'   : 'ira-portfolio-dashboard',
            'X-Bin-Private': 'true',
          },
          body: JSON.stringify(state),
        });
        if (res.ok) {
          const json = await res.json();
          _settings.jsonbinId  = json.metadata.id;
          state.settings.jsonbinId = json.metadata.id;
          /* Persist the new binId back to localStorage. */
          try { localStorage.setItem(_LS_KEY, JSON.stringify(state)); } catch (_) {}
          /* Show the bin ID in the Settings modal if it is open. */
          _updateJsonbinIdDisplay();
        }
      } else {
        /* Subsequent saves — update the existing bin. */
        await fetch(`https://api.jsonbin.io/v3/b/${_settings.jsonbinId}`, {
          method : 'PUT',
          headers: {
            'Content-Type': 'application/json',
            'X-Master-Key': _settings.jsonbinKey,
          },
          body: JSON.stringify(state),
        });
      }
    } catch (e) {
      console.warn('JSONBin save failed:', e);
    }
  }

  /* Also persist to the local server when running locally. */
  if (_IS_LOCAL) {
    try {
      await fetch('/api/data', {
        method : 'POST',
        headers: { 'Content-Type': 'application/json' },
        body   : JSON.stringify(state),
      });
    } catch (e) {
      console.warn('Server save failed:', e);
    }
  }
}

/**
 * Load portfolio data.
 * Tries the local server first; falls back to localStorage (GitHub Pages).
 * Returns true if valid saved data was found, false otherwise.
 */
async function loadState() {
  const _defaultSettings = () => ({ finnhubKey: '', jsonbinKey: '', jsonbinId: '' });

  /* 1. Peek at localStorage to get credentials before attempting remote loads. */
  let lsState = null;
  try {
    const raw = localStorage.getItem(_LS_KEY);
    if (raw) lsState = JSON.parse(raw);
  } catch (_) {}

  /* 2. Try local server (only when running locally). */
  if (_IS_LOCAL) {
    try {
      const res = await fetch('/api/data');
      if (res.ok) {
        const state = await res.json();
        if (Array.isArray(state?.portfolios)) {
          PORTFOLIOS    = state.portfolios.length > 0 ? state.portfolios : getDefaultPortfolios();
          _portfolioSeq = state.portfolioSeq ?? PORTFOLIOS.length;
          _settings     = { ..._defaultSettings(), ...state.settings };
          return true;
        }
      }
    } catch (e) {
      /* Server not available — fall through. */
    }
  }

  /* 3. Try JSONBin if credentials are stored in localStorage. */
  const creds = lsState?.settings;
  if (creds?.jsonbinKey && creds?.jsonbinId) {
    try {
      const res = await fetch(
        `https://api.jsonbin.io/v3/b/${creds.jsonbinId}/latest`,
        { headers: { 'X-Master-Key': creds.jsonbinKey } }
      );
      if (res.ok) {
        const json  = await res.json();
        const state = json.record;
        if (Array.isArray(state?.portfolios)) {
          PORTFOLIOS    = state.portfolios.length > 0 ? state.portfolios : getDefaultPortfolios();
          _portfolioSeq = state.portfolioSeq ?? PORTFOLIOS.length;
          _settings     = { ..._defaultSettings(), ...state.settings,
                            jsonbinKey: creds.jsonbinKey, jsonbinId: creds.jsonbinId };
          /* Refresh localStorage with the latest remote data. */
          try { localStorage.setItem(_LS_KEY, JSON.stringify(state)); } catch (_) {}
          return true;
        }
      }
    } catch (e) {
      console.warn('JSONBin load failed:', e);
    }
  }

  /* 4. Fall back to localStorage. */
  if (lsState && Array.isArray(lsState?.portfolios)) {
    PORTFOLIOS    = lsState.portfolios.length > 0 ? lsState.portfolios : getDefaultPortfolios();
    _portfolioSeq = lsState.portfolioSeq ?? PORTFOLIOS.length;
    _settings     = { ..._defaultSettings(), ...lsState.settings };
    return true;
  }
  return false;
}

/**
 * Reset to default demo portfolios and persist the reset to disk.
 * Called by the "Reset Demo" button in the site header.
 */
function resetToDemo() {
  if (!confirm('Reset all portfolios to demo data? All your changes will be lost.')) return;
  PORTFOLIOS    = getDefaultPortfolios();
  _portfolioSeq = PORTFOLIOS.length;
  renderDashboard();
  _persistState().catch(e => console.warn('resetToDemo save failed:', e));
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
document.addEventListener('DOMContentLoaded', async () => {
  await loadState();
  setHeaderDate();
  renderDashboard();
  initModal();
  initSettingsModal();
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
                title="Fetch live market prices — set a free Finnhub API key in Cloud Sync for best results">
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
        <button class="btn btn-tax-layer"
                onclick="generateTaxAwareSuggestion('${escAttr(portfolio.id)}')"
                title="Stage 8 — Tax-Aware Rebalancing Analysis">
          <span aria-hidden="true">&#9650;</span> Tax-Aware Analysis
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
      <input type="number"
             class="tbl-input inp-mkt-price${mktPriceRaw > 0 ? ' mkt-price-live' : ''}"
             data-mkt-price
             value="${escAttr(mktPriceRaw > 0 ? String(mktPriceRaw) : '')}"
             min="0" step="any" placeholder="—"
             onchange="recalculate('${escAttr(portfolioId)}')"
             onblur="recalculate('${escAttr(portfolioId)}')"
             aria-label="Market price (editable)" />
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
    const mktPrice   = toNum(row.querySelector('[data-mkt-price]')?.value);
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
 * Path A (preferred) — Finnhub quote API: direct CORS, no proxy, parallel.
 *   Requires a free Finnhub API key stored via the Cloud Sync modal.
 *   Sign up at finnhub.io — free tier gives 60 calls/min (plenty).
 *
 * Path B (fallback) — proxy cascade when no Finnhub key is set:
 *   1. Yahoo Finance v7 batch (query1, then query2) — 1 request for all.
 *   2. Sequential per-ticker: Stooq CSV then Yahoo v8, 300 ms apart.
 *   Proxy timeouts are kept short (4 s) so failures fail fast.
 *
 * Proxies: allorigins.win → corsproxy.io → codetabs.com
 */
async function fetchPrices(portfolioId) {
  const card = document.getElementById(`card-${portfolioId}`);
  if (!card) return;

  const rows = Array.from(card.querySelectorAll('tbody tr[data-row]'));
  const tickerRowMap = {};
  rows.forEach(row => {
    const t = (row.querySelector('[data-ticker]')?.value || '').trim().toUpperCase();
    if (t) (tickerRowMap[t] = tickerRowMap[t] || []).push(row);
  });

  const uniqueTickers = Object.keys(tickerRowMap);
  if (uniqueTickers.length === 0) return;

  // Build the actual fetch list with proxy substitution:
  // FZROX, FZILX, etc. have no public API — substitute the proxy ETF for fetching,
  // then map the price back to the original ticker before applying to DOM.
  const fetchTickers = [...new Set(uniqueTickers.map(t => PRICE_PROXY_MAP[t] || t))];

  /* ── Loading state ─────────────────────────────────────── */
  const btn = card.querySelector(`[data-fetch-btn="${portfolioId}"]`);
  const originalHTML = btn ? btn.innerHTML : '';
  if (btn) {
    btn.disabled = true;
    btn.innerHTML = `<span aria-hidden="true" class="spin-icon">⟳</span> Fetching…`;
  }

  const priceMap = {};
  const finnhubKey = getFinnhubKey();

  /* ── Path A: Finnhub (direct CORS, parallel, no proxy needed) ── */
  if (finnhubKey) {
    const results = await Promise.allSettled(
      fetchTickers.map(async ticker => {
        const ctrl  = new AbortController();
        const timer = setTimeout(() => ctrl.abort(), 8000);
        try {
          const res = await fetch(
            `https://finnhub.io/api/v1/quote?symbol=${encodeURIComponent(ticker)}&token=${encodeURIComponent(finnhubKey)}`,
            { signal: ctrl.signal },
          );
          clearTimeout(timer);
          if (!res.ok) throw new Error(`HTTP ${res.status}`);
          const data = await res.json();
          /* c = current price; 0 means unknown symbol or closed/no data */
          if (data?.c > 0) return { ticker, price: data.c };
          throw new Error('No valid price from Finnhub (c=0)');
        } catch (e) {
          clearTimeout(timer);
          throw e;
        }
      }),
    );
    results.forEach((outcome, idx) => {
      if (outcome.status === 'fulfilled')
        priceMap[outcome.value.ticker] = outcome.value.price;
      else
        console.warn(`Finnhub — ${fetchTickers[idx]}:`, outcome.reason?.message);
    });
  }

  /* ── Path B: Proxy fallback for any ticker still missing ── */
  const proxyWrap = [
    url => `https://api.allorigins.win/raw?url=${encodeURIComponent(url)}`,
    url => `https://corsproxy.io/?${encodeURIComponent(url)}`,
    url => `https://api.codetabs.com/v1/proxy?quest=${encodeURIComponent(url)}`,
  ];

  /* Short timeout — fail fast so we don't hang for minutes */
  async function tryProxies(targetUrl, parseMode, ms = 4000) {
    for (const makeProxy of proxyWrap) {
      const ctrl  = new AbortController();
      const timer = setTimeout(() => ctrl.abort(), ms);
      try {
        const res = await fetch(makeProxy(targetUrl), { signal: ctrl.signal });
        clearTimeout(timer);
        if (!res.ok) continue;
        return parseMode === 'json' ? await res.json() : await res.text();
      } catch (e) {
        clearTimeout(timer);
      }
    }
    return null;
  }

  const missingAfterFinnhub = fetchTickers.filter(t => priceMap[t] == null);
  if (missingAfterFinnhub.length > 0) {

    /* B1: Yahoo v7 batch — query1 then query2 */
    for (const host of ['query1', 'query2']) {
      const still = missingAfterFinnhub.filter(t => priceMap[t] == null);
      if (still.length === 0) break;
      const symList = still.map(encodeURIComponent).join(',');
      const batch   = await tryProxies(
        `https://${host}.finance.yahoo.com/v7/finance/quote?symbols=${symList}&fields=regularMarketPrice`,
        'json', 6000,
      );
      (batch?.quoteResponse?.result || []).forEach(q => {
        if (q?.regularMarketPrice != null)
          priceMap[q.symbol.toUpperCase()] = q.regularMarketPrice;
      });
    }

    /* B2: Sequential per-ticker (Stooq → Yahoo v8) for anything still missing */
    const delay = ms => new Promise(r => setTimeout(r, ms));
    for (const ticker of missingAfterFinnhub.filter(t => priceMap[t] == null)) {
      let resolved = false;

      /* Try Stooq CSV: .US suffix first, then plain ticker */
      for (const sym of [`${ticker}.US`, ticker]) {
        const text = await tryProxies(
          `https://stooq.com/q/l/?s=${encodeURIComponent(sym)}&f=sd2t2ohlcv&e=csv`,
          'text', 4000,
        );
        if (text) {
          const fields = text.trim().split('\n').pop().split(',');
          const close  = parseFloat(fields[6]);
          if (!isNaN(close) && close > 0) { priceMap[ticker] = close; resolved = true; break; }
        }
      }

      /* Last resort: Yahoo v8 chart per-ticker */
      if (!resolved) {
        const yData = await tryProxies(
          `https://query2.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(ticker)}` +
          `?interval=1d&range=1d&includePrePost=false`,
          'json', 4000,
        );
        const price = yData?.chart?.result?.[0]?.meta?.regularMarketPrice;
        if (price != null) priceMap[ticker] = price;
        else console.warn(`fetchPrices — ${ticker}: all sources failed`);
      }

      await delay(300);
    }
  }

  // Fill proxied prices back to the original tickers
  // e.g. priceMap['FZROX'] = priceMap['VTI'] after fetching 'VTI'
  uniqueTickers.forEach(t => {
    if (PRICE_PROXY_MAP[t] != null && priceMap[t] == null) {
      const proxyTicker = PRICE_PROXY_MAP[t];
      if (priceMap[proxyTicker] != null) priceMap[t] = priceMap[proxyTicker];
    }
  });

  /* ── Apply prices to DOM ───────────────────────────────── */
  let fetched = 0;
  uniqueTickers.forEach(ticker => {
    const price       = priceMap[ticker];
    const proxySource = PRICE_PROXY_MAP[ticker];   // e.g. 'VTI' for FZROX
    const isProxied   = proxySource != null && price != null;
    (tickerRowMap[ticker] || []).forEach(row => {
      const mktEl = row.querySelector('[data-mkt-price]');
      if (!mktEl) return;
      if (price != null) {
        mktEl.value     = price;
        mktEl.className = `tbl-input inp-mkt-price mkt-price-live${isProxied ? ' mkt-price-proxied' : ''}`;
        mktEl.title     = isProxied
          ? `Price proxied from ${proxySource} — ${ticker} has no public market data`
          : '';
        fetched++;
      } else {
        mktEl.value     = '';
        mktEl.className = 'tbl-input inp-mkt-price mkt-price-na';
        mktEl.title     = '';
      }
    });
  });

  recalculate(portfolioId);

  /* ── Button feedback ───────────────────────────────────── */
  if (btn) {
    btn.innerHTML = fetched === 0
      ? `<span aria-hidden="true">⚠</span> No prices — set Finnhub key in Cloud Sync`
      : `<span aria-hidden="true">✓</span> Updated (${fetched}/${uniqueTickers.length})`;
    setTimeout(() => {
      btn.innerHTML = originalHTML;
      btn.disabled  = false;
    }, 4000);
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
    const mktRaw      = toNum(row.querySelector('[data-mkt-price]')?.value);
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
   SETTINGS MODAL  (Finnhub API key)
   ================================================================ */

function initSettingsModal() {
  document.getElementById('settings-modal-close')
    ?.addEventListener('click', closeSettingsModal);
  document.getElementById('settings-modal-backdrop')
    ?.addEventListener('click', e => {
      if (e.target.id === 'settings-modal-backdrop') closeSettingsModal();
    });
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' &&
        document.getElementById('settings-modal-backdrop')?.classList.contains('active'))
      closeSettingsModal();
  });
}

function openSettingsModal() {
  const backdrop = document.getElementById('settings-modal-backdrop');
  if (!backdrop) return;
  const fhInput  = document.getElementById('finnhub-key-input');
  const fhStatus = document.getElementById('finnhub-key-status');
  if (fhInput)  fhInput.value        = getFinnhubKey() ? '••••••••••••••••••••••••••••••••' : '';
  if (fhStatus) fhStatus.textContent = getFinnhubKey() ? '✓ Key saved — Fetch Prices will use Finnhub.' : '';
  /* Populate JSONBin fields. */
  const jbInput  = document.getElementById('jsonbin-key-input');
  const jbStatus = document.getElementById('jsonbin-key-status');
  if (jbInput)  jbInput.value        = _settings.jsonbinKey ? '••••••••••••••••••••••••••••••••' : '';
  if (jbStatus) jbStatus.textContent = _settings.jsonbinKey ? '✓ Key saved — data syncs to JSONBin.' : '';
  _updateJsonbinIdDisplay();
  backdrop.classList.add('active');
  if (!getFinnhubKey()) fhInput?.focus();
}

function closeSettingsModal() {
  document.getElementById('settings-modal-backdrop')?.classList.remove('active');
}

/** Save or clear the Finnhub API key; persists to server via saveState(). */
function saveFinnhubKey() {
  const input    = document.getElementById('finnhub-key-input');
  const statusEl = document.getElementById('finnhub-key-status');
  const val      = input?.value?.trim() ?? '';
  if (!val || val.startsWith('•')) {
    if (statusEl) statusEl.textContent = 'No change — paste a new key to update.';
    return;
  }
  if (val === 'clear' || val === 'remove') {
    _settings.finnhubKey = '';
    if (input)    input.value = '';
    if (statusEl) statusEl.textContent = 'Key removed.';
    saveState();
    return;
  }
  _settings.finnhubKey = val;
  if (input)    input.value = '••••••••••••••••••••••••••••••••';
  if (statusEl) statusEl.textContent = '✓ Key saved — Fetch Prices will now use Finnhub.';
  saveState();
}

/** Show or hide the Bin ID display row in the Settings modal. */
function _updateJsonbinIdDisplay() {
  const input = document.getElementById('jsonbin-id-input');
  if (!input) return;
  if (_settings.jsonbinId) {
    input.value = _settings.jsonbinId;
  }
}

/** Save or clear the JSONBin X-Master-Key; triggers a save to create the bin if needed. */
function saveJsonbinKey() {
  const input    = document.getElementById('jsonbin-key-input');
  const idInput  = document.getElementById('jsonbin-id-input');
  const statusEl = document.getElementById('jsonbin-key-status');
  const val      = input?.value?.trim() ?? '';
  if (!val || val.startsWith('•')) {
    /* No new key — but still allow updating the Bin ID alone. */
    const newId = idInput?.value?.trim() ?? '';
    if (newId && newId !== _settings.jsonbinId) {
      _settings.jsonbinId = newId;
      if (statusEl) statusEl.textContent = '↺ Bin ID updated — loading your data…';
      _loadFromJsonbin(statusEl);
    } else {
      if (statusEl) statusEl.textContent = 'No change — paste a new key to update.';
    }
    return;
  }
  if (val === 'clear' || val === 'remove') {
    _settings.jsonbinKey = '';
    _settings.jsonbinId  = '';
    if (input)    input.value = '';
    if (idInput)  idInput.value = '';
    if (statusEl) statusEl.textContent = 'JSONBin key removed — using localStorage only.';
    saveState();
    return;
  }
  _settings.jsonbinKey = val;
  if (input) input.value = '••••••••••••••••••••••••••••••••';

  /* If the user also provided an existing Bin ID, use it to load their data. */
  const providedId = idInput?.value?.trim() ?? '';
  if (providedId) {
    _settings.jsonbinId = providedId;
    if (statusEl) statusEl.textContent = '↺ Key saved — loading data from your existing bin…';
    _loadFromJsonbin(statusEl);
    return;
  }

  if (statusEl) statusEl.textContent = '✓ Key saved — syncing to JSONBin…';
  /* No Bin ID provided — trigger a save so a new bin is created. */
  _persistState().then(() => {
    if (statusEl) statusEl.textContent = _settings.jsonbinId
      ? `✓ Connected — Bin ID: ${_settings.jsonbinId}`
      : '✓ Key saved — bin will be created on next save.';
  }).catch(() => {
    if (statusEl) statusEl.textContent = '⚠ Key saved but JSONBin sync failed — check the key.';
  });
}

/**
 * Fetch data from the configured JSONBin bin and replace the current portfolios.
 * Called when the user provides an existing Bin ID on a new device.
 *
 * @param {HTMLElement|null} statusEl  Status message element to update.
 */
async function _loadFromJsonbin(statusEl) {
  try {
    const res = await fetch(
      `https://api.jsonbin.io/v3/b/${_settings.jsonbinId}/latest`,
      { headers: { 'X-Master-Key': _settings.jsonbinKey } }
    );
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json  = await res.json();
    const state = json.record;
    if (!Array.isArray(state?.portfolios)) throw new Error('Unexpected bin format.');
    PORTFOLIOS    = state.portfolios.length > 0 ? state.portfolios : getDefaultPortfolios();
    _portfolioSeq = state.portfolioSeq ?? PORTFOLIOS.length;
    /* Keep current credentials; merge everything else from the bin. */
    _settings = {
      ..._settings,
      finnhubKey: state.settings?.finnhubKey || _settings.finnhubKey,
    };
    /* Persist locally so next page-load skips this step. */
    try { localStorage.setItem(_LS_KEY, JSON.stringify({ ...state, settings: _settings })); } catch (_) {}
    _updateJsonbinIdDisplay();
    renderDashboard();
    if (statusEl) statusEl.textContent = `✓ Data loaded — Bin ID: ${_settings.jsonbinId}`;
  } catch (e) {
    console.warn('_loadFromJsonbin failed:', e);
    if (statusEl) statusEl.textContent = `⚠ Could not load bin — check the Bin ID and key. (${e.message})`;
  }
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
