/* ================================================================
   Stage 8 — Tax-Aware Rebalancing Layer
   ================================================================
   Self-contained plug-in that sits AFTER the Stage 7 rebalancing
   engine.  It accepts the Stage 7 action list and returns a
   tax-aware version with adjusted recommendations, tax scores,
   and plain-English explanations.

   This file does NOT touch any Stage 7 logic.

   Public API (called from script.js UI):
     generateTaxAwareSuggestion(portfolioId)

   Core function (pure, testable):
     applyTaxLayer(rebalanceActions, portfolio, cashFlows?)
       → { actions[], warnings[], accountType }
   ================================================================ */

'use strict';

/* ================================================================
   ACCOUNT TYPE CLASSIFICATION
   ================================================================ */

/**
 * Infer the tax treatment of a portfolio from its name/subtitle.
 * Returns one of: 'roth_ira' | 'trad_ira' | 'k401' | 'taxable'
 *
 * Roth and Traditional IRA accounts are tax-advantaged:
 *   selling inside them triggers zero capital-gains tax.
 * Taxable (brokerage) accounts ARE subject to capital-gains tax.
 */
function classifyAccountType(portfolio) {
  const label = ((portfolio.name || '') + ' ' + (portfolio.subtitle || '')).toLowerCase();
  if (label.includes('roth'))                          return 'roth_ira';
  if (label.includes('traditional') ||
      label.includes('trad-ira')    ||
      label.includes('trad ira'))                      return 'trad_ira';
  if (label.includes('401k') || label.includes('401(k)')) return 'k401';
  return 'taxable';
}

/** Human-readable label for each account type. */
const ACCOUNT_TYPE_LABELS = {
  roth_ira:  'Roth IRA (tax-free growth)',
  trad_ira:  'Traditional IRA (tax-deferred)',
  k401:      '401(k) (tax-deferred)',
  taxable:   'Taxable Brokerage Account',
};

/* ================================================================
   TAX RATES
   ================================================================ */

/**
 * Capital-gains tax rates by account type and holding period.
 *
 * IRA / Roth / 401(k): no capital-gains tax on trades inside the account.
 * Taxable: assumes a typical moderate-income household filing jointly
 *   (15% long-term, 22% short-term ordinary income rate).
 * Families with different income levels can update these two numbers.
 */
const TAX_RATES = {
  roth_ira : { short_term: 0,    long_term: 0    },
  trad_ira : { short_term: 0,    long_term: 0    },
  k401     : { short_term: 0,    long_term: 0    },
  taxable  : { short_term: 0.22, long_term: 0.15 },
};

/* ================================================================
   HOLDING-PERIOD ESTIMATION
   ================================================================ */

/**
 * Estimate whether a gain is short-term (< 1 year) or long-term (≥ 1 year).
 *
 * If the holding has a purchaseDate field, we use the actual elapsed time.
 * Otherwise we fall back to a conservative heuristic:
 *   unrealized gain > 15% of cost  →  likely held more than a year
 *   (appreciated positions tend to be older; this errs on the side of calm)
 */
function estimateHoldingPeriod(holding) {
  if (holding.purchaseDate) {
    const days = (Date.now() - new Date(holding.purchaseDate).getTime()) / 86400000;
    return days >= 365 ? 'long_term' : 'short_term';
  }
  const cost = holding.price || 0;
  const mkt  = holding.mktPrice > 0 ? holding.mktPrice : cost;
  if (cost > 0 && mkt > cost && ((mkt - cost) / cost) > 0.15) {
    return 'long_term';
  }
  return 'long_term'; // default: assume long-term (less alarming, more realistic)
}

/* ================================================================
   TAX-COST COMPUTATION
   ================================================================ */

/**
 * Compute the estimated tax cost of selling a position.
 *
 * @param {Object} holding      — { ticker, shares, price (cost basis), mktPrice }
 * @param {string} accountType  — one of the keys in TAX_RATES
 * @returns {{ unrealizedGain, taxCost, taxCostPct, holdingPeriod }}
 */
function computeTaxCost(holding, accountType) {
  const rates        = TAX_RATES[accountType] ?? TAX_RATES.taxable;
  const costBasis    = holding.price   || 0;
  const mktPrice     = holding.mktPrice > 0 ? holding.mktPrice : costBasis;
  const shares       = holding.shares  || 0;

  const currentValue   = shares * mktPrice;
  const unrealizedGain = (mktPrice - costBasis) * shares;

  if (unrealizedGain <= 0 || currentValue <= 0) {
    return { unrealizedGain, taxCost: 0, taxCostPct: 0, holdingPeriod: 'n/a' };
  }

  const holdingPeriod = estimateHoldingPeriod(holding);
  const rate          = rates[holdingPeriod] ?? 0;
  const taxCost       = unrealizedGain * rate;
  const taxCostPct    = (taxCost / currentValue) * 100;

  return { unrealizedGain, taxCost, taxCostPct, holdingPeriod };
}

/* ================================================================
   TAX COST SCORE
   ================================================================ */

/**
 * Assign a Tax Cost Score based on what percentage of position
 * value would be lost to taxes if this position were sold today.
 *
 *   LOW      < 1%   — negligible; sell freely
 *   MODERATE 1–3%   — worth considering; sell only if drift is large
 *   HIGH     > 3%   — significant; prefer tax-efficient alternatives
 */
function scoreTaxCost(taxCostPct) {
  if (taxCostPct < 1)  return 'low';
  if (taxCostPct <= 3) return 'moderate';
  return 'high';
}

/* ================================================================
   DRIFT SEVERITY
   ================================================================ */

/**
 * Classify how far off-target a position is.
 *
 *   trivial  < 1 pp  — rounding noise
 *   minor    1–3 pp  — worth monitoring
 *   moderate 3–7 pp  — should be corrected
 *   extreme  > 7 pp  — must be corrected
 */
function scoreDrift(driftPct) {
  const abs = Math.abs(driftPct);
  if (abs < 1) return 'trivial';
  if (abs < 3) return 'minor';
  if (abs < 7) return 'moderate';
  return 'extreme';
}

/* ================================================================
   ACTION REWRITING (CORE DECISION LOGIC)
   ================================================================ */

/**
 * Convert a single Stage 7 action into a tax-aware action.
 *
 * Decision rules:
 *  1. Tax-advantaged account   → allow all trades; tax_cost_score = 'none'
 *  2. Taxable + buy/hold       → allow; tax_cost_score = 'none'
 *  3. Taxable + sell + LOW tax → allow; tax_cost_score = 'low'
 *  4. Taxable + sell + MOD tax + large drift → allow partial; tax_cost_score = 'moderate'
 *  5. Taxable + sell + MOD tax + small drift → redirect_cash_flows; tax_cost_score = 'moderate'
 *  6. Taxable + sell + HIGH tax + extreme drift → partial sell (50%); tax_cost_score = 'high'
 *  7. Taxable + sell + HIGH tax + non-extreme   → avoid_selling; tax_cost_score = 'high'
 */
function buildTaxAwareAction(suggestedTrade, holding, accountType, driftPct) {

  /* ── Tax-advantaged accounts: zero tax on any trade ── */
  if (accountType !== 'taxable') {
    const typeLabel = ACCOUNT_TYPE_LABELS[accountType] || accountType;
    let reason;
    switch (suggestedTrade.type) {
      case 'sell':
        reason = `This is a ${typeLabel} account — selling here has no capital-gains impact. Feel free to trim as suggested.`;
        break;
      case 'buy':
        reason = `No tax cost on purchases in a ${typeLabel}. Buy to bring this position back to target.`;
        break;
      default:
        reason = 'Position is within target range — no action needed right now.';
    }
    return {
      account_id     : holding.portfolioId || '',
      ticker         : holding.ticker      || '',
      action         : suggestedTrade.type,
      approx_amount  : suggestedTrade.amount || 0,
      tax_cost_score : 'none',
      reason,
    };
  }

  /* ── Taxable account: buy or hold — no tax ── */
  if (suggestedTrade.type === 'buy') {
    return {
      account_id     : holding.portfolioId || '',
      ticker         : holding.ticker      || '',
      action         : 'buy',
      approx_amount  : suggestedTrade.amount || 0,
      tax_cost_score : 'none',
      reason         : 'Buying never triggers capital gains. Go ahead and purchase to close the gap.',
    };
  }

  if (suggestedTrade.type === 'hold') {
    return {
      account_id     : holding.portfolioId || '',
      ticker         : holding.ticker      || '',
      action         : 'hold',
      approx_amount  : 0,
      tax_cost_score : 'none',
      reason         : 'Already close to target — no trade needed.',
    };
  }

  /* ── Taxable account: SELL — evaluate tax cost ── */
  const { unrealizedGain, taxCost, taxCostPct, holdingPeriod } =
    computeTaxCost(holding, accountType);

  const taxScore  = scoreTaxCost(taxCostPct);
  const driftScore = scoreDrift(driftPct);
  const periodLabel = holdingPeriod === 'long_term' ? 'long-term' : 'short-term';
  const gainStr = fmtCurrency(Math.max(0, unrealizedGain));
  const costStr = `~${taxCostPct.toFixed(1)}% of position value`;

  /* LOW tax cost — sell freely */
  if (taxScore === 'low') {
    return {
      account_id     : holding.portfolioId || '',
      ticker         : holding.ticker      || '',
      action         : 'sell',
      approx_amount  : suggestedTrade.amount || 0,
      tax_cost_score : 'low',
      reason         : `Estimated tax cost is small (${costStr}). Selling is a good choice here — the rebalancing benefit outweighs the tax drag.`,
    };
  }

  /* MODERATE tax cost */
  if (taxScore === 'moderate') {
    if (driftScore === 'moderate' || driftScore === 'extreme') {
      return {
        account_id     : holding.portfolioId || '',
        ticker         : holding.ticker      || '',
        action         : 'sell',
        approx_amount  : suggestedTrade.amount || 0,
        tax_cost_score : 'moderate',
        reason         : `The drift on ${holding.ticker} is large enough that trimming makes sense even with a moderate tax cost (${costStr}, ${periodLabel} gain of ${gainStr}). A partial trim is fine — you don't need to sell everything at once.`,
      };
    }
    /* Small drift + moderate tax — use cash flows instead */
    return {
      account_id     : holding.portfolioId || '',
      ticker         : holding.ticker      || '',
      action         : 'redirect_cash_flows',
      approx_amount  : suggestedTrade.amount || 0,
      tax_cost_score : 'moderate',
      reason         : `The drift is modest and the tax cost is moderate (${costStr}). Rather than selling ${holding.ticker}, simply redirect your next contributions or dividends toward your underweight positions until the balance naturally returns.`,
    };
  }

  /* HIGH tax cost */
  if (driftScore === 'extreme') {
    /* Extreme drift forces a partial trim despite high taxes */
    const partialAmount = (suggestedTrade.amount || 0) * 0.5;
    return {
      account_id     : holding.portfolioId || '',
      ticker         : holding.ticker      || '',
      action         : 'sell',
      approx_amount  : partialAmount,
      tax_cost_score : 'high',
      reason         : `The drift on ${holding.ticker} is extreme, so a small trim is still warranted — but the tax cost is high (${costStr}, ${periodLabel} gain of ${gainStr}). Consider selling only about half the suggested amount to limit the tax hit. Spread the rest over future contributions.`,
    };
  }

  /* Non-extreme drift + high tax → avoid selling entirely */
  return {
    account_id     : holding.portfolioId || '',
    ticker         : holding.ticker      || '',
    action         : 'avoid_selling',
    approx_amount  : 0,
    tax_cost_score : 'high',
    reason         : `Selling ${holding.ticker} right now would cost an estimated ${costStr} in taxes (${periodLabel} gain of ${gainStr}). The drift is manageable — skip the sale and redirect incoming cash to your underweight positions instead.`,
  };
}

/* ================================================================
   WARNING GENERATOR
   ================================================================ */

/**
 * Build a list of plain-English warnings from the final action set.
 */
function generateWarnings(actions, meta) {
  const warnings = [];

  const highTaxTickers   = actions.filter(a => a.tax_cost_score === 'high').map(a => a.ticker);
  const avoidedTickers   = actions.filter(a => a.action === 'avoid_selling').map(a => a.ticker);
  const redirectTickers  = actions.filter(a => a.action === 'redirect_cash_flows').map(a => a.ticker);

  if (highTaxTickers.length > 0) {
    warnings.push({
      type    : 'tax',
      message : `${listTickers(highTaxTickers)} ${highTaxTickers.length === 1 ? 'has' : 'have'} a high estimated tax cost in a taxable account. Selling now would meaningfully reduce your net proceeds — consider the alternatives listed above.`,
    });
  }

  if (avoidedTickers.length > 0) {
    warnings.push({
      type    : 'recommendation',
      message : `For ${listTickers(avoidedTickers)}: skip the sale and use new contributions or dividend payments to buy underweight positions instead. This achieves the same rebalancing result with no tax bill.`,
    });
  }

  if (redirectTickers.length > 0) {
    warnings.push({
      type    : 'recommendation',
      message : `For ${listTickers(redirectTickers)}: redirect your next contribution or dividend away from these positions so they naturally drift back toward target over time.`,
    });
  }

  if (meta?.hasIRACapacity) {
    warnings.push({
      type    : 'opportunity',
      message : `This portfolio is an IRA or Roth IRA. Selling inside a tax-advantaged account has zero capital-gains impact — rebalance here freely before touching any taxable accounts you may hold.`,
    });
  }

  return warnings;
}

/** Format a ticker list as a readable string: "VTI, QQQ, and GLD" */
function listTickers(tickers) {
  if (tickers.length === 0) return '';
  if (tickers.length === 1) return tickers[0];
  if (tickers.length === 2) return `${tickers[0]} and ${tickers[1]}`;
  return tickers.slice(0, -1).join(', ') + ', and ' + tickers[tickers.length - 1];
}

/* ================================================================
   STAGE 8 — MAIN API
   apply_tax_layer(rebalanceActions, portfolio, cashFlows?)
   ================================================================ */

/**
 * Apply tax-aware logic to Stage 7 rebalancing suggestions.
 *
 * @param {Object[]} rebalanceActions
 *   Stage 7 output — each item: { ticker, type, amount, driftPct }
 *   where type = 'buy' | 'sell' | 'hold'
 *
 * @param {Object} portfolio
 *   Full portfolio object: { id, name, subtitle, holdings[] }
 *   holdings items: { ticker, shares, price (cost basis), mktPrice, targetPct }
 *
 * @param {Object} [cashFlows]
 *   Optional: { newContributions: number, dividends: number }
 *   Dollar amounts of incoming cash available to direct at underweight positions.
 *
 * @returns {{ actions: Object[], warnings: Object[], accountType: string }}
 */
function applyTaxLayer(rebalanceActions, portfolio, cashFlows = {}) {
  const accountType = classifyAccountType(portfolio);

  /* Index holdings by ticker for fast lookup */
  const holdingMap = {};
  (portfolio.holdings || []).forEach(h => {
    holdingMap[(h.ticker || '').toUpperCase()] = { ...h, portfolioId: portfolio.id };
  });

  /* Rewrite each Stage 7 action */
  const actions = (rebalanceActions || []).map(suggested => {
    const holding = holdingMap[(suggested.ticker || '').toUpperCase()] || {
      ticker     : suggested.ticker,
      portfolioId: portfolio.id,
      shares     : 0,
      price      : 0,
      mktPrice   : 0,
    };
    return buildTaxAwareAction(suggested, holding, accountType, suggested.driftPct || 0);
  });

  const hasIRACapacity = accountType === 'roth_ira' || accountType === 'trad_ira' || accountType === 'k401';
  const warnings       = generateWarnings(actions, { hasIRACapacity });

  /* Prepend a cash-flow advisory when incoming money is available */
  const totalCash = (cashFlows.newContributions || 0) + (cashFlows.dividends || 0);
  if (totalCash > 0) {
    warnings.unshift({
      type    : 'cash_flow',
      message : `You have ${fmtCurrency(totalCash)} in available cash (contributions + dividends). Direct this toward underweight positions first — it may be all you need to rebalance without selling anything.`,
    });
  }

  return { actions, warnings, accountType };
}

/* ================================================================
   REPORT FORMATTER
   ================================================================ */

/**
 * Format the Stage 8 output into a clean, human-readable text report.
 */
function formatTaxAwareReport(portfolio, totalValue, actions, warnings, accountType) {
  const dateStr = new Date().toLocaleString('en-US', {
    month: 'long', day: 'numeric', year: 'numeric',
    hour: '2-digit', minute: '2-digit',
  });

  const typeLabel  = ACCOUNT_TYPE_LABELS[accountType] || accountType;

  /* Action label display */
  const ACTION_DISPLAY = {
    sell              : 'Trim (Sell)',
    buy               : 'Buy',
    hold              : 'Hold',
    redirect_cash_flows: 'Redirect Cash',
    avoid_selling     : 'Skip — Use Cash',
  };

  /* Tax score display */
  const SCORE_DISPLAY = {
    none    : '—',
    low     : 'LOW',
    moderate: 'MODERATE',
    high    : 'HIGH',
  };

  /* ── Header ── */
  const BORDER = '═'.repeat(72);
  const lines  = [
    BORDER,
    `  TAX-AWARE REBALANCING ANALYSIS — ${portfolio.name.toUpperCase()}`,
    `  Stage 8: Tax-Aware Rebalancing Layer`,
    `  Account Type : ${typeLabel}`,
    `  Portfolio Value: ${fmtCurrency(totalValue)}`,
    `  Generated    : ${dateStr}`,
    BORDER,
    '',
  ];

  /* ── Action table ── */
  const W = { ticker: 8, action: 18, amount: 14, score: 12 };
  const colHeader = [
    'Ticker'.padEnd(W.ticker),
    'Recommendation'.padEnd(W.action),
    'Est. Amount'.padStart(W.amount),
    'Tax Score'.padStart(W.score),
  ].join('  ');

  const rowSep = '─'.repeat(colHeader.length);

  lines.push('  RECOMMENDED ACTIONS');
  lines.push(`  ${rowSep}`);
  lines.push(`  ${colHeader}`);
  lines.push(`  ${rowSep}`);

  actions.forEach(a => {
    const actionLabel = ACTION_DISPLAY[a.action] || a.action;
    const amountStr   = a.approx_amount > 0 ? fmtCurrency(a.approx_amount) : '—';
    const scoreStr    = SCORE_DISPLAY[a.tax_cost_score] || a.tax_cost_score;

    lines.push([
      `  ${(a.ticker || '').padEnd(W.ticker)}`,
      actionLabel.padEnd(W.action),
      amountStr.padStart(W.amount),
      scoreStr.padStart(W.score),
    ].join('  '));
  });

  lines.push(`  ${rowSep}`);
  lines.push('');

  /* ── Per-action explanations ── */
  lines.push('  WHY THESE RECOMMENDATIONS?');
  lines.push('');
  actions.forEach(a => {
    lines.push(`  ${(a.ticker || '').padEnd(6)}  ${a.reason}`);
    lines.push('');
  });

  /* ── Warnings / Opportunities ── */
  if (warnings.length > 0) {
    lines.push(`  ${'─'.repeat(68)}`);
    lines.push('  NOTES & OPPORTUNITIES');
    lines.push('');
    warnings.forEach((w, i) => {
      const prefix = w.type === 'tax'          ? '⚠  TAX ALERT   '
                   : w.type === 'opportunity'  ? '✦  OPPORTUNITY  '
                   : w.type === 'cash_flow'    ? '💰  CASH FLOW   '
                   : '→  NOTE         ';
      lines.push(`  ${i + 1}. ${prefix}`);
      /* Word-wrap the message at ~65 chars */
      const words   = w.message.split(' ');
      let   lineOut = '     ';
      words.forEach(word => {
        if (lineOut.length + word.length + 1 > 68) {
          lines.push(lineOut.trimEnd());
          lineOut = '     ' + word + ' ';
        } else {
          lineOut += word + ' ';
        }
      });
      if (lineOut.trim()) lines.push(lineOut.trimEnd());
      lines.push('');
    });
  }

  /* ── Tax score legend ── */
  lines.push(`  ${'─'.repeat(68)}`);
  lines.push('  TAX SCORE GUIDE');
  lines.push('');
  lines.push('  LOW      (< 1% of position value)  — tax cost is negligible; sell freely.');
  lines.push('  MODERATE (1–3% of position value)  — weigh tax cost vs. drift severity.');
  lines.push('  HIGH     (> 3% of position value)  — prefer tax-free alternatives when possible.');
  lines.push('  NONE                               — tax-advantaged account; no capital-gains impact.');
  lines.push('');
  lines.push(`  Note: Tax estimates assume ~15% long-term / ~22% short-term rates for taxable`);
  lines.push(`  accounts. Adjust in stage8_tax_layer.js → TAX_RATES if your rates differ.`);
  lines.push('');
  lines.push(BORDER);

  return lines.join('\n');
}

/* ================================================================
   UI INTEGRATION — called by the "Tax-Aware Analysis" button
   ================================================================ */

/**
 * Main entry point from the dashboard UI.
 *
 * Pipeline:
 *   Stage 7 → rebalancing actions (derived from live DOM state)
 *   Stage 8 → tax-aware actions  (via applyTaxLayer)
 *   Output  → formatted report shown in the existing modal
 *
 * Stage 7 logic in script.js is NOT modified — we read the same
 * computed DOM attributes that Stage 7 already produced.
 */
function generateTaxAwareSuggestion(portfolioId) {
  /* Ensure computations are fresh */
  if (typeof recalculate === 'function') recalculate(portfolioId);

  const card      = document.getElementById(`card-${portfolioId}`);
  const portfolio = (typeof PORTFOLIOS !== 'undefined' ? PORTFOLIOS : [])
    .find(p => p.id === portfolioId);
  if (!card || !portfolio) return;

  const rows = Array.from(card.querySelectorAll('tbody tr[data-row]'));
  if (rows.length === 0) {
    if (typeof showModal === 'function') {
      showModal(`Tax-Aware Analysis — ${portfolio.name}`, '  No holdings to analyse.');
    }
    return;
  }

  /* ── Total portfolio value ── */
  const totalValue = rows.reduce(
    (s, r) => s + (parseFloat(r.dataset.computedCurrentValue) || 0), 0,
  );

  /* ── Build Stage 7 action list from DOM (mirrors calcSuggestedTrade logic) ── */
  const stage7Actions = rows.map(row => {
    const ticker    = (row.querySelector('[data-ticker]')?.value    || '').toUpperCase().trim();
    const driftPct  = parseFloat(row.dataset.computedDriftPct)      || 0;
    const curVal    = parseFloat(row.dataset.computedCurrentValue)  || 0;
    const targetPct = parseFloat(row.querySelector('[data-target-pct]')?.value) || 0;

    const targetValue = (targetPct / 100) * totalValue;
    const dollarDiff  = targetValue - curVal;
    const absDrift    = Math.abs(driftPct);

    let type   = 'hold';
    let amount = 0;

    if (absDrift >= 0.1 && totalValue > 0) {
      type   = dollarDiff > 0 ? 'buy' : 'sell';
      amount = Math.abs(dollarDiff);
    }

    return { ticker, type, amount, driftPct };
  });

  /* ── Apply Stage 8 tax layer ── */
  const { actions, warnings, accountType } = applyTaxLayer(stage7Actions, portfolio);

  /* ── Format and display ── */
  const reportText = formatTaxAwareReport(portfolio, totalValue, actions, warnings, accountType);

  if (typeof showModal === 'function') {
    showModal(`Tax-Aware Analysis — ${portfolio.name}`, reportText);
  }

  if (typeof copyText === 'function') {
    copyText(reportText)
      .then(() => {
        if (typeof showCopyFeedback === 'function') showCopyFeedback('Auto-copied to clipboard');
      })
      .catch(() => { /* user can copy manually */ });
  }
}
