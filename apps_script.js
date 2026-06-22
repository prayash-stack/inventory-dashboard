// ============================================================
// INVENTORY DASHBOARD — Google Apps Script
// Reads your Google Sheet → pushes data.json to GitHub
// Run this daily via a time-based trigger
// ============================================================

// ── CONFIGURATION ── fill these in (Step 4 of the guide) ──
const CONFIG = {
  SHEET_ID:        'YOUR_GOOGLE_SHEET_ID',     // from your sheet URL
  STOCK_SHEET:     'Stock',                    // name of the Stock tab
  INOUT_SHEET:     'IN/OUT',                   // name of the IN/OUT tab

  // ── Optional 2025 history archive ──
  // The archive's IN/OUT log is merged in (read-only) to extend consumption
  // history to ~15 months for seasonality/forecasting. It is fail-safe: if the
  // sheet can't be opened the live dashboard is unaffected (history just hides).
  ARCHIVE_SHEET_ID:    '1rCr_wg0mVYzyD5HV6YYUb8-GG3RnlZZghCQkURXhe4c',
  ARCHIVE_INOUT_SHEET: 'IN/OUT',
  ARCHIVE_CUTOFF:      '2026-03-31',           // keep archive rows BEFORE this date (live owns the rest → no double-count)
  CODE_REMAP:          { DS001: 'DB001', OTH074: 'OIL001', DG043: 'DG038' }, // archive SKU code → live SKU code (safety net; data already aligned)

  GITHUB_TOKEN:    'YOUR_GITHUB_TOKEN',        // personal access token
  GITHUB_OWNER:    'YOUR_GITHUB_USERNAME',     // e.g. prayash
  GITHUB_REPO:     'inventory-dashboard',      // repo name you create
  GITHUB_FILE:     'data.json',               // filename in the repo
};
// ──────────────────────────────────────────────────────────


function pushDataToGitHub() {
  try {
    const data = buildDashboardData();
    const json = JSON.stringify(data, null, 2);
    uploadToGitHub(json);
    Logger.log('✅ Dashboard data updated successfully at ' + new Date());
  } catch (e) {
    Logger.log('❌ Error: ' + e.message);
    // Optional: email yourself on failure
    // MailApp.sendEmail('your@email.com', 'Dashboard update failed', e.message);
  }
}


// ── Web App endpoint: lets the dashboard's "Refresh" button rebuild on demand ──
// Deploy:  Apps Script editor → Deploy → New deployment → select type "Web app",
//          Execute as: Me,  Who has access: Anyone.
//          Copy the /exec URL it gives you into REBUILD_URL in index.html.
function doGet(e) {
  const cache = CacheService.getScriptCache();
  if (cache.get('rebuilding')) {
    // Throttle: ignore repeat triggers within 20s so the button can't spam rebuilds.
    return jsonOut({ ok: true, throttled: true, note: 'A rebuild already ran in the last 20s' });
  }
  cache.put('rebuilding', '1', 20);
  try {
    const data = buildDashboardData();
    uploadToGitHub(JSON.stringify(data, null, 2));
    return jsonOut({ ok: true, updatedAt: data.meta.updatedAt, totalSkus: data.meta.totalSkus });
  } catch (err) {
    return jsonOut({ ok: false, error: String((err && err.message) || err) });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// Robust date parser: handles real Date cells AND "01 Apr 26"-style text.
function parseSheetDate(raw) {
  if (raw instanceof Date) return isNaN(raw.getTime()) ? null : raw;
  if (raw == null || raw === '') return null;
  const s = String(raw).trim();
  const m = s.match(/^(\d{1,2})\s+([A-Za-z]{3})\w*\s+(\d{2,4})$/);
  if (m) {
    const MM = { jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11 };
    let y = +m[3]; if (y < 100) y += 2000;
    const d = new Date(y, MM[m[2].toLowerCase()], +m[1]);
    return isNaN(d.getTime()) ? null : d;
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function buildDashboardData() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

  // ── 1. Read Stock sheet (summary) ──
  const stockSheet = ss.getSheetByName(CONFIG.STOCK_SHEET);
  const stockData  = stockSheet.getDataRange().getValues();
  const stockHeaders = stockData[0];

  const col = (name) => stockHeaders.findIndex(h =>
    h.toString().toLowerCase().trim() === name.toLowerCase()
  );

  const nameCol    = col('Product Name');
  const catCol     = col('Product Category');
  const locCol     = col('Location');
  const inCol      = col('In');
  const outCol     = col('Out');
  const stockCol   = col('Current Stock');
  const unitCol    = col('Unit');
  // Optional unit-cost column (enables the inventory-value / money view). Add any of these
  // headers to your Stock sheet to switch it on.
  let   costCol    = col('Cost'); if (costCol < 0) costCol = col('Unit Cost'); if (costCol < 0) costCol = col('Price');

  // IN/OUT sheet columns: A Products, B Location, C In/Out, D Quantity, E Units, F Date
  const PROD_C = 0, TYPE_C = 2, QTY_C = 3, dateCol = 5;

  // SKU code = text in the trailing (...) of a name, e.g. "... (PNT013)" → "PNT013".
  // The log has inconsistent name casing ("Touch Up" vs "Touch up"), so we key on the code.
  const skuCode = (s) => {
    const m = String(s || '').match(/\(([^)]+)\)\s*$/);
    return m ? m[1].trim().toUpperCase() : null;
  };
  const norm = (s) => String(s || '').trim().toLowerCase();

  // ── 2. Read the live IN/OUT log ──
  const inoutSheet = ss.getSheetByName(CONFIG.INOUT_SHEET);
  const liveData   = inoutSheet.getDataRange().getValues();
  const liveRows   = [];
  for (let r = 1; r < liveData.length; r++) liveRows.push(liveData[r]);

  // ── 2a. Merge in the 2025 history archive (read-only, fail-safe) ──
  // Older transactions extend consumption history to ~15 months. We cut the
  // overlap at ARCHIVE_CUTOFF (live owns dates from there on) so nothing is
  // double-counted, and remap any stale SKU codes onto the live codes.
  const cutoff = CONFIG.ARCHIVE_CUTOFF ? parseSheetDate(CONFIG.ARCHIVE_CUTOFF) : null;
  const remap  = CONFIG.CODE_REMAP || {};
  const remapName = (s) => {
    const c = skuCode(s);
    return (c && remap[c]) ? String(s).replace(/\(([^)]+)\)\s*$/, '(' + remap[c] + ')') : s;
  };
  const archRows = [];
  if (CONFIG.ARCHIVE_SHEET_ID) {
    try {
      const aSheet = SpreadsheetApp.openById(CONFIG.ARCHIVE_SHEET_ID)
        .getSheetByName(CONFIG.ARCHIVE_INOUT_SHEET);
      if (aSheet) {
        const aData = aSheet.getDataRange().getValues();
        for (let r = 1; r < aData.length; r++) {
          const d = parseSheetDate(aData[r][dateCol]);
          if (!d) continue;
          if (cutoff && d.getTime() >= cutoff.getTime()) continue; // live owns dates >= cutoff
          const row = aData[r].slice();
          row[PROD_C] = remapName(row[PROD_C]);
          archRows.push(row);
        }
      }
    } catch (e) {
      // Archive unreachable → history simply won't show; live dashboard unaffected.
    }
  }

  // Combined log powers the usage windows + 15-month history.
  const combined = liveRows.concat(archRows);

  // ── 3. Live period (drives headline stock rates — unchanged behavior) ──
  let minDate = new Date();
  let maxDate = new Date(2000, 0, 1);
  for (let i = 0; i < liveRows.length; i++) {
    const d = parseSheetDate(liveRows[i][dateCol]);
    if (!d) continue;
    if (d < minDate) minDate = d;
    if (d > maxDate) maxDate = d;
  }
  const today      = new Date();
  const periodDays = Math.max(1, Math.round((today - minDate) / (1000 * 60 * 60 * 24)));
  const WEEKS      = Math.max(1, Math.ceil(periodDays / 7)); // live-period consumption-trend chart

  // ── 3a. Full-history month list (drives the seasonality / forecasting view) ──
  let histMin = new Date();
  for (let i = 0; i < combined.length; i++) {
    const d = parseSheetDate(combined[i][dateCol]);
    if (d && d < histMin) histMin = d;
  }
  const monthKey = (d) => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
  const months = [];
  { let y = histMin.getFullYear(), m = histMin.getMonth();
    const ey = today.getFullYear(), em = today.getMonth();
    while (y < ey || (y === ey && m <= em)) { months.push(y + '-' + String(m + 1).padStart(2, '0')); m++; if (m > 11) { m = 0; y++; } } }
  const monthIdx = {}; months.forEach((k, i) => monthIdx[k] = i);
  const liveStartMonth = monthKey(minDate);

  // ── 4. Usage windows + per-window OUT (powers the dashboard's "usage window" filter) ──
  const WINDOWS = [7, 30, 90];

  // Codes used by more than one distinct product in the Stock sheet are data-entry dupes;
  // for those we fall back to name matching so their windows don't get merged.
  const codeOwners = {};
  for (let r = 1; r < stockData.length; r++) {
    const nm = stockData[r][nameCol] ? stockData[r][nameCol].toString().trim() : '';
    if (!nm) continue;
    const c = skuCode(nm);
    if (!c) continue;
    (codeOwners[c] = codeOwners[c] || {})[norm(nm)] = true;
  }
  const dupCodes = {};
  Object.keys(codeOwners).forEach(c => { if (Object.keys(codeOwners[c]).length > 1) dupCodes[c] = true; });

  // Aggregate OUT quantity by code and by name — windows + monthly history (over the
  // COMBINED log), and weekly (over the LIVE log only, so the short-term chart stays clean).
  const DAY = 1000 * 60 * 60 * 24;
  const outByCode = {}, outByName = {}, monByCode = {}, monByName = {};
  const overallMonthly = new Array(months.length).fill(0);
  const addWin = (map, key, ageDays, qty) => {
    if (!map[key]) map[key] = { 7: 0, 30: 0, 90: 0 };
    for (let i = 0; i < WINDOWS.length; i++) if (ageDays <= WINDOWS[i]) map[key][WINDOWS[i]] += qty;
  };
  const addMon = (map, key, mi, qty) => {
    if (!map[key]) map[key] = new Array(months.length).fill(0);
    map[key][mi] += qty;
  };
  for (let i = 0; i < combined.length; i++) {
    const row = combined[i];
    if ((row[TYPE_C] || '').toString().trim().toLowerCase() !== 'out') continue;
    const d = parseSheetDate(row[dateCol]);
    if (!d) continue;
    const qty = parseFloat(row[QTY_C]) || 0;
    const ageDays = (today - d) / DAY;
    const mi = monthIdx[monthKey(d)];
    const c = skuCode(row[PROD_C]);
    if (c) { addWin(outByCode, c, ageDays, qty); if (mi != null) addMon(monByCode, c, mi, qty); }
    addWin(outByName, norm(row[PROD_C]), ageDays, qty);
    if (mi != null) { addMon(monByName, norm(row[PROD_C]), mi, qty); overallMonthly[mi] += qty; }
  }

  // Weekly buckets — LIVE log only (keeps the short-term trend chart readable).
  const wkByCode = {}, wkByName = {};
  const overallWeekly = new Array(WEEKS).fill(0);
  const addWk = (map, key, wi, qty) => {
    if (!map[key]) map[key] = new Array(WEEKS).fill(0);
    map[key][wi] += qty;
  };
  for (let i = 0; i < liveRows.length; i++) {
    const row = liveRows[i];
    if ((row[TYPE_C] || '').toString().trim().toLowerCase() !== 'out') continue;
    const d = parseSheetDate(row[dateCol]);
    if (!d) continue;
    const qty = parseFloat(row[QTY_C]) || 0;
    let wi = Math.floor((d - minDate) / (7 * DAY));
    if (wi < 0) wi = 0; if (wi >= WEEKS) wi = WEEKS - 1;
    const c = skuCode(row[PROD_C]);
    if (c) addWk(wkByCode, c, wi, qty);
    addWk(wkByName, norm(row[PROD_C]), wi, qty);
    overallWeekly[wi] += qty;
  }

  const winFor = (name) => {
    const c = skuCode(name);
    if (c && !dupCodes[c] && outByCode[c]) return outByCode[c];
    return outByName[norm(name)] || { 7: 0, 30: 0, 90: 0 };
  };
  const wkFor = (name) => {
    const c = skuCode(name);
    if (c && !dupCodes[c] && wkByCode[c]) return wkByCode[c];
    return wkByName[norm(name)] || new Array(WEEKS).fill(0);
  };
  const monFor = (name) => {
    const c = skuCode(name);
    if (c && !dupCodes[c] && monByCode[c]) return monByCode[c];
    return monByName[norm(name)] || new Array(months.length).fill(0);
  };

  // ── 3. Build items array ──
  const items = [];
  const categoryTotals = {};
  let totalStockValue = 0, deadStockValue = 0;

  for (let r = 1; r < stockData.length; r++) {
    const row = stockData[r];
    const name  = row[nameCol]  ? row[nameCol].toString().trim()  : '';
    const cat   = row[catCol]   ? row[catCol].toString().trim()   : '';
    const loc   = row[locCol]   ? row[locCol].toString().trim()   : '';
    const inQty = parseFloat(row[inCol])    || 0;
    const outQty= parseFloat(row[outCol])   || 0;
    const stock = parseFloat(row[stockCol]) || 0;
    const unit  = row[unitCol]  ? row[unitCol].toString().trim()  : 'PCS';

    if (!name || !cat) continue;
    if (inQty === 0 && outQty === 0 && stock === 0) continue;

    const dailyRate   = outQty / periodDays;
    const daysLeft    = dailyRate > 0 ? Math.round(stock / dailyRate) : 9999;
    const monthlyRate = Math.round(outQty / (periodDays / 30));
    const reorderQty  = Math.round(dailyRate * 45); // 45-day supply

    // Per-window figures: same stock, but rate/days-left/reorder from each window's OUT.
    const win = winFor(name);
    const periods = {};
    WINDOWS.forEach(w => {
      const o  = win[w] || 0;
      const dr = o / w;
      periods[w] = {
        totalOut:    Math.round(o),
        dailyRate:   parseFloat(dr.toFixed(1)),
        monthlyRate: Math.round(o / (w / 30)),
        daysLeft:    dr > 0 ? Math.max(-999, Math.round(stock / dr)) : 9999,
        reorderQty:  Math.round(dr * 45),
      };
    });

    const cost = costCol >= 0 ? (parseFloat(row[costCol]) || 0) : 0;
    const stockValue = Math.round(cost * stock);

    const item = {
      name, cat, loc, unit,
      totalIn:     Math.round(inQty),
      totalOut:    Math.round(outQty),
      currentStock:Math.round(stock),
      dailyRate:   parseFloat(dailyRate.toFixed(1)),
      monthlyRate,
      daysLeft:    Math.max(-999, daysLeft),
      reorderQty,
      periods,
      trend:       wkFor(name).map(v => Math.round(v)), // units out per week (live period)
    };
    // 15-month consumption history (units out per calendar month, aligned to meta.months).
    const mon = monFor(name).map(v => Math.round(v));
    item.monthlyOut = mon;
    // Year-over-year on the last COMPLETE month (current month is partial → skip it).
    const lastFull = months.length - 2, yearAgo = lastFull - 12;
    if (lastFull >= 0 && yearAgo >= 0 && (mon[lastFull] || mon[yearAgo])) {
      item.yoy = { month: months[lastFull], cur: mon[lastFull], prev: mon[yearAgo] };
    }
    if (costCol >= 0) { item.cost = parseFloat(cost.toFixed(2)); item.stockValue = stockValue; }
    items.push(item);

    if (costCol >= 0) {
      totalStockValue += stockValue;
      if (stock > 0 && outQty === 0) deadStockValue += stockValue;
    }

    if (!categoryTotals[cat]) categoryTotals[cat] = 0;
    categoryTotals[cat] += outQty;
  }

  // ── 4. Sort items by units sold descending ──
  items.sort((a, b) => b.totalOut - a.totalOut);

  // ── 5. Category summary ──
  const categories = Object.entries(categoryTotals)
    .map(([name, totalOut]) => ({ name, totalOut: Math.round(totalOut) }))
    .sort((a, b) => b.totalOut - a.totalOut);

  const totalUnitsOut = items.reduce((s, i) => s + i.totalOut, 0);
  const criticalCount = items.filter(i => i.daysLeft <= 7  && i.totalOut > 0).length;
  const warningCount  = items.filter(i => i.daysLeft > 7 && i.daysLeft <= 30 && i.totalOut > 0).length;

  // Weekly buckets for the consumption-trend chart.
  const weekly = [];
  for (let i = 0; i < WEEKS; i++) {
    const s = new Date(minDate.getTime() + i * 7 * DAY);
    weekly.push({ start: s.toISOString().split('T')[0], totalOut: Math.round(overallWeekly[i]) });
  }

  // 15-month consumption history (seasonality / forecasting). isHistory flags months
  // that come from the 2025 archive (before the live sheet's first month).
  const monthly = months.map((k, i) => ({
    month:     k,
    totalOut:  Math.round(overallMonthly[i]),
    isHistory: k < liveStartMonth,
  }));

  return {
    meta: {
      updatedAt:   new Date().toISOString(),
      periodDays,
      startDate:   minDate.toISOString().split('T')[0],
      endDate:     today.toISOString().split('T')[0],
      totalSkus:   items.length,
      totalUnitsOut,
      criticalCount,
      warningCount,
      windows:     WINDOWS,
      weeks:       WEEKS,
      weekly,
      // 15-month history (archive-merged). hasHistory is false if the archive was unreachable.
      months,
      monthly,
      hasHistory:      archRows.length > 0,
      historyStart:    histMin.toISOString().split('T')[0],
      liveStartMonth,
      archiveRows:     archRows.length,
      hasCost:         costCol >= 0,
      totalStockValue: Math.round(totalStockValue),
      deadStockValue:  Math.round(deadStockValue),
    },
    items,
    categories,
  };
}


function uploadToGitHub(jsonContent) {
  const url = `https://api.github.com/repos/${CONFIG.GITHUB_OWNER}/${CONFIG.GITHUB_REPO}/contents/${CONFIG.GITHUB_FILE}`;

  // Get current file SHA (needed to update an existing file)
  let sha = null;
  try {
    const getResp = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        Authorization: `Bearer ${CONFIG.GITHUB_TOKEN}`,
        Accept: 'application/vnd.github+json',
      },
      muteHttpExceptions: true,
    });
    if (getResp.getResponseCode() === 200) {
      sha = JSON.parse(getResp.getContentText()).sha;
    }
  } catch (e) {
    // File doesn't exist yet — that's fine, sha stays null
  }

  const payload = {
    message: `Dashboard update ${new Date().toISOString().split('T')[0]}`,
    content: Utilities.base64Encode(jsonContent, Utilities.Charset.UTF_8),
    ...(sha ? { sha } : {}),
  };

  const resp = UrlFetchApp.fetch(url, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${CONFIG.GITHUB_TOKEN}`,
      Accept: 'application/vnd.github+json',
      'Content-Type': 'application/json',
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const code = resp.getResponseCode();
  if (code !== 200 && code !== 201) {
    throw new Error(`GitHub API returned ${code}: ${resp.getContentText()}`);
  }
}


// ── TEST: run this manually first to check everything works ──
function testRun() {
  Logger.log('Building data...');
  const data = buildDashboardData();
  Logger.log('Period: ' + data.meta.periodDays + ' days');
  Logger.log('Items: '  + data.meta.totalSkus);
  Logger.log('Total out: ' + data.meta.totalUnitsOut);
  Logger.log('Critical: ' + data.meta.criticalCount);
  Logger.log('Top 3:');
  data.items.slice(0, 3).forEach((it, i) => {
    Logger.log(`  ${i+1}. ${it.name} — ${it.totalOut} out, ${it.daysLeft}d left`);
  });
  const s = data.items[0];
  Logger.log('Windows: ' + data.meta.windows.join('/') + 'd | top item OUT 7/30/90d: ' +
    s.periods[7].totalOut + '/' + s.periods[30].totalOut + '/' + s.periods[90].totalOut);
  Logger.log('Categories: ' + data.categories.map(c => c.name + ':' + c.totalOut).join(', '));
  // History / archive integration
  Logger.log('History: ' + (data.meta.hasHistory ? 'ON' : 'off') +
    ' | archive rows merged: ' + data.meta.archiveRows +
    ' | months: ' + data.meta.months.length + ' (' + data.meta.historyStart + ' → ' + data.meta.endDate + ')');
  Logger.log('Monthly OUT: ' + data.meta.monthly.map(m => m.month + (m.isHistory ? '*' : '') + ':' + m.totalOut).join('  '));
}
