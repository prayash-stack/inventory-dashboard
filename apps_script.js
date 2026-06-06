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

  // ── 2. Determine period from IN/OUT sheet ──
  const inoutSheet = ss.getSheetByName(CONFIG.INOUT_SHEET);
  const inoutData  = inoutSheet.getDataRange().getValues();
  const dateCol    = 5; // column F (0-indexed) = Date

  let minDate = new Date();
  let maxDate = new Date(2000, 0, 1);

  for (let r = 1; r < inoutData.length; r++) {
    const raw = inoutData[r][dateCol];
    if (!raw) continue;
    const d = raw instanceof Date ? raw : new Date(raw);
    if (isNaN(d.getTime())) continue;
    if (d < minDate) minDate = d;
    if (d > maxDate) maxDate = d;
  }

  const today      = new Date();
  const periodDays = Math.max(1, Math.round((today - minDate) / (1000 * 60 * 60 * 24)));

  // ── 3. Build items array ──
  const items = [];
  const categoryTotals = {};

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

    items.push({
      name, cat, loc, unit,
      totalIn:     Math.round(inQty),
      totalOut:    Math.round(outQty),
      currentStock:Math.round(stock),
      dailyRate:   parseFloat(dailyRate.toFixed(1)),
      monthlyRate,
      daysLeft:    Math.max(-999, daysLeft),
      reorderQty,
    });

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
  Logger.log('Categories: ' + data.categories.map(c => c.name + ':' + c.totalOut).join(', '));
}
