// ═══════════════════════════════════════════════════════════════
// GAME STATION LOUNGE POS — Google Apps Script Backend v2
// Paste this entire file into Extensions > Apps Script
// ═══════════════════════════════════════════════════════════════

const SHEET_ID    = '14XuYMKu5fphq9w0ZqVANYUDQ3-hJvrTsLvYL_hqpO9E';
const ALERT_EMAIL = 'gamestationlounge9@gmail.com';

// Items >= 10,000 RWF → alert at 2 remaining
// Items <  10,000 RWF → alert at 5 remaining
const HIGH_VALUE_THRESHOLD = 10000;
const ALERT_LOW  = 5;
const ALERT_HIGH = 2;

// ── ALERT TRACKING (once per product per day) ─────────────────────
function getAlertsSentToday() {
  const props = PropertiesService.getScriptProperties();
  const today = new Date().toLocaleDateString('en-GB');
  const raw   = props.getProperty('alerts_' + today);
  return raw ? JSON.parse(raw) : {};
}
function markAlertSent(product) {
  const props = PropertiesService.getScriptProperties();
  const today = new Date().toLocaleDateString('en-GB');
  const sent  = getAlertsSentToday();
  sent[product] = true;
  props.setProperty('alerts_' + today, JSON.stringify(sent));
}

// ── ROUTING ───────────────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    return respond(getResult(data));
  } catch (err) {
    return respond({ ok: false, error: err.toString() });
  }
}
function doGet(e) {
  try {
    let result;
    if (e.parameter.payload) {
      result = getResult(JSON.parse(e.parameter.payload));
    } else {
      const action = e.parameter.action;
      if      (action === 'GET_LAST_STOCK') result = getLastStock();
      else if (action === 'GET_PRICES')     result = getPrices();
      else                                  result = { ok: false, error: 'Unknown action' };
    }
    // JSONP: wrap result in callback(json) so <script> tag injection works
    const cb = e.parameter.callback;
    if (cb) {
      return ContentService
        .createTextOutput(cb + '(' + JSON.stringify(result) + ')')
        .setMimeType(ContentService.MimeType.TEXT);
    }
    return respond(result);
  } catch (err) {
    return respond({ ok: false, error: err.toString() });
  }
}
// Returns a plain object — used by both doPost and JSONP doGet
function getResult(data) {
  const action = data.action;
  if (action === 'LOG_SALE')            return logSale(data);
  if (action === 'SAVE_TAB')            return saveTab(data);
  if (action === 'END_DAY')             return endDay(data);
  if (action === 'END_DAY_SUMMARY')     return endDaySummary(data);
  if (action === 'SAVE_CLOSING_STOCK')  return saveClosingStock(data);
  if (action === 'GET_LAST_STOCK')      return getLastStock();
  if (action === 'UPDATE_PRICES')       return updatePrices(data);
  if (action === 'GET_PRICES')          return getPrices();
  if (action === 'CHECK_STOCK_ALERT')   return checkStockAlert(data);
  return { ok: false, error: 'Unknown action' };
}
function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 1. LOG SALE ───────────────────────────────────────────────────
function logSale(data) {
  const wb    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(wb, 'Sales Log');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Date','Time','Bartender','Product','Qty Sold','Unit Price (RWF)','Revenue (RWF)','Remaining Stock','Payment']);
    sheet.getRange(1,1,1,9).setFontWeight('bold').setBackground('#1A5276').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1,9,150);
  }
  sheet.appendRow([data.date,data.time,data.bartender,data.product,data.qty,data.unitPrice||0,data.revenue||0,data.remaining,data.paymentMethod||'CASH']);
  if (data.revenue > 0) sheet.getRange(sheet.getLastRow(),7).setBackground('#D5F5E3');
  return { ok: true };
}

// ── 1b. SAVE TAB (one entry per closed tab) ───────────────────────
function saveTab(data) {
  const wb    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(wb, 'Sales Log');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Date','Time','Bartender','Table','Items (JSON)','Total (RWF)','Payment']);
    sheet.getRange(1,1,1,7).setFontWeight('bold').setBackground('#1A5276').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1,7,160);
  }
  sheet.appendRow([data.date, data.time, data.bartender, data.tableName,
                   data.items, Number(data.total)||0, data.paymentMethod]);
  sheet.getRange(sheet.getLastRow(), 6).setBackground('#D5F5E3');
  return { ok: true };
}

// ── 2. CHECK STOCK & SEND ALERT EMAIL ────────────────────────────
function checkStockAlert(data) {
  const product   = data.product;
  const remaining = data.remaining;
  const unitPrice = data.unitPrice || 0;

  // Determine correct threshold based on price
  const threshold = unitPrice >= HIGH_VALUE_THRESHOLD ? ALERT_HIGH : ALERT_LOW;

  // Not low enough — no alert needed
  if (remaining > threshold) return { ok: true, alerted: false };

  // Already sent alert for this product today — don't spam
  if (getAlertsSentToday()[product]) return { ok: true, alerted: false, reason: 'already sent today' };

  const timeNow = new Date().toLocaleString('en-GB');
  const subject  = `⚠️ Low Stock Alert — ${product} | Game Station Lounge`;

  // Plain text version
  const body = `⚠️ KABWA Eric,

Low Stock Alert — ${product} is running low (${remaining} remaining). Please restock.

  Product   : ${product}
  Remaining : ${remaining} bottle${remaining !== 1 ? 's' : ''}
  Threshold : ${threshold} bottles
  Time      : ${timeNow}

Please arrange a restock as soon as possible to avoid running out during service.

Regards,
Game Station Lounge POS System`;

  // HTML version
  const htmlBody = `
<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;background:#ffffff;border:1px solid #e5e7eb;border-radius:12px;overflow:hidden">

  <!-- Header -->
  <div style="background:#1a1a2e;padding:22px 28px;text-align:center">
    <h2 style="color:#f59e0b;margin:0;font-size:20px;letter-spacing:3px">GAME STATION LOUNGE</h2>
    <p style="color:#94a3b8;margin:5px 0 0;font-size:11px;letter-spacing:2px">POINT OF SALE SYSTEM</p>
  </div>

  <!-- Alert banner -->
  <div style="background:#fef3c7;border-bottom:2px solid #f59e0b;padding:16px 28px;text-align:center">
    <p style="font-size:32px;margin:0">⚠️</p>
    <h3 style="color:#92400e;margin:6px 0 0;font-size:17px;letter-spacing:1px">LOW STOCK ALERT</h3>
  </div>

  <!-- Body -->
  <div style="padding:28px">
    <p style="color:#111827;font-size:15px;margin:0 0 8px">Dear <strong>KABWA Eric</strong>,</p>
    <p style="color:#374151;font-size:14px;line-height:1.7;margin:0 0 24px">
      ⚠️ KABWA Eric, Low Stock Alert — <strong style="color:#ef4444">${product}</strong> is running low
      (<strong>${remaining} remaining</strong>). Please restock.
    </p>

    <!-- Details table -->
    <table style="width:100%;border-collapse:collapse;border-radius:8px;overflow:hidden;margin-bottom:24px;font-size:13px">
      <tr style="background:#f8fafc">
        <td style="padding:11px 16px;border-bottom:1px solid #e5e7eb;color:#64748b;width:38%">Product</td>
        <td style="padding:11px 16px;border-bottom:1px solid #e5e7eb;font-weight:bold;color:#111827">${product}</td>
      </tr>
      <tr>
        <td style="padding:11px 16px;border-bottom:1px solid #e5e7eb;color:#64748b;background:#f8fafc">Remaining</td>
        <td style="padding:11px 16px;border-bottom:1px solid #e5e7eb;font-weight:bold;color:#ef4444;font-size:17px">${remaining} bottle${remaining !== 1 ? 's' : ''}</td>
      </tr>
      <tr style="background:#f8fafc">
        <td style="padding:11px 16px;border-bottom:1px solid #e5e7eb;color:#64748b">Alert Level</td>
        <td style="padding:11px 16px;border-bottom:1px solid #e5e7eb;color:#111827">${threshold} bottles or below</td>
      </tr>
      <tr>
        <td style="padding:11px 16px;color:#64748b;background:#f8fafc">Date & Time</td>
        <td style="padding:11px 16px;color:#111827">${timeNow}</td>
      </tr>
    </table>

    <!-- Warning box -->
    <div style="background:#fee2e2;border:1px solid #fca5a5;border-radius:8px;padding:14px 18px;text-align:center">
      <p style="color:#991b1b;font-weight:bold;margin:0;font-size:13px">
        🚨 Please arrange a restock as soon as possible to avoid running out during service.
      </p>
    </div>
  </div>

  <!-- Footer -->
  <div style="background:#f1f5f9;padding:14px 28px;text-align:center;border-top:1px solid #e5e7eb">
    <p style="color:#94a3b8;font-size:11px;margin:0">Game Station Lounge POS — Automated Stock Alert System</p>
  </div>

</div>`;

  try {
    Logger.log(`Attempting to send email to ${ALERT_EMAIL} for product: ${product}`);
    MailApp.sendEmail({ to: ALERT_EMAIL, subject, body, htmlBody });
    markAlertSent(product);
    Logger.log('Email sent successfully and alert marked.');
    return { ok: true, alerted: true, product, remaining };
  } catch(err) {
    Logger.log(`Error sending email: ${err.toString()}`);
    return { ok: false, error: err.toString() };
  }
}

// ── 3. END OF DAY ─────────────────────────────────────────────────
function endDay(data) {
  const wb = SpreadsheetApp.openById(SHEET_ID);

  const summary = getOrCreateSheet(wb, 'Daily Summary');
  if (summary.getLastRow() === 0) {
    summary.appendRow(['Date','Bartender','Total Units Sold','Total Revenue (RWF)','Top Product','Shift Start','Shift End']);
    summary.getRange(1,1,1,7).setFontWeight('bold').setBackground('#1E8449').setFontColor('#FFFFFF');
    summary.setFrozenRows(1);
    summary.setColumnWidths(1,7,160);
  }

  const top = data.sales && data.sales.length > 0
    ? data.sales.reduce((a,b) => a.sold > b.sold ? a : b).product : '—';
  summary.appendRow([data.date,data.bartender,data.totalUnits,data.totalRevenue,top,data.shiftStart,data.shiftEnd]);

  const closing = getOrCreateSheet(wb, 'Closing Stock');
  closing.clearContents();
  closing.appendRow(['Product','Closing Stock','Date']);
  closing.getRange(1,1,1,3).setFontWeight('bold').setBackground('#7D3C98').setFontColor('#FFFFFF');
  // closingStock is a slim all-products array; fall back to stock if not present
  const stockForClosing = data.closingStock || data.stock;
  if (stockForClosing) stockForClosing.forEach(item => closing.appendRow([item.product,item.remaining,data.date]));

  const daySheet = getOrCreateSheet(wb, data.date);
  daySheet.clearContents();
  daySheet.appendRow(['Product','Opening Stock','Restocked','Total Available','Sold','Unit Price (RWF)','Revenue (RWF)','Closing Stock']);
  daySheet.getRange(1,1,1,8).setFontWeight('bold').setBackground('#2E86C1').setFontColor('#FFFFFF');
  daySheet.setFrozenRows(1);
  daySheet.setColumnWidths(1,8,150);
  if (data.stock) {
    data.stock.forEach((item,i) => {
      const bg = i%2===0 ? '#EBF5FB' : '#FFFFFF';
      daySheet.appendRow([item.product,item.opening,item.restocked||0,
        item.totalAvailable||(item.opening+(item.restocked||0)),
        item.sold,item.unitPrice||0,item.revenue||0,item.remaining]);
      daySheet.getRange(i+2,1,1,8).setBackground(bg);
    });
    const lastRow = data.stock.length + 2;
    daySheet.appendRow(['TOTAL','','','',data.totalUnits,'',data.totalRevenue,'']);
    daySheet.getRange(lastRow,1,1,8).setFontWeight('bold').setBackground('#D6EAF8');
  }
  return { ok: true, message: 'Day closed successfully' };
}

// ── 3b. END DAY SUMMARY (split call 1 — small payload) ───────────────
function endDaySummary(data) {
  const wb = SpreadsheetApp.openById(SHEET_ID);

  // Daily Summary — single row append (always one row, no loop needed)
  const summary = getOrCreateSheet(wb, 'Daily Summary');
  if (summary.getLastRow() === 0) {
    summary.appendRow(['Date','Bartender','Total Units Sold','Total Revenue (RWF)','Top Product','Shift Start','Shift End']);
    summary.getRange(1,1,1,7).setFontWeight('bold').setBackground('#1E8449').setFontColor('#FFFFFF');
    summary.setFrozenRows(1);
    summary.setColumnWidths(1,7,160);
  }
  const top = data.sales && data.sales.length > 0
    ? data.sales.reduce((a,b) => a.sold > b.sold ? a : b).product : '—';
  summary.appendRow([data.date,data.bartender,data.totalUnits,data.totalRevenue,top,data.shiftStart,data.shiftEnd]);

  // Dated daily breakdown — one setValues call instead of N appendRow calls
  if (data.stock && data.stock.length > 0) {
    const daySheet = getOrCreateSheet(wb, data.date);
    daySheet.clearContents();

    const dataRows = data.stock.map(item => [
      item.product,
      item.opening,
      item.restocked || 0,
      item.opening + (item.restocked || 0),
      item.sold,
      item.unitPrice || 0,
      item.sold * (item.unitPrice || 0),
      item.remaining
    ]);
    const totalRow = ['TOTAL','','','',data.totalUnits,'',data.totalRevenue,''];
    const headers  = ['Product','Opening Stock','Restocked','Total Available','Sold','Unit Price (RWF)','Revenue (RWF)','Closing Stock'];

    daySheet.getRange(1, 1, dataRows.length + 2, 8)
      .setValues([headers, ...dataRows, totalRow]);

    // All formatting in bulk — one call per style, not one per row
    daySheet.getRange(1,1,1,8).setFontWeight('bold').setBackground('#2E86C1').setFontColor('#FFFFFF');
    daySheet.setFrozenRows(1);
    daySheet.setColumnWidths(1,8,150);
    const bgGrid = dataRows.map((_,i) => Array(8).fill(i%2===0 ? '#EBF5FB' : '#FFFFFF'));
    daySheet.getRange(2, 1, dataRows.length, 8).setBackgrounds(bgGrid);
    daySheet.getRange(dataRows.length + 2, 1, 1, 8).setFontWeight('bold').setBackground('#D6EAF8');
  }

  formatDailySummary(wb);
  return { ok: true };
}

// ── 3c. SAVE CLOSING STOCK (split call 2 — small payload) ────────────
function saveClosingStock(data) {
  const wb = SpreadsheetApp.openById(SHEET_ID);
  const closing = getOrCreateSheet(wb, 'Closing Stock');
  closing.clearContents();

  const rows = data.stock ? data.stock.map(item => [item.product, item.remaining, data.date]) : [];
  closing.getRange(1, 1, rows.length + 1, 3)
    .setValues([['Product','Closing Stock','Date'], ...rows]);
  closing.getRange(1,1,1,3).setFontWeight('bold').setBackground('#7D3C98').setFontColor('#FFFFFF');

  setupSheets(wb);
  return { ok: true };
}

// ── 4. GET LAST CLOSING STOCK ─────────────────────────────────────
function getLastStock() {
  const wb    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = wb.getSheetByName('Closing Stock');
  if (!sheet || sheet.getLastRow() <= 1) return { ok: true, stock: [], message: 'No previous stock found' };
  const rows  = sheet.getRange(2,1,sheet.getLastRow()-1,3).getValues();
  const stock = rows.filter(r=>r[0]!=='').map(r=>({ product:r[0], remaining:r[1], date:r[2] }));
  return { ok: true, stock };
}

// ── 5. SAVE PRICES ────────────────────────────────────────────────
function updatePrices(data) {
  const wb    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(wb, 'Prices');
  sheet.clearContents();
  const rows = data.prices ? data.prices.map(p => [p.product, p.price]) : [];
  sheet.getRange(1, 1, rows.length + 1, 2)
    .setValues([['Product','Unit Price (RWF)'], ...rows]);
  sheet.getRange(1,1,1,2).setFontWeight('bold').setBackground('#F39C12').setFontColor('#FFFFFF');
  return { ok: true };
}

// ── 6. GET PRICES ─────────────────────────────────────────────────
function getPrices() {
  const wb    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = wb.getSheetByName('Prices');
  if (!sheet || sheet.getLastRow() <= 1) return { ok: true, prices: [] };
  const rows   = sheet.getRange(2,1,sheet.getLastRow()-1,2).getValues();
  const prices = rows.filter(r=>r[0]!=='').map(r=>({ product:r[0], price:r[1] }));
  return { ok: true, prices };
}

// ══════════════════════════════════════════════════════════════════
// SHEET FORMATTING & DASHBOARD
// setupSheets() is called only from saveClosingStock() — both END DAY
// calls run in parallel so Dashboard must only be written once.
// ══════════════════════════════════════════════════════════════════

function setupSheets(wb) {
  formatDailySummary(wb);
  formatClosingStock(wb);
  formatSalesLog(wb);
  formatPrices(wb);
  updateDashboard(wb);
}

// ── Ensure a branded title row exists at row 1 ───────────────────
function ensureTitleRow(sheet, title, numCols, bgColor) {
  const firstVal = sheet.getRange(1,1).getValue().toString().trim();
  if (firstVal !== title) sheet.insertRowBefore(1);
  sheet.getRange(1,1,1,numCols).merge()
    .setValue(title)
    .setBackground(bgColor || '#1a1a2e')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(13)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);
}

// ── DAILY SUMMARY ────────────────────────────────────────────────
function formatDailySummary(wb) {
  const sheet = wb.getSheetByName('Daily Summary');
  if (!sheet || sheet.getLastRow() === 0) return;

  ensureTitleRow(sheet, 'GAME STATION LOUNGE — DAILY REPORT', 7);

  // Row 2: column headers — amber
  sheet.getRange(2,1,1,7)
    .setBackground('#f59e0b').setFontColor('#000000').setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Data rows: alternating #f8fafc / white
  const dataRows = sheet.getLastRow() - 2;
  if (dataRows > 0) {
    const bgGrid = Array.from({length: dataRows}, (_,i) =>
      Array(7).fill(i%2===0 ? '#f8fafc' : '#FFFFFF'));
    sheet.getRange(3, 1, dataRows, 7).setBackgrounds(bgGrid);
    // Revenue column (#,##0 "RWF")
    sheet.getRange(3, 4, dataRows, 1).setNumberFormat('#,##0" RWF"');
  }

  sheet.setFrozenRows(2);
  sheet.autoResizeColumns(1, 7);
}

// ── CLOSING STOCK ────────────────────────────────────────────────
function formatClosingStock(wb) {
  const sheet = wb.getSheetByName('Closing Stock');
  if (!sheet || sheet.getLastRow() <= 1) return;

  // Header row 1: green
  sheet.getRange(1,1,1,3)
    .setBackground('#10b981').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Conditional row colours based on remaining (col B)
  const dataRows = sheet.getLastRow() - 1;
  if (dataRows > 0) {
    const vals = sheet.getRange(2, 2, dataRows, 1).getValues(); // col B
    const bgGrid = vals.map(r => {
      const rem = Number(r[0]);
      if (rem <= 3)  return ['#fee2e2'];  // red
      if (rem <= 10) return ['#fef3c7'];  // yellow
      return ['#FFFFFF'];
    });
    // Apply colours across all 3 columns
    const fullBg = bgGrid.map(r => [r[0], r[0], r[0]]);
    sheet.getRange(2, 1, dataRows, 3).setBackgrounds(fullBg);
  }

  sheet.autoResizeColumns(1, 3);
}

// ── SALES LOG ────────────────────────────────────────────────────
function formatSalesLog(wb) {
  const sheet = wb.getSheetByName('Sales Log');
  if (!sheet || sheet.getLastRow() <= 1) return;

  sheet.getRange(1,1,1,8)
    .setBackground('#1e40af').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  const dataRows = sheet.getLastRow() - 1;
  if (dataRows > 0) {
    const bgGrid = Array.from({length: dataRows}, (_,i) =>
      Array(8).fill(i%2===0 ? '#f8fafc' : '#FFFFFF'));
    sheet.getRange(2, 1, dataRows, 8).setBackgrounds(bgGrid);
  }

  sheet.autoResizeColumns(1, 8);
}

// ── PRICES ───────────────────────────────────────────────────────
function formatPrices(wb) {
  const sheet = wb.getSheetByName('Prices');
  if (!sheet || sheet.getLastRow() <= 1) return;

  sheet.getRange(1,1,1,2)
    .setBackground('#7c3aed').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  const dataRows = sheet.getLastRow() - 1;
  if (dataRows > 0) {
    sheet.getRange(2, 2, dataRows, 1).setNumberFormat('#,##0" RWF"');
    const bgGrid = Array.from({length: dataRows}, (_,i) =>
      [i%2===0 ? '#f8fafc' : '#FFFFFF', i%2===0 ? '#f8fafc' : '#FFFFFF']);
    sheet.getRange(2, 1, dataRows, 2).setBackgrounds(bgGrid);
  }

  sheet.autoResizeColumns(1, 2);
}

// ── DASHBOARD ────────────────────────────────────────────────────
function updateDashboard(wb) {
  const dash = getOrCreateSheet(wb, 'DASHBOARD');
  dash.clearContents();
  dash.clearFormats();

  // ── Today's summary (last row of Daily Summary) ─────────────
  const summarySheet = wb.getSheetByName('Daily Summary');
  let todayDate='—', bartender='—', totalUnits=0, totalRevenue=0, topProduct='—', shiftStart='—', shiftEnd='—';
  if (summarySheet && summarySheet.getLastRow() > 1) {
    const lastRow = summarySheet.getRange(summarySheet.getLastRow(), 1, 1, 7).getValues()[0];
    // Handle both 1-header format and 2-header format (with title row)
    const firstVal = summarySheet.getRange(1,1).getValue().toString();
    const dataStart = firstVal.indexOf('GAME STATION') === 0 ? 3 : 2;
    const actualLast = summarySheet.getRange(summarySheet.getLastRow(), 1, 1, 7).getValues()[0];
    todayDate    = actualLast[0]; bartender  = actualLast[1];
    totalUnits   = actualLast[2]; totalRevenue = actualLast[3];
    topProduct   = actualLast[4]; shiftStart = actualLast[5]; shiftEnd = actualLast[6];
  }

  // ── Top 5 products (all-time from Sales Log) ─────────────────
  const top5 = [];
  const logSheet = wb.getSheetByName('Sales Log');
  if (logSheet && logSheet.getLastRow() > 1) {
    const logData = logSheet.getRange(2, 1, logSheet.getLastRow()-1, 5).getValues();
    const totals = {};
    logData.forEach(r => {
      const prod = r[3]; const qty = Number(r[4]);
      if (prod) totals[prod] = (totals[prod]||0) + qty;
    });
    Object.entries(totals).sort((a,b)=>b[1]-a[1]).slice(0,5)
      .forEach(([p,q]) => top5.push([p, q]));
  }

  // ── Low stock items (remaining <= 10) ────────────────────────
  const lowStock = [];
  const csSheet = wb.getSheetByName('Closing Stock');
  if (csSheet && csSheet.getLastRow() > 1) {
    const csData = csSheet.getRange(2, 1, csSheet.getLastRow()-1, 2).getValues();
    csData.filter(r => r[0] && Number(r[1]) <= 10)
      .sort((a,b) => Number(a[1])-Number(b[1]))
      .forEach(r => lowStock.push([r[0], r[1]]));
  }

  // ── Build dashboard rows ──────────────────────────────────────
  const rows = [];
  const EMPTY = ['',''];

  // Title
  rows.push(['GAME STATION LOUNGE POS', '']);
  rows.push(EMPTY);

  // Today's performance
  rows.push(['📊  TODAY\'S PERFORMANCE', '']);
  rows.push(['Date',        todayDate]);
  rows.push(['Bartender',   bartender]);
  rows.push(['Revenue',     totalRevenue]);
  rows.push(['Units Sold',  totalUnits]);
  rows.push(['Shift',       shiftStart + ' — ' + shiftEnd]);
  rows.push(['Top Product', topProduct]);
  rows.push(EMPTY);

  // Top 5 products
  rows.push(['🏆  TOP 5 PRODUCTS (ALL TIME)', '']);
  rows.push(['Product', 'Units Sold']);
  top5.forEach(r => rows.push(r));
  if (top5.length === 0) rows.push(['No sales recorded', '']);
  rows.push(EMPTY);

  // Low stock
  rows.push(['⚠️  LOW STOCK ALERTS  (≤ 10 remaining)', '']);
  rows.push(['Product', 'Remaining']);
  lowStock.forEach(r => rows.push(r));
  if (lowStock.length === 0) rows.push(['All stock levels OK', '']);
  rows.push(EMPTY);

  // Timestamp
  rows.push(['Last Updated', new Date().toLocaleString('en-GB')]);

  // Write all at once
  dash.getRange(1, 1, rows.length, 2).setValues(rows);

  // ── Apply formatting ─────────────────────────────────────────
  // Title
  dash.getRange(1,1,1,2).merge()
    .setBackground('#1a1a2e').setFontColor('#f59e0b')
    .setFontWeight('bold').setFontSize(16)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  dash.setRowHeight(1, 48);

  // Section headers (rows 3, 11, top5+13, lowStock+…)
  let r = 3;
  [r, r+8+top5.length+2+1].forEach(() => {}); // recalc below

  // Build row index map for section headers
  const secRows = [3, 11, 11 + top5.length + 2, 11 + top5.length + 2 + lowStock.length + 2 + 1];
  const secActual = [3, 11, 11 + (top5.length||1) + 2, 11 + (top5.length||1) + 2 + (lowStock.length||1) + 2 + 1];

  // Simpler: just scan rows array for section header markers
  rows.forEach((row, i) => {
    const label = row[0].toString();
    const rowNum = i + 1;
    if (label.startsWith('📊') || label.startsWith('🏆') || label.startsWith('⚠️')) {
      dash.getRange(rowNum,1,1,2).merge()
        .setBackground('#374151').setFontColor('#FFFFFF')
        .setFontWeight('bold').setFontSize(11)
        .setHorizontalAlignment('left');
    } else if (label === 'Product' || label === 'Date' || label === 'Bartender') {
      // Sub-header rows
      dash.getRange(rowNum,1,1,2)
        .setBackground('#f3f4f6').setFontWeight('bold');
    } else if (label === 'Revenue') {
      dash.getRange(rowNum, 2, 1, 1).setNumberFormat('#,##0" RWF"');
    } else if (label === 'Last Updated') {
      dash.getRange(rowNum, 1, 1, 2)
        .setBackground('#f8fafc').setFontStyle('italic');
    }
  });

  // Red highlight rows where remaining <= 3
  rows.forEach((row, i) => {
    if (row[0] && lowStock.some(ls => ls[0] === row[0] && Number(ls[1]) <= 3)) {
      dash.getRange(i+1, 1, 1, 2).setBackground('#fee2e2');
    }
  });

  dash.setColumnWidth(1, 220);
  dash.setColumnWidth(2, 180);
  dash.setFrozenRows(1);
}

// ── HELPER ────────────────────────────────────────────────────────
function getOrCreateSheet(wb, name) {
  let sheet = wb.getSheetByName(name);
  if (!sheet) sheet = wb.insertSheet(name);
  return sheet;
}