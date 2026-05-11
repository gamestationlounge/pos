// ═══════════════════════════════════════════════════════════════
// GAME STATION LOUNGE POS — Google Apps Script Backend v2
// Paste this entire file into Extensions > Apps Script
// ═══════════════════════════════════════════════════════════════

const SHEET_ID    = '14XuYMKu5fphq9w0ZqVANYUDQ3-hJvrTsLvYL_hqpO9E';
const ALERT_EMAIL = 'gamestationlounge9@gmail.com';

var OWNER_NAME  = 'KABWA';
var OWNER_EMAIL = 'gamestationlounge9@gmail.com';

// Items >= 10,000 RWF → alert at 2 remaining
// Items <  10,000 RWF → alert at 5 remaining
const HIGH_VALUE_THRESHOLD = 10000;
const ALERT_LOW  = 5;
const ALERT_HIGH = 2;

function sendLowStockEmail(product, remaining, bartender) {
  try {
    MailApp.sendEmail({
      to: OWNER_EMAIL,
      subject: '⚠️ LOW STOCK: ' + product + ' (' + remaining + ' left) - Game Station Lounge',
      htmlBody: '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">' +
        '<div style="background: #1a1a2e; padding: 20px; text-align: center;">' +
        '<h1 style="color: #f59e0b; margin: 0;">GAME STATION LOUNGE</h1>' +
        '<p style="color: #ffffff; margin: 5px 0;">Point of Sale System</p></div>' +
        '<div style="background: #fee2e2; padding: 15px; text-align: center;">' +
        '<h2 style="color: #dc2626; margin: 0;">⚠️ LOW STOCK ALERT</h2></div>' +
        '<div style="background: #ffffff; padding: 30px;">' +
        '<p>Dear <b>' + OWNER_NAME + '</b>,</p>' +
        '<p>This is an automatic alert from your POS system.</p>' +
        '<table style="width:100%; border-collapse: collapse;">' +
        '<tr style="background: #f8fafc;"><td style="padding:10px; border:1px solid #e2e8f0;"><b>Product</b></td>' +
        '<td style="padding:10px; border:1px solid #e2e8f0; color:#dc2626;">' + product + '</td></tr>' +
        '<tr><td style="padding:10px; border:1px solid #e2e8f0;"><b>Remaining Stock</b></td>' +
        '<td style="padding:10px; border:1px solid #e2e8f0; color:#dc2626; font-size:20px;"><b>' + remaining + ' units</b></td></tr>' +
        '<tr style="background:#f8fafc;"><td style="padding:10px; border:1px solid #e2e8f0;"><b>Sold by</b></td>' +
        '<td style="padding:10px; border:1px solid #e2e8f0;">' + bartender + '</td></tr>' +
        '<tr><td style="padding:10px; border:1px solid #e2e8f0;"><b>Time</b></td>' +
        '<td style="padding:10px; border:1px solid #e2e8f0;">' + new Date().toLocaleString() + '</td></tr></table>' +
        '<div style="background:#fef3c7; padding:15px; margin-top:20px; border-radius:5px;">' +
        '<p style="margin:0; color:#92400e;"><b>⚡ Please restock ' + product + ' as soon as possible!</b></p></div></div>' +
        '<div style="background:#1a1a2e; padding:15px; text-align:center;">' +
        '<p style="color:#9ca3af; margin:0; font-size:12px;">Game Station Lounge POS - Automatic Alert System</p></div></div>'
    });
  } catch(e) {
    Logger.log('Email alert failed: ' + e);
  }
}

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
  if (action === 'SAVE_CREDIT')         return saveCredit(data);
  if (action === 'GET_CREDITS')         return getCredits(data);
  if (action === 'MARK_CREDIT_PAID')    return markCreditPaid(data);
  if (action === 'GET_MANAGER_REPORT')  return getManagerReport(data);
  if (action === 'GET_STOCK_LIVE')      return getStockLive();
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
  if (Number(data.remaining) <= 10) {
    sendLowStockEmail(data.product, data.remaining, data.bartender || 'Bartender');
  }
  return { ok: true };
}

// ── 1b. SAVE TAB (one row per product) ───────────────────────────
function saveTab(data) {
  const wb    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(wb, 'Sales Log');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Date','Time','Bartender','Table','Product','Qty','Unit Price (RWF)','Revenue (RWF)','Payment']);
    sheet.getRange(1,1,1,9).setFontWeight('bold').setBackground('#1A5276').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1,9,150);
  }

  // Parse items JSON and write one row per product
  let items = [];
  try { items = JSON.parse(data.items); } catch(e) { items = []; }

  items.forEach(function(item) {
    const lineTotal = (item.qty || 0) * (item.unitPrice || 0);
    sheet.appendRow([
      data.date, data.time, data.bartender, data.tableName,
      item.product, item.qty || 0, item.unitPrice || 0,
      lineTotal, data.paymentMethod
    ]);
    if (lineTotal > 0) sheet.getRange(sheet.getLastRow(), 8).setBackground('#D5F5E3');
  });

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

// ── 7. SAVE CREDIT ────────────────────────────────────────────────
function saveCredit(data) {
  const wb    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(wb, 'Credits');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Date','Time','Bartender','Customer Name','Phone','Table','Items','Amount (RWF)','Status','Paid Date','Paid Method','Credit ID']);
    sheet.getRange(1,1,1,12).setFontWeight('bold').setBackground('#f59e0b').setFontColor('#000000');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1,12,130);
  }

  sheet.appendRow([
    data.date, data.time, data.bartender,
    data.customerName, data.phone || '', data.tableName,
    data.items, data.amount, 'PENDING', '', '', data.creditId || ''
  ]);

  // Highlight PENDING row in light red
  const row = sheet.getLastRow();
  sheet.getRange(row, 1, 1, 12).setBackground('#fee2e2');
  sheet.autoResizeColumns(1, 12);

  // Send email alert to owner
  sendCreditEmail(data);

  return { ok: true };
}

// ── 8. GET CREDITS ─────────────────────────────────────────────────
function getCredits(data) {
  const wb    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = wb.getSheetByName('Credits');
  if (!sheet || sheet.getLastRow() <= 1) return { ok: true, credits: [] };

  const rows = sheet.getRange(2, 1, sheet.getLastRow()-1, 12).getValues();
  const statusFilter = data.status || 'ALL';

  const credits = rows
    .filter(r => r[0] !== '')
    .filter(r => statusFilter === 'ALL' || r[8] === statusFilter)
    .map(r => ({
      date: r[0], time: r[1], bartender: r[2],
      customerName: r[3], phone: r[4], tableName: r[5],
      items: r[6], amount: r[7], status: r[8],
      paidDate: r[9], paidMethod: r[10], creditId: r[11]
    }));

  return { ok: true, credits };
}

// ── 9. MARK CREDIT PAID ───────────────────────────────────────────
function markCreditPaid(data) {
  const wb    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = wb.getSheetByName('Credits');
  if (!sheet || sheet.getLastRow() <= 1) return { ok: false, error: 'Credits sheet not found' };

  const rows = sheet.getRange(2, 1, sheet.getLastRow()-1, 12).getValues();
  let updated = false;

  for (let i = 0; i < rows.length; i++) {
    // Match by creditId (col L, index 11), fall back to PENDING status
    const matchById = data.creditId && rows[i][11] === data.creditId;
    const matchFallback = !data.creditId && rows[i][8] === 'PENDING';
    if (matchById || matchFallback) {
      const sheetRow = i + 2;
      sheet.getRange(sheetRow, 9).setValue('PAID');
      sheet.getRange(sheetRow, 10).setValue(data.paidDate);
      sheet.getRange(sheetRow, 11).setValue(data.paidMethod);
      sheet.getRange(sheetRow, 1, 1, 12).setBackground('#dcfce7');
      updated = true;
      break;
    }
  }

  sheet.autoResizeColumns(1, 12);
  return { ok: true, updated };
}

// ── CREDIT EMAIL ALERT ────────────────────────────────────────────
function sendCreditEmail(data) {
  try {
    const itemsList = (function() {
      try { return JSON.parse(data.items).join('<br>  · '); } catch(e) { return data.items; }
    })();

    const subject = '📒 NEW CREDIT: ' + data.customerName + ' owes ' + Number(data.amount).toLocaleString() + ' RWF';

    const htmlBody = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#ffffff;border:1px solid #e5e7eb;border-radius:12px;overflow:hidden">' +

      '<div style="background:#1a1a2e;padding:22px 28px;text-align:center">' +
      '<h2 style="color:#f59e0b;margin:0;font-size:20px;letter-spacing:3px">GAME STATION LOUNGE</h2>' +
      '<p style="color:#94a3b8;margin:5px 0 0;font-size:11px;letter-spacing:2px">POINT OF SALE SYSTEM</p></div>' +

      '<div style="background:#fef3c7;border-bottom:2px solid #f59e0b;padding:16px 28px;text-align:center">' +
      '<p style="font-size:32px;margin:0">📒</p>' +
      '<h3 style="color:#92400e;margin:6px 0 0;font-size:17px;letter-spacing:1px">NEW CREDIT RECORDED</h3></div>' +

      '<div style="padding:28px">' +
      '<p style="color:#111827;font-size:15px;margin:0 0 20px">Dear <strong>KABWA</strong>,</p>' +
      '<p style="color:#374151;font-size:14px;margin:0 0 24px">A new credit has been recorded at Game Station Lounge.</p>' +

      '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:24px">' +
      '<tr style="background:#f8fafc"><td style="padding:11px 16px;border-bottom:1px solid #e5e7eb;color:#64748b;width:35%">Customer</td>' +
      '<td style="padding:11px 16px;border-bottom:1px solid #e5e7eb;font-weight:bold;color:#111827">' + data.customerName + '</td></tr>' +
      (data.phone ? '<tr><td style="padding:11px 16px;border-bottom:1px solid #e5e7eb;color:#64748b;background:#f8fafc">Phone</td>' +
      '<td style="padding:11px 16px;border-bottom:1px solid #e5e7eb;color:#111827">' + data.phone + '</td></tr>' : '') +
      '<tr style="background:#f8fafc"><td style="padding:11px 16px;border-bottom:1px solid #e5e7eb;color:#64748b">Amount Owed</td>' +
      '<td style="padding:11px 16px;border-bottom:1px solid #e5e7eb;font-weight:bold;color:#dc2626;font-size:18px">' + Number(data.amount).toLocaleString() + ' RWF</td></tr>' +
      '<tr><td style="padding:11px 16px;border-bottom:1px solid #e5e7eb;color:#64748b;background:#f8fafc">Table</td>' +
      '<td style="padding:11px 16px;border-bottom:1px solid #e5e7eb;color:#111827">' + data.tableName + '</td></tr>' +
      '<tr style="background:#f8fafc"><td style="padding:11px 16px;border-bottom:1px solid #e5e7eb;color:#64748b">Bartender</td>' +
      '<td style="padding:11px 16px;border-bottom:1px solid #e5e7eb;color:#111827">' + data.bartender + '</td></tr>' +
      '<tr><td style="padding:11px 16px;border-bottom:1px solid #e5e7eb;color:#64748b;background:#f8fafc">Date & Time</td>' +
      '<td style="padding:11px 16px;border-bottom:1px solid #e5e7eb;color:#111827">' + data.date + ' at ' + data.time + '</td></tr>' +
      '<tr style="background:#f8fafc"><td style="padding:11px 16px;color:#64748b;vertical-align:top">Items Consumed</td>' +
      '<td style="padding:11px 16px;color:#111827;font-size:12px">· ' + itemsList + '</td></tr></table>' +

      '<div style="background:#fef2f2;border:1px solid #fca5a5;border-radius:8px;padding:14px 18px;text-align:center">' +
      '<p style="color:#991b1b;font-weight:bold;margin:0;font-size:13px">⚠️ Please follow up with the customer to collect payment.</p></div></div>' +

      '<div style="background:#f1f5f9;padding:14px 28px;text-align:center;border-top:1px solid #e5e7eb">' +
      '<p style="color:#94a3b8;font-size:11px;margin:0">Game Station Lounge POS — Credit Alert System</p></div></div>';

    MailApp.sendEmail({ to: OWNER_EMAIL, subject, htmlBody });
  } catch(e) {
    Logger.log('Credit email failed: ' + e);
  }
}

// ── WEEKLY CREDIT SUMMARY EMAIL ───────────────────────────────────
// To enable weekly Monday emails, run setupWeeklyCreditTrigger() once
// from the Apps Script editor (Run > setupWeeklyCreditTrigger).
function setupWeeklyCreditTrigger() {
  // Remove any existing weekly credit triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendWeeklyCreditSummary') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendWeeklyCreditSummary')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
}

function sendWeeklyCreditSummary() {
  try {
    const wb    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = wb.getSheetByName('Credits');
    if (!sheet || sheet.getLastRow() <= 1) return;

    const rows = sheet.getRange(2, 1, sheet.getLastRow()-1, 11).getValues();
    const pending = rows.filter(r => r[0] !== '' && r[8] === 'PENDING');
    const totalOwed = pending.reduce((a, r) => a + Number(r[7]), 0);

    const rows_html = pending.map(r =>
      '<tr style="border-bottom:1px solid #e5e7eb">' +
      '<td style="padding:9px 12px;color:#111827">' + r[3] + '</td>' +
      '<td style="padding:9px 12px;color:#64748b">' + r[4] + '</td>' +
      '<td style="padding:9px 12px;color:#64748b">' + r[0] + '</td>' +
      '<td style="padding:9px 12px;color:#64748b">' + r[5] + '</td>' +
      '<td style="padding:9px 12px;font-weight:bold;color:#dc2626">' + Number(r[7]).toLocaleString() + ' RWF</td>' +
      '</tr>'
    ).join('');

    const subject = '📊 Weekly Credit Summary - Game Station Lounge';
    const htmlBody = '<div style="font-family:Arial,sans-serif;max-width:660px;margin:0 auto;background:#fff;border:1px solid #e5e7eb;border-radius:12px;overflow:hidden">' +
      '<div style="background:#1a1a2e;padding:22px 28px;text-align:center">' +
      '<h2 style="color:#f59e0b;margin:0;font-size:20px;letter-spacing:3px">GAME STATION LOUNGE</h2>' +
      '<p style="color:#94a3b8;margin:5px 0 0;font-size:11px;letter-spacing:2px">WEEKLY CREDIT SUMMARY</p></div>' +
      '<div style="padding:24px">' +
      '<p style="color:#111827;font-size:15px;margin:0 0 6px">Dear <strong>KABWA</strong>,</p>' +
      '<p style="color:#374151;font-size:13px;margin:0 0 20px">Here is the weekly outstanding credit summary for Game Station Lounge.</p>' +
      '<div style="background:#fee2e2;border:1px solid #fca5a5;border-radius:8px;padding:14px;text-align:center;margin-bottom:20px">' +
      '<div style="color:#64748b;font-size:11px;letter-spacing:.1em;text-transform:uppercase;margin-bottom:4px">Total Outstanding</div>' +
      '<div style="color:#dc2626;font-size:26px;font-weight:bold">' + totalOwed.toLocaleString() + ' RWF</div>' +
      '<div style="color:#64748b;font-size:12px">' + pending.length + ' pending credit' + (pending.length!==1?'s':'') + '</div></div>' +
      (pending.length > 0 ?
        '<table style="width:100%;border-collapse:collapse;font-size:13px">' +
        '<tr style="background:#f59e0b"><th style="padding:10px 12px;text-align:left;color:#000">Customer</th><th style="padding:10px 12px;text-align:left;color:#000">Phone</th><th style="padding:10px 12px;text-align:left;color:#000">Date</th><th style="padding:10px 12px;text-align:left;color:#000">Table</th><th style="padding:10px 12px;text-align:left;color:#000">Amount</th></tr>' +
        rows_html + '</table>' :
        '<p style="text-align:center;color:var(--accent2);font-size:14px">✅ No outstanding credits!</p>') +
      '</div>' +
      '<div style="background:#f1f5f9;padding:14px 28px;text-align:center;border-top:1px solid #e5e7eb">' +
      '<p style="color:#94a3b8;font-size:11px;margin:0">Game Station Lounge POS — Automated Weekly Report</p></div></div>';

    MailApp.sendEmail({ to: OWNER_EMAIL, subject, htmlBody });
  } catch(e) {
    Logger.log('Weekly credit summary failed: ' + e);
  }
}

// ── GET MANAGER REPORT ────────────────────────────────────────────
function getManagerReport(data) {
  var wb = SpreadsheetApp.openById(SHEET_ID);

  function parseDate(s) {
    if (!s) return null;
    if (s instanceof Date) return s;
    var p = String(s).split('/');
    if (p.length === 3) return new Date(+p[2], +p[1]-1, +p[0]);
    return new Date(s);
  }

  var startD = parseDate(data.startDate) || new Date(0);
  var endD   = parseDate(data.endDate)   || new Date();
  endD.setHours(23, 59, 59, 999);

  var totalRevenue = 0, totalTransactions = 0, totalUnits = 0;
  var cash = 0, momo = 0, credit = 0;
  var prodTotals = {}, bartenderMap = {}, dailyMap = {};

  var salesSheet = wb.getSheetByName('Sales Log');
  if (salesSheet && salesSheet.getLastRow() > 1) {
    // Sales Log columns: Date(0) | Time(1) | Bartender(2) | Table(3) | Product(4) | Qty(5) | UnitPrice(6) | Revenue(7) | Payment(8)
    var rows = salesSheet.getRange(2, 1, salesSheet.getLastRow() - 1, 9).getValues();
    rows.forEach(function(r) {
      if (!r[0]) return;
      var rowDate = r[0] instanceof Date ? r[0] : parseDate(r[0]);
      if (!rowDate || rowDate < startD || rowDate > endD) return;

      var bartender = r[2] || '—';
      var product   = r[4] || '';   // Column E — product name
      var qty       = Number(r[5]) || 0;  // Column F — quantity
      var revenue   = Number(r[7]) || 0;  // Column H — revenue
      var method    = String(r[8] || 'CASH').toUpperCase();
      var dateKey   = r[0] instanceof Date
        ? r[0].toLocaleDateString('en-GB')
        : String(r[0]);

      totalRevenue += revenue;
      totalUnits   += qty;
      totalTransactions++;

      if (method === 'CASH')        cash   += revenue;
      else if (method === 'MOMO')   momo   += revenue;
      else if (method === 'CREDIT') credit += revenue;

      if (product) prodTotals[product] = (prodTotals[product] || 0) + qty;

      if (!bartenderMap[bartender]) bartenderMap[bartender] = { name: bartender, revenue: 0, transactions: 0 };
      bartenderMap[bartender].revenue      += revenue;
      bartenderMap[bartender].transactions++;

      if (!dailyMap[dateKey]) dailyMap[dateKey] = { date: dateKey, revenue: 0, transactions: 0, cash: 0, momo: 0, credit: 0 };
      dailyMap[dateKey].revenue      += revenue;
      dailyMap[dateKey].transactions++;
      if (method === 'CASH')        dailyMap[dateKey].cash   += revenue;
      else if (method === 'MOMO')   dailyMap[dateKey].momo   += revenue;
      else if (method === 'CREDIT') dailyMap[dateKey].credit += revenue;
    });
  }

  var topProducts = Object.keys(prodTotals)
    .map(function(p) { return { product: p, sold: prodTotals[p] }; })
    .sort(function(a, b) { return b.sold - a.sold; })
    .slice(0, 10);

  var bartenders = Object.values(bartenderMap)
    .sort(function(a, b) { return b.revenue - a.revenue; });

  var dailyBreakdown = Object.values(dailyMap)
    .sort(function(a, b) {
      var da = a.date.split('/'), db = b.date.split('/');
      return new Date(+da[2], +da[1]-1, +da[0]) - new Date(+db[2], +db[1]-1, +db[0]);
    });

  return {
    ok: true,
    totalRevenue: totalRevenue,
    totalTransactions: totalTransactions,
    totalUnits: totalUnits,
    cash: cash,
    momo: momo,
    credit: credit,
    topProducts: topProducts,
    bartenders: bartenders,
    dailyBreakdown: dailyBreakdown
  };
}

// ── GET STOCK LIVE ────────────────────────────────────────────────
function getStockLive() {
  var wb    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = wb.getSheetByName('Closing Stock');
  if (!sheet || sheet.getLastRow() <= 1) return { ok: true, stock: [] };

  // Closing Stock columns: Product(A=0) | Remaining(B=1) | Date(C=2)
  var rows  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  var stock = rows
    .filter(function(r) { return r[0] !== ''; })
    .map(function(r) { return { product: String(r[0]), remaining: Number(r[1]) || 0 }; });

  return { ok: true, stock: stock };
}

// ── HELPER ────────────────────────────────────────────────────────
function getOrCreateSheet(wb, name) {
  let sheet = wb.getSheetByName(name);
  if (!sheet) sheet = wb.insertSheet(name);
  return sheet;
}