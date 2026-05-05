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
  if (action === 'LOG_SALE')          return logSale(data);
  if (action === 'END_DAY')           return endDay(data);
  if (action === 'GET_LAST_STOCK')    return getLastStock();
  if (action === 'UPDATE_PRICES')     return updatePrices(data);
  if (action === 'GET_PRICES')        return getPrices();
  if (action === 'CHECK_STOCK_ALERT') return checkStockAlert(data);
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
    sheet.appendRow(['Date','Time','Bartender','Product','Qty Sold','Unit Price (RWF)','Revenue (RWF)','Remaining Stock']);
    sheet.getRange(1,1,1,8).setFontWeight('bold').setBackground('#1A5276').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1,8,150);
  }
  sheet.appendRow([data.date,data.time,data.bartender,data.product,data.qty,data.unitPrice||0,data.revenue||0,data.remaining]);
  if (data.revenue > 0) sheet.getRange(sheet.getLastRow(),7).setBackground('#D5F5E3');
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
  sheet.appendRow(['Product','Unit Price (RWF)']);
  sheet.getRange(1,1,1,2).setFontWeight('bold').setBackground('#F39C12').setFontColor('#FFFFFF');
  data.prices.forEach(p => sheet.appendRow([p.product, p.price]));
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

// ── HELPER ────────────────────────────────────────────────────────
function getOrCreateSheet(wb, name) {
  let sheet = wb.getSheetByName(name);
  if (!sheet) sheet = wb.insertSheet(name);
  return sheet;
}