/**
 * Tarzift — receive website orders into this Google Sheet.
 * Setup: Extensions > Apps Script, paste this file, run setup() once,
 * then Deploy > New deployment > Web app (Execute as: Me, Access: Anyone).
 */
const SHEET_NAME = 'Orders';
const HEADERS = [
  'Date', 'Order ID', 'Status', 'Customer name', 'Phone', 'City', 'Address',
  'Name (Arabic)', 'Name (Latin)', 'Card message', 'Quantity', 'Unit price',
  'Delivery fee', 'Total (DH)', 'Note', 'Language', 'Page', 'Product', 'Color'
];
const FIELDS = [
  null, 'order_id', null, 'customer_name', 'phone', 'city', 'address',
  'engrave_ar', 'engrave_latin', 'card_message', 'quantity', 'unit_price',
  'delivery_fee', 'total', 'note', 'language', 'page', 'product', 'color'
];
const STATUSES = ['New', 'Confirmed', 'In production', 'Shipped', 'Delivered', 'Returned', 'Cancelled'];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#EFE3C8').setFontColor('#2B211A');
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, HEADERS.length, 140);
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build();
  sh.getRange(2, 3, sh.getMaxRows() - 1, 1).setDataValidation(rule);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const p = (e && e.parameter) || {};
    if (p.website) return json({ ok: true }); // spam trap
    if (!p.customer_name || !p.phone) return json({ ok: false, error: 'missing fields' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) { setup(); sh = ss.getSheetByName(SHEET_NAME); }

    const row = FIELDS.map((f, i) => {
      if (i === 0) return new Date();
      if (i === 2) return 'New';
      let v = String(p[f] || '').slice(0, 300);
      if (/^[=+\-@]/.test(v)) v = "'" + v; // block formula injection
      return v;
    });
    // Keep phone as text so the leading 0 stays
    row[4] = "'" + String(p.phone || '').replace(/[^\d+]/g, '');
    sh.appendRow(row);
    return json({ ok: true, order_id: p.order_id || '' });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json({ ok: true, service: 'Tarzift orders' });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
