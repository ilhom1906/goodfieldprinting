var OWNER_EMAIL = 'caps.ilhom@gmail.com';
var STATUSES = ['Accepted', 'Printing', 'Ready', 'Delivered', 'Cancelled'];

// Default prices (TJS). They are copied into the "Prices" tab the first time and can be edited there.
var DEFAULT_PRICES = [
  ['bw_single', 0.80, 'A4 black & white, single-sided (per sheet)'],
  ['bw_double', 1.50, 'A4 black & white, double-sided (per sheet)'],
  ['colour_single', 2.40, 'A4 colour, single-sided (per sheet)'],
  ['colour_double', 4.50, 'A4 colour, double-sided (per sheet)'],
  ['card_per_sheet', 1.00, 'Extra for card stock (per sheet)'],
  ['cut_a5', 2.00, 'Cutting to A5 (per copy)'],
  ['binding_glue', 5.00, 'Glue binding (per copy)'],
  ['binding_spiral', 5.00, 'Spiral binding (per copy)'],
  ['staff_discount_percent', 30, 'Discount in Staff print log (percent)']
];

function clean(v) { return String(v || '').substring(0, 200); }
function esc(v) { return String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

function getSheet(name, header) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function ordersSheet() {
  return getSheet('Orders', ['No.', 'Received', 'Status', 'Name', 'Phone (WhatsApp/Telegram)', 'Deadline',
    'Binding', 'Printing', 'Pages per copy', 'Copies', 'Total TJS']);
}
function logSheet() {
  return getSheet('Print Log', ['Saved at', 'Work by', 'Date', 'For', 'Discount', 'Binding', 'Printing',
    'Pages per copy', 'Copies', 'Total sheets', 'Total TJS']);
}

function pricesJson() {
  var sh = getSheet('Prices', ['Key', 'Value', 'Description']);
  var have = {};
  sh.getDataRange().getValues().slice(1).forEach(function (r) { have[String(r[0])] = true; });
  DEFAULT_PRICES.forEach(function (d) { if (!have[d[0]]) sh.appendRow(d); });
  var out = {};
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    if (r[0] === '') return;
    var n = Number(String(r[1]).replace(',', '.'));
    out[String(r[0])] = isNaN(n) ? r[1] : n;
  });
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function reportPage(p) {
  var code = PropertiesService.getScriptProperties().getProperty('REPORT_CODE');
  if (!code || p.code !== code) {
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:30px">Wrong or missing report code.</p>');
  }
  var rows = logSheet().getDataRange().getValues().slice(1);
  var tz = Session.getScriptTimeZone();
  var data = {};
  rows.forEach(function (r) {
    var d = r[2];
    var month = (d instanceof Date) ? Utilities.formatDate(d, tz, 'yyyy-MM') : String(d).substring(0, 7);
    var who = String(r[1] || '?');
    data[month] = data[month] || {};
    var x = data[month][who] = data[month][who] || { n: 0, sheets: 0, tjs: 0 };
    x.n += 1; x.sheets += Number(r[9]) || 0; x.tjs += Number(r[10]) || 0;
  });
  var html = '<div style="font-family:sans-serif;max-width:640px;margin:0 auto;padding:20px"><h2>Staff print report</h2>';
  var months = Object.keys(data).sort().reverse();
  if (!months.length) html += '<p>No records yet.</p>';
  months.forEach(function (m) {
    html += '<h3>' + esc(m) + '</h3><table style="width:100%;border-collapse:collapse">' +
      '<tr style="text-align:left;border-bottom:2px solid #ccc"><th>Name</th><th>Records</th><th>Sheets</th><th>TJS</th></tr>';
    var tn = 0, ts = 0, tt = 0;
    Object.keys(data[m]).sort().forEach(function (w) {
      var x = data[m][w]; tn += x.n; ts += x.sheets; tt += x.tjs;
      html += '<tr style="border-bottom:1px solid #eee"><td>' + esc(w) + '</td><td>' + x.n + '</td><td>' + x.sheets + '</td><td>' + x.tjs.toFixed(2) + '</td></tr>';
    });
    html += '<tr style="font-weight:bold"><td>Total</td><td>' + tn + '</td><td>' + ts + '</td><td>' + tt.toFixed(2) + '</td></tr></table>';
  });
  return HtmlService.createHtmlOutput(html + '</div>').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.type === 'prices') return pricesJson();
  if (p.type === 'report') return reportPage(p);
  if (!p.name) return HtmlService.createHtmlOutput('No data.');
  var now = new Date();

  // Staff print log
  if (p.type === 'log') {
    logSheet().appendRow([now, clean(p.name), clean(p.date), clean(p.forwho), clean(p.discount),
      clean(p.binding), clean(p.printing), clean(p.pages), clean(p.copies), clean(p.sheets), Number(p.total) || 0]);
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;text-align:center;padding:40px"><h2>Print log saved.</h2></div>');
  }

  // Customer order
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  var no;
  try {
    var sh = ordersSheet();
    no = sh.getLastRow();
    sh.appendRow([no, now, STATUSES[0], clean(p.name), "'" + clean(p.contact), clean(p.date), clean(p.binding),
      clean(p.printing), clean(p.pages), clean(p.copies), Number(p.total) || 0]);
    sh.getRange(sh.getLastRow(), 3).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build());
  } finally {
    lock.releaseLock();
  }
  var num = ('000' + no).slice(-4);

  var body = 'New book printing order #' + num + '\n' +
    'Name: ' + clean(p.name) + '\n' +
    'Phone (WhatsApp/Telegram): ' + clean(p.contact) + '\n' +
    'Deadline: ' + clean(p.date) + '\n' +
    'Binding: ' + clean(p.binding) + '\n' +
    'Printing: ' + clean(p.printing) + '\n' +
    'Pages per copy: ' + clean(p.pages) + '\n' +
    'Number of copies: ' + clean(p.copies) + '\n' +
    'Total: ' + clean(p.total) + ' TJS';
  MailApp.sendEmail(OWNER_EMAIL, 'New order #' + num + ': ' + clean(p.name), body);

  return HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;text-align:center;padding:40px">' +
    '<h2>Order received. Thank you!</h2><p>Your order number: <b>#' + num + '</b></p>' +
    '<p>We will contact you soon.</p></div>');
}

// Run once from the editor: creates the Prices, Orders and Print Log tabs
function setup() {
  pricesJson();
  ordersSheet();
  logSheet();
}

// Optional test: sends a sample order to your email
function testOrder() {
  doGet({ parameter: {
    type: 'order', name: 'Test Client', contact: '+992 501810696', date: '2026-12-01',
    binding: 'Glue', printing: 'B&W, Single-sided', pages: '100', copies: '1', total: '85.00'
  }});
}
