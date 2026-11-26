var OWNER_EMAIL = 'caps.ilhom@gmail.com';
var STATUSES = ['Accepted', 'Printing', 'Ready', 'Delivered', 'Cancelled'];
var ORDER_PAY = ['Unpaid', 'Partly paid', 'Paid'];
var MAX_UPLOAD_MB = 500;                // largest file a customer may attach
var BLOCKED = /\.(exe|bat|cmd|com|scr|msi|js|vbs|ps1|sh|apk|jar|dll)$/i;

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
  ['staff_discount_percent', 30, 'Discount in Staff print log (percent)'],
  // Contact numbers shown to the customer after an order (with country code). Leave empty if not needed.
  ['contact_phone_1', "'+992501810696", 'Contact number 1 (with country code)'],
  ['contact_phone_2', '', 'Contact number 2 (optional)'],
  ['contact_phone_3', '', 'Contact number 3 (optional)'],
  ['contact_telegram', '', 'Telegram username for the Telegram button, e.g. ilhom_djumaev (optional; otherwise number 1 is used)']
];

var ORDER_HEADER = ['No.', 'Received', 'Status', 'Name', 'Phone (WhatsApp/Telegram)', 'Deadline', 'Binding',
  'Printing', 'Pages per copy', 'Copies', 'Total TJS', 'Payment', 'Paid TJS', 'Paid at', 'Files', 'Comment'];
var LOG_HEADER = ['Saved at', 'Work by', 'Date', 'For', 'Discount', 'Binding', 'Printing', 'Pages per copy',
  'Copies', 'Total sheets', 'Total TJS', 'Payment', 'Paid at', 'ID'];

// ---------- helpers ----------
function clean(v) { return String(v || '').substring(0, 200); }
function cleanLong(v) { return String(v || '').substring(0, 800); }
function cleanComment(v) { return String(v || '').substring(0, 700); }
function esc(v) { return String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function round2(n) { return Math.round((Number(n) || 0) * 100) / 100; }
function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function fmtDate(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '');
}
function monthOf(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM');
  return String(v || '').substring(0, 7);
}
function withLock(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// One code for staff access, print log, payments and staff report: script property REPORT_CODE
function checkCode(code) {
  var c = PropertiesService.getScriptProperties().getProperty('REPORT_CODE');
  var ok = !!c && String(code || '') === c;
  if (!ok) Utilities.sleep(800);
  return ok;
}

// ---------- sheets (columns are found by header name, so you may reorder them) ----------
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
function colOf(sh, name) {
  var cols = Math.max(sh.getLastColumn(), 1);
  var h = sh.getRange(1, 1, 1, cols).getValues()[0];
  var i = h.indexOf(name);
  if (i >= 0) return i + 1;
  sh.getRange(1, cols + 1).setValue(name).setFontWeight('bold');
  return cols + 1;
}
function ensureCols(sh, names) { names.forEach(function (n) { colOf(sh, n); }); }
function table(sh) {
  var last = sh.getLastRow(), cols = sh.getLastColumn();
  var vals = last ? sh.getRange(1, 1, last, cols).getValues() : [[]];
  var h = vals[0].map(String);
  var rows = [];
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === 'TOTAL') continue; // the totals row is not an order
    var o = { _row: i + 1 };
    h.forEach(function (k, j) { if (k) o[k] = vals[i][j]; });
    rows.push(o);
  }
  return { h: h, rows: rows };
}
function appendObj(sh, obj) {
  var cols = sh.getLastColumn();
  var h = sh.getRange(1, 1, 1, cols).getValues()[0];
  sh.appendRow(h.map(function (k) { return obj.hasOwnProperty(k) ? obj[k] : ''; }));
  return sh.getLastRow();
}
function setCells(sh, row, obj) {
  Object.keys(obj).forEach(function (k) { sh.getRange(row, colOf(sh, k)).setValue(obj[k]); });
}
function ordersSheet() {
  var sh = getSheet('Orders', ORDER_HEADER);
  ensureCols(sh, ['Payment', 'Paid TJS', 'Paid at', 'Files', 'Comment']);
  return sh;
}
function logSheet() {
  var sh = getSheet('Print Log', LOG_HEADER);
  ensureCols(sh, ['Payment', 'Paid at', 'ID']);
  return sh;
}
function dropdown(sh, row, name, list) {
  sh.getRange(row, colOf(sh, name)).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(list, true).build());
}

// ---------- totals row at the end of a sheet (accounting style, recalculated automatically) ----------
function colLetter(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
function removeTotals(sh) {
  var last = sh.getLastRow();
  if (last < 2) return;
  var col = sh.getRange(1, 1, last, 1).getValues();
  for (var i = col.length - 1; i >= 1; i--) if (String(col[i][0]) === 'TOTAL') sh.deleteRow(i + 1);
}
function writeTotals(sh, kind) {
  var n = sh.getLastRow(), t = n + 1;                 // data rows are 2..n, the totals row is right after
  var cell = function (name) { return sh.getRange(t, colOf(sh, name)); };
  var L = function (name) { return colLetter(colOf(sh, name)); };
  var rng = function (name) { return L(name) + '2:' + L(name) + n; };
  if (kind === 'orders') {
    cell('Name').setValue('Total amount');
    if (n >= 2) {
      cell('Total TJS').setFormula('=SUMIF(' + rng('Status') + ',"<>Cancelled",' + rng('Total TJS') + ')');
      cell('Paid TJS').setFormula('=SUM(' + rng('Paid TJS') + ')');
      cell('Payment').setFormula('="Outstanding: "&TEXT(' + L('Total TJS') + t + '-' + L('Paid TJS') + t + ',"0.00")');
    } else {
      cell('Total TJS').setValue(0); cell('Paid TJS').setValue(0);
    }
  } else {
    cell('Work by').setValue('Total amount');
    if (n >= 2) {
      cell('Total sheets').setFormula('=SUM(' + rng('Total sheets') + ')');
      cell('Total TJS').setFormula('=SUM(' + rng('Total TJS') + ')');
      cell('Payment').setFormula('="Unpaid: "&TEXT(SUMIF(' + rng('Payment') + ',"<>Paid",' + rng('Total TJS') + '),"0.00")');
    } else {
      cell('Total sheets').setValue(0); cell('Total TJS').setValue(0);
    }
  }
  sh.getRange(t, 1).setValue('TOTAL');
  sh.getRange(t, 1, 1, sh.getLastColumn()).setFontWeight('bold').setBackground('#E8EEF5');
}
function refreshTotals(sh, kind) { removeTotals(sh); writeTotals(sh, kind); }

// ---------- files attached by customers go to a Drive folder ----------
function orderFolder() {
  var it = DriveApp.getFoldersByName('Goodfield Orders');
  return it.hasNext() ? it.next() : DriveApp.createFolder('Goodfield Orders');
}
function safeName(n) { return String(n || 'file').replace(/[^\w.\- ()\u0400-\u04FF]+/g, '_').substring(0, 120); }

// Large files arrive in pieces of about 5 MB: up_start opens a resumable Drive upload,
// every up_chunk sends the next piece to it, and the last piece completes the file.
function upStart(b) {
  var name = safeName(b.name), size = Number(b.size) || 0;
  if (BLOCKED.test(name)) return jsonOut({ ok: false, error: 'type' });
  if (!(size > 0) || size > MAX_UPLOAD_MB * 1024 * 1024) return jsonOut({ ok: false, error: 'size' });
  var mime = String(b.mime || 'application/octet-stream');
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  var res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&fields=id', {
    method: 'post',
    contentType: 'application/json; charset=UTF-8',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'X-Upload-Content-Type': mime },
    payload: JSON.stringify({ name: stamp + '_' + name, parents: [orderFolder().getId()] }),
    muteHttpExceptions: true
  });
  var h = res.getAllHeaders(), loc = h['Location'] || h['location'];
  if (res.getResponseCode() !== 200 || !loc) return jsonOut({ ok: false, error: 'init' });
  var id = Utilities.getUuid();
  CacheService.getScriptCache().put('up_' + id, JSON.stringify({ url: loc, size: size, mime: mime }), 21600);
  return jsonOut({ ok: true, id: id });
}
function upChunk(b) {
  var raw = CacheService.getScriptCache().get('up_' + b.id);
  if (!raw) return jsonOut({ ok: false, error: 'session' });
  var s = JSON.parse(raw), start = Number(b.start);
  var bytes = Utilities.base64Decode(String(b.data || ''));
  if (!(start >= 0) || !bytes.length || start + bytes.length > s.size) return jsonOut({ ok: false, error: 'range' });
  var res = UrlFetchApp.fetch(s.url, {
    method: 'put',
    contentType: s.mime,
    payload: Utilities.newBlob(bytes, s.mime),
    followRedirects: false,
    muteHttpExceptions: true,
    headers: { 'Content-Range': 'bytes ' + start + '-' + (start + bytes.length - 1) + '/' + s.size }
  });
  var code = res.getResponseCode();
  if (code === 308) return jsonOut({ ok: true, done: false });
  if (code === 200 || code === 201) {
    CacheService.getScriptCache().remove('up_' + b.id);
    return jsonOut({ ok: true, done: true, url: 'https://drive.google.com/file/d/' + JSON.parse(res.getContentText()).id + '/view' });
  }
  return jsonOut({ ok: false, error: 'put', code: code });
}
function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return jsonOut({ ok: false, error: 'bad' }); }
  if (body && body.type === 'up_start') return upStart(body);
  if (body && body.type === 'up_chunk') return upChunk(body);
  return jsonOut({ ok: false, error: 'type' });
}

// ---------- prices ----------
function pricesJson() {
  var sh = getSheet('Prices', ['Key', 'Value', 'Description']);
  var have = {};
  sh.getDataRange().getValues().slice(1).forEach(function (r) { have[String(r[0])] = true; });
  DEFAULT_PRICES.forEach(function (d) { if (!have[d[0]]) sh.appendRow(d); });
  var out = {};
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    if (r[0] === '' || String(r[0]).indexOf('contact_') === 0) return; // contact numbers are not published here
    var n = Number(String(r[1]).replace(',', '.'));
    out[String(r[0])] = isNaN(n) ? r[1] : n;
  });
  return jsonOut(out);
}

// ---------- contact buttons on the order confirmation page ----------
function contactsHtml(num, name) {
  var sh = getSheet('Prices', ['Key', 'Value', 'Description']);
  var map = {};
  sh.getDataRange().getValues().slice(1).forEach(function (r) { map[String(r[0])] = r[1]; });
  var phones = ['contact_phone_1', 'contact_phone_2', 'contact_phone_3']
    .map(function (k) { return String(map[k] || '').replace(/\D/g, ''); })
    .filter(function (d) { return d.length >= 8; });
  if (!phones.length) return '';
  var msg = encodeURIComponent('Hello! My order is #' + num + ' (' + name + ').');
  var tg = String(map['contact_telegram'] || '').replace(/^@/, '').replace(/[^A-Za-z0-9_]/g, '');
  var base = 'display:block;margin:8px 0;padding:14px;border-radius:12px;text-decoration:none;font-weight:600;color:#fff;';
  var green = base + 'background:#3F7D31;', navy = base + 'background:#1B3358;';
  var h = '<p style="margin:22px 0 6px;font-weight:600">Contact us about your order</p>';
  phones.forEach(function (d) {
    h += '<div style="margin:10px 0;padding:12px;border:1px solid #ddd;border-radius:14px">' +
      '<div style="font-weight:600;margin-bottom:6px">+' + d + '</div>' +
      '<a target="_top" style="' + green + '" href="tel:+' + d + '">Call</a>' +
      '<a target="_top" style="' + green + '" href="https://wa.me/' + d + '?text=' + msg + '">WhatsApp</a></div>';
  });
  h += '<a target="_top" style="' + navy + '" href="' + (tg ? 'https://t.me/' + tg : 'https://t.me/+' + phones[0]) + '">Telegram</a>';
  return h;
}

// ---------- payments: customer orders ----------
function ordersJson(p) {
  var wantPaid = p.filter === 'paid';
  var left = 0, out = [];
  table(ordersSheet()).rows.reverse().forEach(function (o) {
    if (o['No.'] === '' || o['No.'] === undefined) return;
    var total = Number(o['Total TJS']) || 0, paid = Number(o['Paid TJS']) || 0;
    var pay = String(o['Payment'] || 'Unpaid');
    var isPaid = pay === 'Paid';
    if (wantPaid !== isPaid) return;
    if (!wantPaid && String(o['Status']) === 'Cancelled') return;
    if (!wantPaid) left += Math.max(0, total - paid);
    out.push({ no: o['No.'], name: String(o['Name'] || ''), status: String(o['Status'] || ''),
      total: round2(total), paid: round2(paid), payment: pay, paidAt: fmtDate(o['Paid at']) });
  });
  return jsonOut({ orders: out.slice(0, 80), totalLeft: round2(left) });
}

function payOrder(p) {
  var amount = Number(String(p.amount).replace(',', '.'));
  if (!(amount > 0)) return jsonOut({ ok: false, error: 'amount' });
  return withLock(function () {
    var sh = ordersSheet();
    var o = table(sh).rows.filter(function (r) { return String(r['No.']) === String(p.no); })[0];
    if (!o) return jsonOut({ ok: false, error: 'notfound' });
    var total = Number(o['Total TJS']) || 0;
    var paid = round2((Number(o['Paid TJS']) || 0) + amount);
    var full = paid >= total - 0.005;
    setCells(sh, o._row, { 'Paid TJS': paid, 'Payment': full ? 'Paid' : 'Partly paid', 'Paid at': new Date() });
    dropdown(sh, o._row, 'Payment', ORDER_PAY);
    return jsonOut({ ok: true, paid: paid, payment: full ? 'Paid' : 'Partly paid' });
  });
}

// ---------- payments: staff debts (grouped by the person in "Work by") ----------
function debtsJson() {
  return withLock(function () {
    var sh = logSheet();
    var idCol = colOf(sh, 'ID');
    var t = table(sh);
    t.rows.forEach(function (o) {
      if (!o['ID']) { o['ID'] = Utilities.getUuid().substring(0, 8); sh.getRange(o._row, idCol).setValue(o['ID']); }
    });
    var people = {}, totalUnpaid = 0;
    t.rows.forEach(function (o) {
      if (String(o['Payment']) === 'Paid') return;
      var who = String(o['Work by'] || '').trim();
      if (!who) return;
      var key = who.toLowerCase();
      var P = people[key] = people[key] || { name: who, key: key, unpaid: 0, months: {} };
      var m = monthOf(o['Date'] || o['Saved at']);
      var M = P.months[m] = P.months[m] || { month: m, unpaid: 0, count: 0, records: [] };
      var tot = Number(o['Total TJS']) || 0;
      P.unpaid += tot; M.unpaid += tot; M.count += 1; totalUnpaid += tot;
      M.records.push({ id: o['ID'], date: fmtDate(o['Date']), forwho: String(o['For'] || ''), total: round2(tot) });
    });
    var list = Object.keys(people).sort().map(function (k) {
      var P = people[k];
      var months = Object.keys(P.months).sort().reverse().map(function (m) {
        var M = P.months[m]; M.unpaid = round2(M.unpaid); return M;
      });
      return { name: P.name, key: P.key, unpaid: round2(P.unpaid), months: months };
    });
    return jsonOut({ people: list, totalUnpaid: round2(totalUnpaid) });
  });
}

function payMonth(p) {
  return withLock(function () {
    var sh = logSheet();
    var key = String(p.who || '').toLowerCase();
    var n = 0, sum = 0;
    table(sh).rows.forEach(function (o) {
      if (String(o['Payment']) === 'Paid') return;
      if (String(o['Work by'] || '').trim().toLowerCase() !== key) return;
      if (monthOf(o['Date'] || o['Saved at']) !== String(p.month)) return;
      setCells(sh, o._row, { 'Payment': 'Paid', 'Paid at': new Date() });
      n += 1; sum += Number(o['Total TJS']) || 0;
    });
    return jsonOut({ ok: n > 0, count: n, total: round2(sum) });
  });
}

function payRecord(p) {
  return withLock(function () {
    var sh = logSheet();
    var o = table(sh).rows.filter(function (r) { return String(r['ID']) === String(p.id); })[0];
    if (!o) return jsonOut({ ok: false, error: 'notfound' });
    setCells(sh, o._row, { 'Payment': 'Paid', 'Paid at': new Date() });
    return jsonOut({ ok: true });
  });
}

// ---------- monthly report page (needs the staff code) ----------
function reportPage(p) {
  if (!checkCode(p.code)) {
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:30px">Wrong or missing code.</p>');
  }
  var data = {};
  table(logSheet()).rows.forEach(function (o) {
    var month = monthOf(o['Date'] || o['Saved at']);
    var who = String(o['Work by'] || '?');
    data[month] = data[month] || {};
    var x = data[month][who] = data[month][who] || { n: 0, sheets: 0, tjs: 0, unpaid: 0 };
    var tot = Number(o['Total TJS']) || 0;
    x.n += 1; x.sheets += Number(o['Total sheets']) || 0; x.tjs += tot;
    if (String(o['Payment']) !== 'Paid') x.unpaid += tot;
  });
  var html = '<div style="font-family:sans-serif;max-width:680px;margin:0 auto;padding:20px"><h2>Staff print report</h2>';
  var months = Object.keys(data).sort().reverse();
  if (!months.length) html += '<p>No records yet.</p>';
  months.forEach(function (m) {
    html += '<h3>' + esc(m) + '</h3><table style="width:100%;border-collapse:collapse">' +
      '<tr style="text-align:left;border-bottom:2px solid #ccc"><th>Name</th><th>Records</th><th>Sheets</th><th>TJS</th><th>Unpaid</th></tr>';
    var tn = 0, ts = 0, tt = 0, tu = 0;
    Object.keys(data[m]).sort().forEach(function (w) {
      var x = data[m][w]; tn += x.n; ts += x.sheets; tt += x.tjs; tu += x.unpaid;
      html += '<tr style="border-bottom:1px solid #eee"><td>' + esc(w) + '</td><td>' + x.n + '</td><td>' + x.sheets +
        '</td><td>' + x.tjs.toFixed(2) + '</td><td>' + x.unpaid.toFixed(2) + '</td></tr>';
    });
    html += '<tr style="font-weight:bold"><td>Total</td><td>' + tn + '</td><td>' + ts + '</td><td>' + tt.toFixed(2) +
      '</td><td>' + tu.toFixed(2) + '</td></tr></table>';
  });
  return HtmlService.createHtmlOutput(html + '</div>').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------- entry point ----------
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.type === 'prices') return pricesJson();
  if (p.type === 'auth') return jsonOut({ ok: checkCode(p.code) });
  if (p.type === 'report') return reportPage(p);

  var admin = { orders: ordersJson, debts: debtsJson, pay_order: payOrder, pay_month: payMonth, pay_record: payRecord };
  if (admin[p.type]) {
    if (!checkCode(p.code)) return jsonOut({ error: 'code' });
    return admin[p.type](p);
  }

  if (!p.name) return HtmlService.createHtmlOutput('No data.');
  var now = new Date();

  // Staff print log (needs the staff code)
  if (p.type === 'log') {
    if (!checkCode(p.code)) {
      return HtmlService.createHtmlOutput(
        '<div style="font-family:sans-serif;text-align:center;padding:40px"><h2>Wrong code.</h2></div>');
    }
    withLock(function () {
      var lsh = logSheet();
      removeTotals(lsh);
      var lrow = appendObj(lsh, {
        'Saved at': now, 'Work by': clean(p.name), 'Date': clean(p.date), 'For': clean(p.forwho),
        'Discount': clean(p.discount), 'Binding': clean(p.binding), 'Printing': cleanLong(p.printing),
        'Pages per copy': clean(p.pages), 'Copies': clean(p.copies), 'Total sheets': Number(p.sheets) || 0,
        'Total TJS': Number(p.total) || 0, 'Payment': 'Unpaid', 'Paid at': '', 'ID': Utilities.getUuid().substring(0, 8)
      });
      dropdown(lsh, lrow, 'Payment', ['Unpaid', 'Paid']);
      writeTotals(lsh, 'log');
    });
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;text-align:center;padding:40px"><h2>Print log saved.</h2></div>');
  }

  // Customer order
  var no = withLock(function () {
    var sh = ordersSheet();
    removeTotals(sh);
    var max = 0;
    table(sh).rows.forEach(function (o) { var n = Number(o['No.']); if (n > max) max = n; });
    var row = appendObj(sh, {
      'No.': max + 1, 'Received': now, 'Status': STATUSES[0], 'Name': clean(p.name),
      'Phone (WhatsApp/Telegram)': "'" + clean(p.contact), 'Deadline': clean(p.date), 'Binding': clean(p.binding),
      'Printing': cleanLong(p.printing), 'Pages per copy': clean(p.pages), 'Copies': clean(p.copies),
      'Total TJS': Number(p.total) || 0, 'Payment': 'Unpaid', 'Paid TJS': 0, 'Paid at': '',
      'Files': cleanLong(p.files).split('|').join('\n'), 'Comment': cleanComment(p.comment)
    });
    dropdown(sh, row, 'Status', STATUSES);
    dropdown(sh, row, 'Payment', ORDER_PAY);
    writeTotals(sh, 'orders');
    return max + 1;
  });
  var num = ('000' + no).slice(-4);

  var body = 'New book printing order #' + num + '\n' +
    'Name: ' + clean(p.name) + '\n' +
    'Phone (WhatsApp/Telegram): ' + clean(p.contact) + '\n' +
    'Deadline: ' + clean(p.date) + '\n' +
    'Binding: ' + clean(p.binding) + '\n' +
    'Parts / printing:\n  ' + cleanLong(p.printing).split('; ').join('\n  ') + '\n' +
    'Pages per copy: ' + clean(p.pages) + '\n' +
    'Number of copies: ' + clean(p.copies) + '\n' +
    'Total sheets: ' + clean(p.sheets) + '\n' +
    (p.comment ? 'Comment: ' + cleanComment(p.comment) + '\n' : '') +
    (p.files ? 'Files:\n  ' + cleanLong(p.files).split('|').join('\n  ') + '\n' : '') +
    'Total: ' + clean(p.total) + ' TJS (unpaid)';
  MailApp.sendEmail(OWNER_EMAIL, 'New order #' + num + ': ' + clean(p.name), body);

  return HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;max-width:420px;margin:0 auto;padding:30px 20px;text-align:center">' +
    '<h2>Order received. Thank you!</h2><p>Your order number: <b>#' + num + '</b></p>' +
    '<p>We will contact you soon.</p>' + contactsHtml(num, clean(p.name)) + '</div>')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Run once from the editor: creates the Prices, Orders and Print Log tabs (and adds the new payment columns)
function setup() {
  pricesJson();
  withLock(function () {
    refreshTotals(ordersSheet(), 'orders');
    refreshTotals(logSheet(), 'log');
  });
  orderFolder(); // creates the "Goodfield Orders" Drive folder and asks for Drive permission
  ScriptApp.getOAuthToken(); // also lets the script upload large files to Drive
}

// Optional test: sends a sample order to your email
function testOrder() {
  doGet({ parameter: {
    type: 'order', name: 'Test Client', contact: '+992 501810696', date: '2026-12-01',
    binding: 'Glue',
    printing: 'Cover: 1 p, Colour, Single-sided, Card stock; Inside: 60 p, B&W, Double-sided, Paper',
    pages: '61', copies: '50', sheets: '1550', total: '2670.00'
  }});
}
