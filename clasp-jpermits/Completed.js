/**
 * Completed lines for the dashboard's Completed view, with the facts OTD/OTIF need.
 *
 * The view lists EVERY line with something supplied against it (decided 2026-11-17).
 * A line is DONE when it also has neither Machine nor Priority. That includes lines that still
 * carry a positive Balance: the planner un-prioritises a line when it is closed, so the missing
 * priority, not the Balance, is the signal (decided 2026-10-15). A supplied line still on the
 * plan is sent with onPlan: true and is not judged for OTD/OTIF. A line with nothing supplied
 * is a PO waiting to be planned and is not listed here.
 *
 * Only facts are sent; the browser computes OTD/OTIF so every view agrees on one rule.
 *   due       = PO Date + Settings "OTD Lead Time (days)"           (default 20)
 *   delivered = date of the LAST invoice that shipped against the line
 *   in full   = Supplied >= Order Qty x (1 - Settings "In-Full Tolerance (%)")   (default 5)
 */
var COMPLETED_DEFAULT_LEAD_DAYS_ = 20;
var COMPLETED_DEFAULT_TOLERANCE_PCT_ = 5;

function getCompletedData_(ss, headers, rows, nameByCode) {
  var col = {
    machine: headers.indexOf('Machine'), priority: headers.indexOf('Priority'),
    po: headers.indexOf('PO Number'), poDate: headers.indexOf('PO Date'),
    code: headers.indexOf('Code'), color: headers.indexOf('Color'),
    orderQty: headers.indexOf('Order Qty'), supplied: headers.indexOf('Supplied'),
    rate: headers.indexOf('Rate'), shape: headers.indexOf('Shape'),
  };
  var invoiceCols = invoiceColumns_(headers);
  var invoiceDates = readInvoiceDates_(ss);
  recordNewInvoiceDates_(ss, invoiceCols, rows, invoiceDates);
  var settings = readCompletedSettings_(ss);
  var usualMachine = readUsualMachines_(ss);

  var blank = function (v) { return v === '' || v === null || v === undefined || String(v).trim() === ''; };
  var invoicesOf = function (r) {
    var invoices = [];
    invoiceCols.forEach(function (c) {
      var qty = Number(String(r[c.index]).replace(/,/g, ''));
      if (!qty) return;
      invoices.push({ no: c.no, date: invoiceDates[c.no] || null, qty: qty });
    });
    return invoices;
  };
  // Off the plan AND something was supplied. A PO line that has just been imported also has
  // no Machine or Priority, but nothing has shipped against it — it is waiting to be planned,
  // not completed.
  var isDone = function (r) {
    return blank(r[col.machine]) && blank(r[col.priority])
      && (Number(r[col.supplied]) > 0 || invoicesOf(r).length > 0);
  };

  // Every shipment on every line, open or completed, for the Summary tab: a month's dispatch
  // includes part-shipments against lines still on the plan. `line` is the row index, the
  // same id the completed lines carry, so a shipment can show its line's OTIF status.
  var shipments = [];
  rows.forEach(function (r, i) {
    if (!r[col.code]) return;
    invoicesOf(r).forEach(function (inv) {
      shipments.push({
        line: i, done: isDone(r), inv: inv.no, date: inv.date, qty: inv.qty,
        rate: Number(r[col.rate]) || 0, code: r[col.code],
        name: nameByCode[r[col.code]] || r[col.code], color: r[col.color] || '(no color)',
        brand: brandOf_(r, col), po: r[col.po], usualMachine: usualMachine[r[col.code]] || '',
      });
    });
  });

  // Every line with anything supplied, whether or not it is still on the plan, so every
  // dispatch is captured (decided 2026-11-17). A line still carrying a Machine or Priority is
  // marked onPlan: the browser shows it as "On plan" and leaves it out of OTD/OTIF, which
  // judge only finished lines.
  var lines = [];
  rows.forEach(function (r, i) {
    if (!r[col.code]) return;
    var invoices = invoicesOf(r);
    if (!(Number(r[col.supplied]) > 0 || invoices.length > 0)) return;
    var line = completedLine_(r, i, col, nameByCode, invoices, usualMachine);
    line.onPlan = !isDone(r);
    lines.push(line);
  });

  return {
    leadDays: settings.leadDays,
    tolerancePct: settings.tolerancePct,
    lines: lines,
    shipments: shipments,
  };
}

function completedLine_(r, id, col, nameByCode, invoices, usualMachine) {
  return {
    id: id, po: r[col.po], poDate: completedDay_(r[col.poDate]),
    code: r[col.code], name: nameByCode[r[col.code]] || r[col.code],
    color: r[col.color] || '(no color)', brand: brandOf_(r, col),
    usualMachine: usualMachine[r[col.code]] || '',
    orderQty: Number(r[col.orderQty]) || 0, supplied: Number(r[col.supplied]) || 0,
    invoices: invoices,
  };
}

// Product Code -> the machine Masters says that product normally runs on. A completed line has
// had its own Machine cleared, so this is the only machine left to filter or group it by.
function readUsualMachines_(ss) {
  var sheet = ss.getSheetByName('Masters');
  if (!sheet || sheet.getLastRow() < 2) return {};
  var values = sheet.getDataRange().getValues();
  var h = values[0].map(function (v) { return String(v).trim(); });
  var iCode = h.indexOf('Product Code'), iMachine = h.indexOf('Machine');
  if (iCode === -1 || iMachine === -1) return {};
  var map = {};
  values.slice(1).forEach(function (r) {
    var code = String(r[iCode]).trim(), m = String(r[iMachine]).trim();
    if (code && m) map[code] = m;
  });
  return map;
}

// The Shape column holds the customer brand the bottle is moulded for (VARROC, ARZOL, myTVS).
function brandOf_(r, col) {
  return col.shape === -1 ? '' : String(r[col.shape] || '').trim();
}

// Invoice columns are every named column to the right of 'Opening'. 'Opening' itself is
// opening stock, not a shipment, so it is excluded.
function invoiceColumns_(headers) {
  var opening = headers.indexOf('Opening');
  if (opening === -1) return [];
  var out = [];
  for (var i = opening + 1; i < headers.length; i++) {
    var no = String(headers[i]).trim();
    if (no) out.push({ index: i, no: no });
  }
  return out;
}

// Invoice Number -> 'yyyy-MM-dd', read by header name so a raw Zoho export pasted into the
// Invoices tab works as well as the trimmed one. Void invoices are skipped. A real date beats
// an automatic one, so pasting a Zoho export under the auto rows corrects them.
var AUTO_INVOICE_STATUS_ = 'Auto (first seen)';

function readInvoiceDates_(ss) {
  var sheet = ss.getSheetByName('Invoices');
  if (!sheet || sheet.getLastRow() < 2) return {};
  var values = sheet.getDataRange().getValues();
  var h = values[0].map(function (v) { return String(v).trim(); });
  var iNo = h.indexOf('Invoice Number'), iDate = h.indexOf('Invoice Date'), iStatus = h.indexOf('Invoice Status');
  if (iNo === -1 || iDate === -1) return {};
  var map = {}, auto = {};
  values.slice(1).forEach(function (r) {
    var no = String(r[iNo]).trim();
    var status = iStatus === -1 ? '' : String(r[iStatus]).trim();
    if (!no || status.toLowerCase() === 'void') return;
    if (map[no] && !auto[no]) return;
    var day = completedDay_(r[iDate]);
    if (!day) return;
    if (map[no] && status === AUTO_INVOICE_STATUS_) return;
    map[no] = day;
    auto[no] = status === AUTO_INVOICE_STATUS_;
  });
  return map;
}

// An invoice column with quantities but no date on the Invoices tab gets today's date, written
// to the tab as "Auto (first seen)". The board polls all day, so the first day it sees the
// column is almost always the day the invoice was typed in — close enough to place it in the
// right month until a Zoho export supplies the real date. Skipped, not waited for, if another
// request holds the lock: the next poll records it.
function recordNewInvoiceDates_(ss, invoiceCols, rows, invoiceDates) {
  var missing = invoiceCols.filter(function (c) {
    if (invoiceDates[c.no]) return false;
    return rows.some(function (r) { return Number(String(r[c.index]).replace(/,/g, '')) > 0; });
  });
  if (!missing.length) return;
  var sheet = ss.getSheetByName('Invoices');
  if (!sheet) return;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    var fresh = readInvoiceDates_(ss);
    var today = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
    var h = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (v) { return String(v).trim(); });
    var iNo = h.indexOf('Invoice Number'), iDate = h.indexOf('Invoice Date'), iStatus = h.indexOf('Invoice Status');
    if (iNo === -1 || iDate === -1) return;
    missing.forEach(function (c) {
      if (fresh[c.no]) { invoiceDates[c.no] = fresh[c.no]; return; }
      var row = h.map(function () { return ''; });
      row[iNo] = c.no; row[iDate] = today;
      if (iStatus !== -1) row[iStatus] = AUTO_INVOICE_STATUS_;
      sheet.appendRow(row);
      invoiceDates[c.no] = today;
    });
  } finally {
    lock.releaseLock();
  }
}

function readCompletedSettings_(ss) {
  var out = { leadDays: COMPLETED_DEFAULT_LEAD_DAYS_, tolerancePct: COMPLETED_DEFAULT_TOLERANCE_PCT_ };
  var sheet = ss.getSheetByName('Settings');
  if (!sheet || sheet.getLastRow() < 2) return out;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function (r) {
    var key = String(r[0]).trim().toLowerCase(), val = Number(r[1]);
    if (!isFinite(val) || String(r[1]).trim() === '') return;
    if (key.indexOf('lead time') !== -1) out.leadDays = val;
    if (key.indexOf('tolerance') !== -1) out.tolerancePct = val;
  });
  return out;
}

// A sheet date to 'yyyy-MM-dd' in IST, or null. PO Date is entered by hand in several shapes:
// real date cells, '03/09/2026', '24/08/26' (both day-first), and '01-Oct-26'.
function completedDay_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, 'Asia/Kolkata', 'yyyy-MM-dd');
  var s = String(v == null ? '' : v).trim();
  if (!s) return null;
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return isoDay_(Number(m[1]), Number(m[2]), Number(m[3]));
  m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/);
  if (m) return isoDay_(fullYear_(m[3]), Number(m[2]), Number(m[1]));
  m = s.match(/^(\d{1,2})[\-\s]([A-Za-z]{3})[A-Za-z]*[\-\s](\d{2,4})$/);
  if (m) {
    var mon = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'].indexOf(m[2].toLowerCase());
    if (mon !== -1) return isoDay_(fullYear_(m[3]), mon + 1, Number(m[1]));
  }
  return null;
}

function fullYear_(y) { return y.length === 2 ? 2000 + Number(y) : Number(y); }

function isoDay_(y, mo, d) {
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  return y + '-' + (mo < 10 ? '0' : '') + mo + '-' + (d < 10 ? '0' : '') + d;
}
