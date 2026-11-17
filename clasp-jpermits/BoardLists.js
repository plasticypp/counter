/**
 * Read-only lists the Dispatch board shows beside Open and Completed:
 *   - Test Reports: the TestReports register, newest first, with each PDF's link.
 *   - Products: the Masters sheet WITHOUT its costing columns.
 *
 * Columns are resolved by header name, never by index: both sheets have had columns moved.
 */

/**
 * The only Masters columns the board may see. The sheet also carries the cost build-up
 * (polymer, labour, finance, profit, rate ...). The board is a public link on the shop
 * floor, so this is an allow-list: a costing column added to Masters later stays hidden.
 */
var BOARD_MASTER_FIELDS_ = [
  ['code', 'Product Code'], ['name', 'APL Item Name'], ['shape', 'SHAPE'],
  ['size', 'PACK SIZE'], ['weight', 'TARE WEIGHT in gms'], ['color', 'COLOR'],
  ['machine', 'Machine'], ['mb', 'Masterbatch'],
  ['perHour', 'Output per Hour'], ['perDay', 'Output per Day'],
];

function boardHeaderIndex_(headers) {
  var idx = {};
  headers.forEach(function (h, i) {
    var key = String(h).trim().toLowerCase();
    if (key && !(key in idx)) idx[key] = i;
  });
  return idx;
}

/** Products for the board, from Masters values already read by getDashboardData_. */
function getBoardProducts_(mastersValues) {
  if (!mastersValues || mastersValues.length < 2) return [];
  var idx = boardHeaderIndex_(mastersValues[0]);
  var cols = BOARD_MASTER_FIELDS_.map(function (f) {
    return [f[0], idx[f[1].toLowerCase()]];
  });
  var codeCol = idx['product code'];
  if (codeCol === undefined) return [];

  return mastersValues.slice(1)
    .filter(function (r) { return String(r[codeCol]).trim() !== ''; })
    .map(function (r) {
      var o = {};
      cols.forEach(function (c) { o[c[0]] = c[1] === undefined ? '' : boardCell_(r[c[1]]); });
      return o;
    });
}

/** Test reports, newest first. Rows without a finished PDF keep url '' (or the failure). */
function getBoardTestReports_(ss) {
  var sh = ss.getSheetByName(TR_REGISTER_SHEET_);
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getDataRange().getValues();
  var idx = boardHeaderIndex_(values[0]);
  var col = function (name) { return idx[name.toLowerCase()]; };
  var c = {
    no: col('Report No'), date: col('Date'), invoice: col('Invoice No'),
    po: col('PO Number'), code: col('Code'), item: col('Item'), qty: col('Qty'),
    url: col('File URL'),
  };

  var out = values.slice(1)
    .filter(function (r) { return String(r[c.no]).trim() !== ''; })
    .map(function (r) {
      var url = String(r[c.url] || '').trim();
      var date = r[c.date] instanceof Date ? r[c.date] : null;
      return {
        no: String(r[c.no]), date: date ? date.toISOString() : String(r[c.date] || ''),
        invoice: String(r[c.invoice] || ''), po: String(r[c.po] || ''),
        code: String(r[c.code] || ''), item: String(r[c.item] || ''),
        qty: Number(r[c.qty]) || 0,
        url: /^https:\/\//.test(url) ? url : '',
        failed: /^FAILED/i.test(url),
      };
    });
  out.reverse();   // register is appended in order, so reversing gives newest first
  return out;
}

function boardCell_(v) {
  if (v instanceof Date) return v.toISOString();
  return typeof v === 'number' ? v : String(v).trim();
}

/**
 * Make the whole "Test Reports" folder viewable by anyone with the link, so the board's
 * "Open PDF" works for people on the floor who are not signed in. Files inherit the folder's
 * link access, including future financial-year subfolders. Idempotent.
 */
function shareTestReportsFolder_(ss) {
  var folder = getOrCreateSubfolder_(ss, 'Test Reports');
  if (folder.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK ||
      folder.getSharingPermission() !== DriveApp.Permission.VIEW) {
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }
  return { folder: folder.getName(), url: folder.getUrl(), access: String(folder.getSharingAccess()) };
}
