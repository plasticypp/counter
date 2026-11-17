function getSpreadsheet_() {
  var cache = PropertiesService.getScriptProperties();
  var id = cache.getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);

  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    cache.setProperty('SPREADSHEET_ID', active.getId());
    return active;
  }

  var file = DriveApp.getFileById(ScriptApp.getScriptId()).getParents();
  if (file.hasNext()) {
    var parent = file.next();
    cache.setProperty('SPREADSHEET_ID', parent.getId());
    return SpreadsheetApp.openById(parent.getId());
  }

  throw new Error('No spreadsheet found: set SPREADSHEET_ID script property manually.');
}

function onOpen() {
  // Only the actions used day to day. The one-time setup routines (code dropdown, operators
  // sheet, trigger install) are still defined below and runnable from the script editor —
  // they are off the menu because they have already been run, and an accidental re-run is
  // the kind of mistake that is tedious to undo.
  SpreadsheetApp.getUi()
    .createMenu('Dispatch Tools')
    .addItem('Import PO from PDF…', 'showPoImportDialog')
    // The old "Print Spec" item is off the menu — the Test Report replaces it. Its functions
    // (printSpecForActiveInvoiceColumn / generateSpecPdfForInvoice_) are left defined and are
    // still runnable from the Apps Script editor, so this is a one-line restore if needed.
    .addItem('Print Test Report for Selected Invoice Column', 'printTestReportForActiveInvoiceColumn')
    .addSeparator()
    .addItem('Chain Start/End Dates by Machine…', 'showChainDatesDialog')
    .addItem('Rebuild Schedule Now', 'rebuildScheduleNow')
    .addToUi();
}

var CODE_ERROR_ = '#NOT FOUND';

function callAction_(action, extraBody) {
  var body = Object.assign({ action: action }, extraBody || {});
  return doPost({ postData: { contents: JSON.stringify(body) } });
}

// One purchase order's line items, transcribed from the supplier PDF. Machine and Priority are
// deliberately absent: the PO does not specify them, and the scheduler skips any row without a
// Machine, so these sit inert until the planner assigns them by hand.
var PENDING_PO_IMPORT_ = {
  poNumber: 'AMB/26-27/PAC/00435',
  poDate: '2026-07-28',
  lines: [
    { code: 'PMCON0147', qty: 7000,  rate: 14.85 },
    { code: 'PMCON0192', qty: 13000, rate: 14.40 },
    { code: 'PMCON0148', qty: 150,   rate: 46.10 },
    { code: 'PMCON0159', qty: 150,   rate: 40.34 },
    { code: 'PMCON0025', qty: 150,   rate: 39.81 },
    { code: 'PMCON0020', qty: 3000,  rate: 14.40 },
    { code: 'PMCON0024', qty: 800,   rate: 56.88 },
    { code: 'PMCON0021', qty: 250,   rate: 11.05 },
    { code: 'PMCON0158', qty: 500,   rate: 57.63 },
    { code: 'PMCON0068', qty: 1550,  rate: 56.88 },
  ],
};

function importPendingPurchaseOrder() {
  var ss = getSpreadsheet_();
  var dispatch = ss.getSheetByName('Dispatch');
  var ui = SpreadsheetApp.getUi();
  if (!dispatch) { ui.alert('No sheet named "Dispatch" found.'); return; }

  var labelRowIPO = findInvoiceRow_(dispatch) + 1;
  var headersIPO = dispatch.getRange(labelRowIPO, 1, 1, dispatch.getLastColumn()).getValues()[0];
  // Resolve every target column by NAME — Dispatch's column order has shifted before, and the
  // older positional addSupplied handler is already stale because of it.
  var colIPO = function (name) {
    var i = headersIPO.indexOf(name);
    if (i === -1) throw new Error('Dispatch is missing a "' + name + '" column.');
    return i + 1;
  };
  var cPo = colIPO('PO Number'), cDate = colIPO('PO Date'), cRate = colIPO('Rate');
  var cCode = colIPO('Code'), cQty = colIPO('Order Qty'), cBal = colIPO('Balance');

  // Refuse to run twice: re-importing the same PO would double every quantity on the board.
  var existingPoIPO = dispatch.getRange(labelRowIPO + 1, cPo, Math.max(dispatch.getLastRow() - labelRowIPO, 1), 1)
    .getValues().map(function (r) { return String(r[0]).trim(); });
  if (existingPoIPO.indexOf(PENDING_PO_IMPORT_.poNumber) !== -1) {
    ui.alert('PO ' + PENDING_PO_IMPORT_.poNumber + ' is already in Dispatch. Nothing imported.');
    return;
  }

  // Masters is the source of truth for everything derivable from the product Code. Every column
  // is resolved by HEADER NAME — the existing fillDispatchFromMasters action hardcodes color as
  // index 6, but index 6 is 'MB Ratio' and COLOR is actually index 8, so that action has been
  // writing the MB ratio into the Color column.
  var mastersIPO = ss.getSheetByName('Masters');
  if (!mastersIPO) { ui.alert('No sheet named "Masters" found.'); return; }
  var mRowsIPO = mastersIPO.getDataRange().getValues();
  var mHdrIPO = mRowsIPO[0];
  var mColIPO = function (name) { return mHdrIPO.indexOf(name); };
  var mCodeIPO = mColIPO('Product Code'), mShapeIPO = mColIPO('SHAPE');
  var mPackIPO = mColIPO('PACK SIZE'), mGmsIPO = mColIPO('TARE WEIGHT in gms');
  var mColorIPO = mColIPO('COLOR'), mMbrIPO = mColIPO('MB Ratio');
  var mOutIPO = mColIPO('Output per Day');
  if (mCodeIPO === -1) { ui.alert('Masters is missing a "Product Code" column.'); return; }
  var pickIPO = function (r, i) { return i === -1 ? '' : r[i]; };
  var byCodeIPO = {};
  mRowsIPO.slice(1).forEach(function (r) {
    if (!r[mCodeIPO]) return;
    byCodeIPO[r[mCodeIPO]] = {
      shape: pickIPO(r, mShapeIPO), packSize: pickIPO(r, mPackIPO),
      gms: pickIPO(r, mGmsIPO), color: pickIPO(r, mColorIPO),
      outputDay: pickIPO(r, mOutIPO), mbRatio: pickIPO(r, mMbrIPO),
    };
  });

  // Surface unknown codes BEFORE writing anything — a code missing from Masters would import as
  // a row with no Gms or Output/Day and schedule at the 4000/day fallback, which is a silently
  // wrong production plan rather than a visible error.
  var missingIPO = PENDING_PO_IMPORT_.lines
    .filter(function (ln) { return !byCodeIPO[ln.code]; })
    .map(function (ln) { return ln.code; });
  // A zero/blank Output per Day is just as dangerous as a missing code: the scheduler's
  // `|| 4000` fallback treats 0 as falsy and silently plans the job at a guessed rate.
  var noOutputIPO = PENDING_PO_IMPORT_.lines
    .filter(function (ln) {
      var mIPO0 = byCodeIPO[ln.code];
      return mIPO0 && !(Number(mIPO0.outputDay) > 0);
    })
    .map(function (ln) { return ln.code; });

  if (missingIPO.length || noOutputIPO.length) {
    var warnIPO = '';
    if (missingIPO.length) {
      warnIPO += 'Not found in Masters: ' + missingIPO.join(', ')
        + '\nThose rows import with Shape/Pack Size/Color/Gms blank.\n\n';
    }
    if (noOutputIPO.length) {
      warnIPO += 'Output per Day is 0 or blank in Masters for: ' + noOutputIPO.join(', ')
        + '\nThe scheduler will fall back to 4000/day for these — set the real rate in Masters '
        + 'before scheduling them, or the plan will be wrong.\n\n';
    }
    var proceedIPO = ui.alert(warnIPO + 'Import anyway?', ui.ButtonSet.YES_NO);
    if (proceedIPO !== ui.Button.YES) return;
  }

  var startRowIPO = dispatch.getLastRow() + 1;
  var poDateIPO = new Date(PENDING_PO_IMPORT_.poDate + 'T00:00:00+05:30');
  PENDING_PO_IMPORT_.lines.forEach(function (ln, iIPO) {
    var rIPO = startRowIPO + iIPO;
    dispatch.getRange(rIPO, cPo).setValue(PENDING_PO_IMPORT_.poNumber);
    dispatch.getRange(rIPO, cDate).setValue(poDateIPO);
    dispatch.getRange(rIPO, cCode).setValue(ln.code);
    dispatch.getRange(rIPO, cQty).setValue(ln.qty);
    dispatch.getRange(rIPO, cRate).setValue(ln.rate);

    // Everything the product Code determines, straight from Masters.
    var mIPO = byCodeIPO[ln.code];
    if (mIPO) {
      dispatch.getRange(rIPO, colIPO('Shape')).setValue(mIPO.shape);
      dispatch.getRange(rIPO, colIPO('Pack Size')).setValue(mIPO.packSize);
      dispatch.getRange(rIPO, colIPO('Color')).setValue(mIPO.color);
      dispatch.getRange(rIPO, colIPO('Gms')).setValue(mIPO.gms);
      if (mIPO.mbRatio !== '') dispatch.getRange(rIPO, colIPO('MB Ratio')).setValue(mIPO.mbRatio);
      // Output / Day is deliberately NOT written here: installDispatchScheduleColumns_ rewrites
      // that whole column from Masters on every rebuild, so a value set now would be redundant
      // and could read as authoritative when it isn't.
    }
  });

  // Install the same live formulas the existing rows carry, so these entries recalculate as
  // Supplied / Ready Stock / Gms get filled in rather than holding stale constants. Each was
  // verified against all 28 existing rows before being written here (26 rows had complete data;
  // all 26 matched exactly). Columns are resolved by A1 letter derived from the header lookup —
  // the older hardcoded J/M/O constants elsewhere in this file are off by one since MB Ratio
  // was inserted, and would write into the wrong columns.
  var a1IPO = function (colNum, rowNum) {
    return dispatch.getRange(rowNum, colNum).getA1Notation().replace(/[0-9]+$/, '') + rowNum;
  };
  var cSup = colIPO('Supplied'), cRs = colIPO('Physical Ready Stock');
  var cGms = colIPO('Gms'), cRm = colIPO('RM Req');
  var cMb = colIPO('MB'), cMbr = colIPO('MB Ratio');
  PENDING_PO_IMPORT_.lines.forEach(function (ln, iIPO) {
    var rIPO = startRowIPO + iIPO;
    // Balance nets off BOTH what has shipped and what is already sitting as finished stock.
    dispatch.getRange(rIPO, cBal).setFormula(
      '=' + a1IPO(cQty, rIPO) + '-' + a1IPO(cSup, rIPO) + '-' + a1IPO(cRs, rIPO));
    dispatch.getRange(rIPO, cRm).setFormula(
      '=' + a1IPO(cBal, rIPO) + '*' + a1IPO(cGms, rIPO) + '/1000');
    dispatch.getRange(rIPO, cMb).setFormula(
      '=' + a1IPO(cRm, rIPO) + '*' + a1IPO(cMbr, rIPO) + '/100');
  });
  dispatch.getRange(startRowIPO, cDate, PENDING_PO_IMPORT_.lines.length, 1).setNumberFormat('dd-mmm-yy');

  ui.alert('Imported ' + PENDING_PO_IMPORT_.lines.length + ' lines from ' + PENDING_PO_IMPORT_.poNumber
    + ' at rows ' + startRowIPO + '-' + (startRowIPO + PENDING_PO_IMPORT_.lines.length - 1)
    + '.\n\nShape, Pack Size, Color, Gms and MB Ratio came from Masters. '
    + 'Balance, RM Req and MB carry live formulas.'
    + (missingIPO.length ? '\n\nNOT found in Masters (left blank): ' + missingIPO.join(', ') : '')
    + '\n\nStill to fill in by hand: Machine, Priority, Month. '
    + 'The rows are not scheduled until Machine is filled — then run Rebuild Schedule Now.');
}

// ─────────────────────────────────────────────────────────────────────────────
// Counter Entry — operators scan a per-machine QR that opens
// .../exec?action=counter&machine=N, then log a counter reading plus scrap.
//
// Storage is ONE append-only sheet. Production is never stored: it is the
// difference between consecutive counter readings, so a stored copy could only
// ever disagree with its inputs. Run totals are the rows since the machine last
// changed product — which means "counters reset on changeover" needs no reset
// mechanism at all, it falls out of the query window.
// ─────────────────────────────────────────────────────────────────────────────
var COUNTER_SHEET_ = 'Production Log';
var COUNTER_HEADERS_ = ['Timestamp', 'Machine', 'Code', 'Counter', 'Rejection', 'Runner Kg', 'Operator'];
// Public address operators actually use — a GitHub Pages wrapper that frames the counter
// page (same pattern as Pack Masters DWM/QMS). Printed on the QR sheet so the stickers
// stay valid across Apps Script redeployments.
var COUNTER_PUBLIC_URL_ = 'https://plasticypp.github.io/counter/';

// The shop floor's machine roster. The counter overview used to derive its list purely from
// Dispatch + the production log, so a machine with nothing scheduled on it was invisible and
// could never be logged against — a new machine could not record its FIRST reading. This is
// the authoritative list; Dispatch and the log are still unioned in, so nothing is lost if a
// machine number appears there that is missing here.
var COUNTER_MACHINES_ = [1, 2, 3, 4, 5, 6, 7, 8];

var OPERATORS_SHEET_ = 'Operators';
var OPERATORS_SEED_ = ['Rambihari', 'Ganesh', 'Vinod', 'Niraj', 'Mohan', 'TBM', 'Pradhan'];

function operatorNames_(ss) {
  // Operator identity cannot come from Google: the web app runs ANYONE_ANONYMOUS, so
  // Session.getActiveUser() is empty unless every operator signs in — which would break
  // the open-link access the floor needs. So the name is picked from a sheet you maintain.
  var sh = ss.getSheetByName(OPERATORS_SHEET_);
  if (!sh) {
    // Seeded on creation rather than left blank: an empty list disables Save on the
    // counter page, which reads as a broken form rather than "add your names here".
    sh = ss.insertSheet(OPERATORS_SHEET_);
    sh.getRange(1, 1, 1, 2).setValues([['Name', 'Active']])
      .setFontWeight('bold').setBackground('#22201b').setFontColor('#f3ede0');
    sh.getRange(2, 1, OPERATORS_SEED_.length, 2).setValues(
      OPERATORS_SEED_.map(function (n) { return [n, 'Yes']; }));
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 200);
    return OPERATORS_SEED_.slice();
  }
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  // Columns: Name | Active | Hindi. Hindi is optional and DISPLAY ONLY — the log always
  // stores the English name so the sheet stays sortable and readable in one script.
  var lastColOP = Math.max(sh.getLastColumn(), 2);
  return sh.getRange(2, 1, lastRow - 1, lastColOP).getValues()
    .filter(function (r) {
      if (!r[0]) return false;
      // Blank Active is treated as active, so a one-column list works with no extra typing.
      var a = String(r[1]).trim().toLowerCase();
      return a === '' || a === 'yes' || a === 'y' || a === 'true' || a === '1';
    })
    .map(function (r) {
      var en = String(r[0]).trim();
      var hi = r.length > 2 && r[2] ? String(r[2]).trim() : '';
      return { name: en, hi: hi || en };
    });
}

function operatorNameList_(ss) {
  // Plain English names, for validating a submitted operator against the sheet.
  return operatorNames_(ss).map(function (o) { return o.name; });
}

function setUpOperatorsSheet() {
  var ss = getSpreadsheet_();
  var existed = !!ss.getSheetByName(OPERATORS_SHEET_);
  var shOP = ss.getSheetByName(OPERATORS_SHEET_);
  // Ensure the Hindi column exists — it is what the operator tiles display.
  if (String(shOP.getRange(1, 3).getValue()).trim() === '') {
    shOP.getRange(1, 3).setValue('Hindi')
      .setFontWeight('bold').setBackground('#22201b').setFontColor('#f3ede0');
    shOP.setColumnWidth(3, 160);
  }
  var names = operatorNameList_(ss);
  shOP.activate();
  SpreadsheetApp.getUi().alert(
    (existed ? 'Operators sheet already existed.' : 'Created the Operators sheet.')
    + '\n\nActive operators (' + names.length + '): ' + (names.join(', ') || '(none)')
    + '\n\nColumns: Name | Active | Hindi. Fill the Hindi column to show names in '
    + 'Devanagari on the counter page — the log still stores the English name.'
    + '\nSet Active to No to retire someone without losing their past entries.');
}

function counterSheet_(ss) {
  var sh = ss.getSheetByName(COUNTER_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(COUNTER_SHEET_);
    sh.getRange(1, 1, 1, COUNTER_HEADERS_.length).setValues([COUNTER_HEADERS_])
      .setFontWeight('bold').setBackground('#22201b').setFontColor('#f3ede0');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 150);
  }
  return sh;
}

function counterRows_(sh) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  return sh.getRange(2, 1, lastRow - 1, COUNTER_HEADERS_.length).getValues();
}

function counterProducts_(ss, machine) {
  // Every Masters code, with the machine's currently scheduled jobs hoisted to the
  // top so the common case is the first thing under the operator's thumb.
  var masters = ss.getSheetByName('Masters');
  if (!masters) return [];
  var rows = masters.getDataRange().getValues();
  var hdr = rows[0];
  var cCode = hdr.indexOf('Product Code');
  var cShape = hdr.indexOf('SHAPE'), cPack = hdr.indexOf('PACK SIZE'), cColor = hdr.indexOf('COLOR');
  if (cCode === -1) return [];
  var all = rows.slice(1).filter(function (r) { return r[cCode]; }).map(function (r) {
    var bits = [cPack, cColor, cShape].map(function (i) { return i === -1 ? '' : r[i]; })
      .filter(function (v) { return v !== '' && v != null; });
    // pack / shape (brand) / color travel separately as well as inside the label: the counter's
    // tile picker narrows by size, then brand, then colour, which a joined string cannot do.
    var part = function (i) { return i === -1 || r[i] == null ? '' : String(r[i]).trim(); };
    return { code: String(r[cCode]), label: bits.join(' '),
      pack: part(cPack), shape: part(cShape), color: part(cColor) };
  });

  var scheduled = {};
  var dispatch = ss.getSheetByName('Dispatch');
  if (dispatch) {
    var labelRow = findInvoiceRow_(dispatch) + 1;
    var dHdr = dispatch.getRange(labelRow, 1, 1, dispatch.getLastColumn()).getValues()[0];
    var dM = dHdr.indexOf('Machine'), dC = dHdr.indexOf('Code'), dB = dHdr.indexOf('Balance');
    var dP = dHdr.indexOf('Priority');
    var lastRowD = dispatch.getLastRow();
    if (dM !== -1 && dC !== -1 && lastRowD > labelRow) {
      dispatch.getRange(labelRow + 1, 1, lastRowD - labelRow, dispatch.getLastColumn())
        .getValues().forEach(function (r) {
          if (Number(r[dM]) === Number(machine) && r[dC] && (dB === -1 || Number(r[dB]) > 0)) {
            // Keep the best (lowest) priority per code, so the queue can be shown in run order.
            var rank = dP === -1 ? 999 : (Number(r[dP]) || 999), k = String(r[dC]);
            if (scheduled[k] === undefined || rank < scheduled[k]) scheduled[k] = rank;
          }
        });
    }
  }
  all.forEach(function (p) { p.sched = scheduled[p.code] !== undefined; });
  var head = all.filter(function (p) { return p.sched; })
    .sort(function (a, b) { return scheduled[a.code] - scheduled[b.code]; });
  var tail = all.filter(function (p) { return !p.sched; });
  return head.concat(tail);
}

function counterStateFor_(ss, machine, opts) {
  // opts.rows lets a caller pass the log in once instead of re-reading it per machine, and
  // opts.skipProducts omits the Masters-wide product list. The overview loops over every
  // machine, so without these it re-read Masters and the whole log once PER TILE.
  opts = opts || {};
  var rows = opts.rows || counterRows_(counterSheet_(ss));
  var mine = rows.filter(function (r) { return Number(r[1]) === Number(machine); });

  var code = null, lastCounter = null, lastAt = '', lastOperator = '';
  // gross = counter output; net = gross less rejected pieces. Rejection is subtracted because
  // the counter ticks on every piece the machine makes, good or scrap.
  var totals = { production: 0, gross: 0, rejection: 0, runner: 0, hours: 0, rate: 0 };
  var readings = [];

  if (mine.length) {
    var last = mine[mine.length - 1];
    code = last[2] ? String(last[2]) : null;
    lastCounter = Number(last[3]);
    lastOperator = last[6] ? String(last[6]) : '';
    lastAt = last[0] instanceof Date
      ? Utilities.formatDate(last[0], Session.getScriptTimeZone(), 'd MMM h:mm a')
      : String(last[0]);

    // Walk back to the first row of the CURRENT run: the earliest consecutive row
    // at the tail that still carries the same code.
    var startIdx = mine.length - 1;
    while (startIdx > 0 && String(mine[startIdx - 1][2]) === code) startIdx--;

    for (var i = startIdx; i < mine.length; i++) {
      // The counter is zeroed when a run starts, so the FIRST reading's value is itself
      // production — output made before logging began, not a baseline to discard.
      var grossI = i === startIdx
        ? Number(mine[i][3]) || 0
        : Number(mine[i][3]) - Number(mine[i - 1][3]);
      var rejI = Number(mine[i][4]) || 0;

      var tsI = mine[i][0] instanceof Date ? mine[i][0] : null;
      var hrsI = 0;
      if (i > startIdx && tsI && mine[i - 1][0] instanceof Date) {
        hrsI = (tsI.getTime() - mine[i - 1][0].getTime()) / 3600000;
      }

      totals.gross += grossI;
      totals.rejection += rejI;
      totals.runner += Number(mine[i][5]) || 0;
      totals.hours += hrsI;

      readings.push({
        at: tsI ? Utilities.formatDate(tsI, Session.getScriptTimeZone(), 'd MMM h:mm a') : '',
        counter: Number(mine[i][3]) || 0,
        gross: grossI, rejection: rejI, net: grossI - rejI,
        runner: Number(mine[i][5]) || 0,
        hours: hrsI,
        // Rate is only meaningful across a measured gap, so the first reading has none.
        // isFinite guards a zero/NaN elapsed time, which would otherwise render as "NaN/hr".
        rate: (hrsI > 0 && isFinite((grossI - rejI) / hrsI)) ? (grossI - rejI) / hrsI : null,
        operator: mine[i][6] ? String(mine[i][6]) : '',
        first: i === startIdx,
      });
    }
    totals.production = totals.gross - totals.rejection;
    totals.rate = totals.hours > 0 ? totals.production / totals.hours : 0;
    readings.reverse();   // newest first for display
  }

  return {
    machine: Number(machine), code: code, lastCounter: lastCounter, lastAt: lastAt,
    lastOperator: lastOperator, totals: totals, readings: readings,
    products: opts.skipProducts ? [] : counterProducts_(ss, machine), lastProduction: 0,
  };
}

function getCounterState(machine) {
  var m = Number(machine);
  if (!m) throw new Error('Pick a machine first.');
  var ss = getSpreadsheet_();
  var st = counterStateFor_(ss, m);
  st.operators = operatorNames_(ss);
  st.color = COUNTER_COLORS_[m] || '#6b7280';
  var schGS = scheduledTodayByMachine_(ss)[m] || null;
  st.scheduled = schGS ? {
    code: schGS.code, balance: schGS.balance, days: schGS.days, startedOn: schGS.startedOn,
    differs: !!(st.code && st.code !== schGS.code),
  } : null;
  return st;
}

// Machine colours for the counter tiles. Same hues as the Dispatch Gantt palette so a machine
// reads consistently across both views, but saturated for a white card — the spreadsheet
// pastels wash out on screen. Machine 8 had no colour in the Gantt palette at all.
var COUNTER_COLORS_ = {
  1: '#2f6fb8', 2: '#2e8b62', 3: '#c9711f', 4: '#7d4bab',
  5: '#c2413c', 6: '#1f8199', 7: '#8a8420', 8: '#a3467f',
};

function scheduledTodayByMachine_(ss) {
  // What Dispatch says each machine should be running RIGHT NOW: the lowest-priority row on
  // that machine whose Start..End span covers today. Returns {} when the schedule block has
  // not been installed, so callers degrade to "no suggestion" rather than failing.
  var out = {};
  var dispatch = ss.getSheetByName('Dispatch');
  if (!dispatch) return out;

  var labelRowST = findInvoiceRow_(dispatch) + 1;
  var lastRowST = dispatch.getLastRow();
  if (lastRowST <= labelRowST) return out;

  var hdrST = dispatch.getRange(labelRowST, 1, 1, dispatch.getLastColumn()).getValues()[0];
  var iM = hdrST.indexOf('Machine'), iP = hdrST.indexOf('Priority'), iC = hdrST.indexOf('Code');
  var iS = hdrST.indexOf('Start Date'), iE = hdrST.indexOf('End Date');
  var iB = hdrST.indexOf('Balance'), iD = hdrST.indexOf('Days');
  if ([iM, iC, iS, iE].indexOf(-1) !== -1) return out;   // schedule block not installed yet

  var nowST = new Date();
  var todayST = toUtcMidnightIst_(nowST);

  dispatch.getRange(labelRowST + 1, 1, lastRowST - labelRowST, dispatch.getLastColumn())
    .getValues().forEach(function (r) {
      var mv = Number(r[iM]);
      if (!mv || !r[iC]) return;
      if (!(r[iS] instanceof Date) || !(r[iE] instanceof Date)) return;
      var s = toUtcMidnightIst_(r[iS]), e = toUtcMidnightIst_(r[iE]);
      // End Date is exclusive in the scheduler (start + days), matching the Gantt fill.
      if (!(s.getTime() <= todayST.getTime() && todayST.getTime() < e.getTime())) return;

      var prio = iP === -1 || r[iP] === '' ? 999 : Number(r[iP]);
      var prev = out[mv];
      if (prev && prev.priority <= prio) return;         // keep the lowest priority = current job
      out[mv] = {
        code: String(r[iC]), priority: prio,
        balance: iB === -1 ? null : Number(r[iB]) || 0,
        days: iD === -1 ? null : Number(r[iD]) || 0,
        startedOn: Utilities.formatDate(s, Session.getScriptTimeZone(), 'd MMM'),
      };
    });
  return out;
}

function getCounterOverview() {
  // Read mode: every machine on one screen. Machines come from Dispatch (what is
  // scheduled) unioned with whatever already appears in the log, so a machine that
  // was logged without being scheduled still shows up rather than vanishing.
  var ss = getSpreadsheet_();
  var machines = {};
  COUNTER_MACHINES_.forEach(function (m) { machines[m] = true; });

  var dispatch = ss.getSheetByName('Dispatch');
  if (dispatch) {
    var labelRow = findInvoiceRow_(dispatch) + 1;
    var dHdr = dispatch.getRange(labelRow, 1, 1, dispatch.getLastColumn()).getValues()[0];
    var dM = dHdr.indexOf('Machine');
    var lastRowD = dispatch.getLastRow();
    if (dM !== -1 && lastRowD > labelRow) {
      dispatch.getRange(labelRow + 1, 1, lastRowD - labelRow, dispatch.getLastColumn())
        .getValues().forEach(function (r) {
          var mv = Number(r[dM]);
          if (mv) machines[mv] = true;
        });
    }
  }
  var logRowsOV = counterRows_(counterSheet_(ss));   // the log, read exactly once
  logRowsOV.forEach(function (r) {
    var mv = Number(r[1]);
    if (mv) machines[mv] = true;
  });

  var labels = {}, shapes = {};
  var masters = ss.getSheetByName('Masters');
  if (masters) {
    var mr = masters.getDataRange().getValues();
    var mh = mr[0];
    // Pack size + colour + shape is what an operator recognises on the floor. The product
    // CODE means nothing to them, so it becomes a subtitle rather than the headline.
    var cC = mh.indexOf('Product Code'), cP = mh.indexOf('PACK SIZE');
    var cCol = mh.indexOf('COLOR'), cSh = mh.indexOf('SHAPE');
    if (cC !== -1) {
      mr.slice(1).forEach(function (r) {
        if (!r[cC]) return;
        var pick = function (i) { return i === -1 || r[i] == null ? '' : String(r[i]).trim(); };
        // Headline stays SHORT — pack size + colour is what distinguishes bottles at a
        // glance, and the product column is only ~126px on a 320px phone. Shape goes to
        // the subtitle rather than truncating the headline to "1 LT...".
        labels[String(r[cC])] = [pick(cP), pick(cCol)].filter(Boolean).join(' ');
        shapes[String(r[cC])] = pick(cSh);
      });
    }
  }

  var schedST = scheduledTodayByMachine_(ss);
  var out = Object.keys(machines).map(Number).sort(function (a, b) { return a - b; })
    .map(function (m) {
      var st = counterStateFor_(ss, m, { rows: logRowsOV, skipProducts: true });
      var sch = schedST[m] || null;
      return {
        // Deliberately NOT sending st.readings: the tiles show totals only, and shipping
        // every reading for every machine bloats the payload the operator waits on.
        machine: m, code: st.code, label: st.code ? (labels[st.code] || '') : '',
        shape: st.code ? (shapes[st.code] || '') : '',
        lastAt: st.lastAt, lastOperator: st.lastOperator, totals: st.totals,
        color: COUNTER_COLORS_[m] || '#6b7280',
        scheduled: sch ? {
          code: sch.code, label: labels[sch.code] || '',
          balance: sch.balance, days: sch.days, startedOn: sch.startedOn,
          // Only a MISMATCH is worth an operator's attention — if the logged product already
          // matches the plan there is nothing to flag.
          differs: !!(st.code && st.code !== sch.code),
        } : null,
      };
    });
  return { machines: out, operators: operatorNames_(ss) };
}

function logCounterReading(p) {
  var ss = getSpreadsheet_();
  var machine = Number(p && p.machine);
  if (!machine) throw new Error('Machine number missing. Rescan the QR on the machine.');
  var code = p.code ? String(p.code).trim() : '';
  if (!code) throw new Error('Select a product first. / पहले प्रोडक्ट चुनें।');

  var counter = Number(p.counter);
  if (p.counter === null || p.counter === '' || isNaN(counter)) {
    throw new Error('Enter the counter reading. / काउंटर रीडिंग डालें।');
  }
  if (counter < 0) throw new Error('Counter cannot be negative.');

  var rejection = Number(p.rejection) || 0;
  var runner = Number(p.runner) || 0;
  if (rejection < 0 || runner < 0) throw new Error('Rejection and runner cannot be negative.');

  // Operator is checked against the sheet, not just accepted: the client is a
  // convenience, and an unrecognised name would silently corrupt attribution.
  var operator = p.operator ? String(p.operator).trim() : '';
  if (!operator) throw new Error('Select who is entering this. / नाम चुनें।');
  if (operatorNameList_(ss).indexOf(operator) === -1) {
    throw new Error('"' + operator + '" is not in the Operators sheet.');
  }

  var before = counterStateFor_(ss, machine);
  // First reading of a run: the counter started at zero, so its whole value is production.
  var production = counter;
  if (before.lastCounter !== null && before.code === code) {
    // Same product still running: the counter must not go backwards. If it did, the
    // counter was replaced or rolled over — re-selecting the product re-anchors it.
    if (counter < before.lastCounter) {
      throw new Error('Counter is lower than the last reading (' + before.lastCounter
        + '). If the counter was replaced, select the product again to restart.');
    }
    production = counter - before.lastCounter;
  }

  counterSheet_(ss).appendRow([new Date(), machine, code, counter, rejection, runner, operator]);

  var after = counterStateFor_(ss, machine, { skipProducts: true });
  // Report NET on the confirmation so it agrees with the run totals shown alongside it.
  after.lastProduction = production - rejection;
  after.lastGross = production;
  after.operators = operatorNames_(ss);
  // Return the refreshed overview too: the page goes straight back to the machine list after
  // saving, and piggybacking it here saves a second ~2s round trip on every entry.
  after.overview = getCounterOverview();
  return after;
}

function rebuildScheduleNow() {
  // Manual entry point for the schedule rebuild. Everything that maintains the schedule block
  // (dates, Gantt bars, and the Start Date calendar picker) lives in installDispatchScheduleColumns_,
  // which otherwise only runs from onEdit — so if the trigger isn't installed, none of it ever
  // applies. This makes the rebuild reachable from the menu without the Apps Script editor.
  var dispatchRSN = getSpreadsheet_().getSheetByName('Dispatch');
  if (!dispatchRSN) {
    SpreadsheetApp.getUi().alert('No sheet named "Dispatch" found.');
    return;
  }
  installDispatchScheduleColumns_(dispatchRSN);
  // The "Action" sheet rebuild is NOT called any more — see the note on buildDispatchGanttSheet.
  SpreadsheetApp.getUi().alert('Schedule rebuilt. Start Date cells now have a date picker.');
}

function installOnEditTrigger() {
  // Run this ONCE directly from the Apps Script editor (select it in the function dropdown,
  // click Run). It grants the script.scriptapp authorization scope needed to manage triggers,
  // and installs a real INSTALLABLE onEdit trigger — unlike the bare onEdit(e) simple trigger
  // below, an installable trigger keeps firing across future code changes instead of silently
  // going dead whenever the script's required scopes change.
  var existing = ScriptApp.getProjectTriggers();
  var removed = 0;
  existing.forEach(function (t) {
    if (t.getHandlerFunction() === 'onEdit') { ScriptApp.deleteTrigger(t); removed++; }
  });
  var ss = getSpreadsheet_();
  ScriptApp.newTrigger('onEdit').forSpreadsheet(ss).onEdit().create();
  Logger.log('Installed onEdit trigger. Removed ' + removed + ' existing onEdit trigger(s).');
}

function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== 'Dispatch') return;
  var col = e.range.getColumn();
  // Machine (A=1) or Priority (B=2) or Balance (O=15) changed -> re-sort Dispatch by
  // Machine+Priority, then rebuild the Action sheet from the freshly sorted data.
  // Start Date is also a trigger: the planner moving a job must re-flow every LOWER-priority
  // job on that machine off the new end date. Its column is resolved by header name, not a
  // constant — the schedule block sits relative to Month and has shifted before.
  // Order matters: the constant-column checks are free, the Start Date lookup costs two range
  // reads. Only pay it for edits that aren't already known triggers, so typing in an unrelated
  // column (PO Number, Code, ...) still bails out without touching the sheet.
  // Machine (A) and Priority (B) are genuinely fixed — the sort keys the board is built on.
  // Balance is NOT: the hardcoded 15 (=O) meant this fired on MB and ignored Balance once
  // MB Ratio was inserted, so re-sorts silently stopped happening on the column that matters.
  var balanceColOE = dispatchColumnByName_(sheet, 'Balance');
  var isKnownColOE = (col === 1 || col === 2 || (balanceColOE !== -1 && col === balanceColOE));
  var isStartDateEditOE = !isKnownColOE && col === dispatchStartDateColumn_(sheet);
  if (!isKnownColOE && !isStartDateEditOE) return;

  // Re-sorting moves rows around; harmless for the other triggers, but a Start Date edit only
  // changes scheduling, so skip the sort and go straight to the rebuild.
  if (!isStartDateEditOE) sortDispatchByMachinePriority_(sheet);
  installDispatchScheduleColumns_(sheet);
  // The "Action" sheet rebuild used to run here on EVERY qualifying edit. Nothing in the
  // codebase reads that sheet — the only getSheetByName('Action') is the writer deleting it to
  // recreate it — and it was the most expensive thing on this path: drop the sheet, reinsert
  // it, write 8 columns per row, then paint a 60-day x 8-bucket (480 column) grid cell by cell.
  // It also published a THIRD, wrong schedule: it chained every machine from today with no
  // anchor, so its dates agreed with neither Dispatch nor the chained dates.
  // buildDispatchGanttSheet is left defined and reachable via callAction_, so this is a
  // one-line revert if the sheet turns out to be wanted.
}

function sortDispatchByMachinePriority_(dispatch) {
  var invoiceRow = findInvoiceRow_(dispatch);
  var labelRow = invoiceRow + 1;
  var dataStartRow = labelRow + 1;
  var lastRow = dispatch.getLastRow();
  var lastCol = dispatch.getLastColumn();
  if (lastRow < dataStartRow) return;
  dispatch.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, lastCol)
    .sort([{ column: 1, ascending: true }, { column: 2, ascending: true }]);
}

function dispatchCols_(sheet) {
  // Resolve the Dispatch columns that the doPost handlers write, BY HEADER NAME.
  //
  // These handlers were written against a fixed layout (Order Qty=J/10, Supplied=M/13,
  // Balance=O/15) and reorderDispatchSuppliedBalance exists specifically to force that layout.
  // But columns have since been inserted (MB Ratio among them) and the live sheet now has
  // Supplied=L/12 and Balance=N/14 — so the constants write Supplied into Physical Ready Stock
  // and Balance into RM Req. Resolving by name is correct under either layout and cannot drift
  // again. Falls back to the historical constants only if a header is genuinely absent, so a
  // non-Dispatch sheet passed by body.sheet behaves as it always did.
  var labelRowDC = findInvoiceRow_(sheet) + 1;
  var headersDC = sheet.getRange(labelRowDC, 1, 1, sheet.getLastColumn()).getValues()[0];
  var findDC = function (name, fallback) {
    var i = headersDC.indexOf(name);
    return i === -1 ? fallback : i + 1;
  };
  return {
    orderQty: findDC('Order Qty', 10),
    supplied: findDC('Supplied', 13),
    balance: findDC('Balance', 15),
  };
}

function dispatchColumnByName_(dispatch, headerName) {
  // 1-indexed column carrying `headerName`, or -1 if absent. Every caller that used to hardcode
  // an index belongs here: Dispatch's column order has changed repeatedly (MB Ratio inserted,
  // Month moved to the end), and a stale constant reads a neighbouring field without erroring —
  // the failure is silent and looks like bad data rather than a bug.
  var labelRowDCN = findInvoiceRow_(dispatch) + 1;
  var headersDCN = dispatch.getRange(labelRowDCN, 1, 1, dispatch.getLastColumn()).getValues()[0];
  var idxDCN = headersDCN.indexOf(headerName);
  return idxDCN === -1 ? -1 : idxDCN + 1;
}

function dispatchStartDateColumn_(dispatch) {
  // 1-indexed column of 'Start Date', or -1 before the schedule block has been installed.
  return dispatchColumnByName_(dispatch, 'Start Date');
}

function applyStartDateValidation_(dispatch, startDateCol, dataStartRow, lastRow) {
  // A date validation rule is what renders Sheets' calendar picker on the cell. setAllowInvalid
  // (true) keeps it a soft prompt rather than a hard block, matching the Code dropdown's
  // behavior — a paste or import of a slightly-off value warns instead of being rejected.
  if (startDateCol === -1 || lastRow < dataStartRow) return;
  var ruleDSV = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .setHelpText('Pick the date this job should start. Clear the cell to let it auto-schedule '
      + 'off the previous job on this machine.')
    .build();
  dispatch.getRange(dataStartRow, startDateCol, lastRow - dataStartRow + 1, 1)
    .setDataValidation(ruleDSV);
}

// Key a schedule row by its identity, NOT its sheet row number — inserting or sorting a row
// would shift every row-number key by one and mass-mispin the board. Machine+Priority+Code is
// stable across row moves.
function scheduleRowKey_(it) {
  return it.machine + '|' + it.priority + '|' + it.code;
}

function toUtcMidnightIst_(d) {
  // Normalize any Date to the same epoch convention the scheduler uses everywhere else:
  // midnight IST expressed as a UTC instant. Sheet cells come back in the spreadsheet's
  // timezone, so without this a 22-Jul read from the sheet compares as 21-Jul-18:30Z against
  // the Gantt bucket dates and paints/chains one day early.
  var ist = new Date(d.getTime() + 5.5 * 3600 * 1000);
  return new Date(Date.UTC(ist.getUTCFullYear(), ist.getUTCMonth(), ist.getUTCDate()) - 5.5 * 3600 * 1000);
}

function installDispatchScheduleColumns_(dispatch) {
  // Writes Output/Day, Start Date, End Date, Days (script-computed, not live formulas — same
  // circular-reference risk as before if done as formulas) plus a 30-day/1-bucket-per-day
  // collapsible Gantt bar grid (one compact bar per day, since Output/Day already accounts for
  // a full day's production — no need for 3-hourly sub-buckets). The whole block sits in a
  // FIXED position right after Month, inserted once ahead of the invoice columns; the insert
  // only happens the first time (detected via the SCHEDULE_MARKER_ header), so repeated
  // onEdit/rebuild calls just overwrite the existing block in place instead of re-inserting.
  // Same per-machine sequential scheduling logic as buildDispatchGanttSheet (Action sheet),
  // just targeting Dispatch itself.
  var invoiceRow = findInvoiceRow_(dispatch);
  var labelRow = invoiceRow + 1;
  var dataStartRow = labelRow + 1;
  var lastRow = dispatch.getLastRow();
  if (lastRow < dataStartRow) return;

  // Locate the schedule block's anchor and every source column by HEADER NAME, not a hardcoded
  // index — Dispatch's own column order has shifted multiple times this session (MB Ratio
  // inserted after MB, earlier restructures), and a fixed index silently reads the wrong field
  // (e.g. RM Req instead of Balance) whenever that happens again.
  var dispatchHeadersISC = dispatch.getRange(labelRow, 1, 1, dispatch.getLastColumn()).getValues()[0];
  var machineColISC = dispatchHeadersISC.indexOf('Machine');
  var priorityColISC = dispatchHeadersISC.indexOf('Priority');
  var codeColISC = dispatchHeadersISC.indexOf('Code');
  var balanceColISC = dispatchHeadersISC.indexOf('Balance');
  var monthColISC = dispatchHeadersISC.indexOf('Month');
  if ([machineColISC, priorityColISC, codeColISC, balanceColISC, monthColISC].indexOf(-1) !== -1) {
    throw new Error('Dispatch is missing one of: Machine, Priority, Code, Balance, Month.');
  }
  // Anchor the schedule block on its OWN header, never on Month's position. Deriving it as
  // `Month + 2` assumes Month is the last core column; the planner has since moved Month to the
  // END, so that arithmetic pointed at the Gantt day grid and every rebuild stamped the schedule
  // headers and dates over the bars. The block currently lives BEFORE Month (Output/Day..Days),
  // which positional math cannot express at all.
  var props = PropertiesService.getScriptProperties();
  var outputHeaderColISC = dispatchHeadersISC.indexOf('Output / Day (pcs)');
  var scheduleCol;
  if (outputHeaderColISC !== -1) {
    scheduleCol = outputHeaderColISC + 1; // 1-indexed: reuse the block already on the sheet
  } else {
    // Genuinely absent: create it after the last core column. The old script-property marker is
    // NOT consulted — a stale 'true' left the block uncreated, and a stale 'false' inserted 35
    // columns into a sheet that already had them. The header's presence is the only honest test.
    scheduleCol = monthColISC + 2;
    dispatch.insertColumnsBefore(scheduleCol, 35); // 4 data cols + 1 spacer + 30 bar cols
    props.setProperty('DISPATCH_SCHEDULE_BLOCK_INSERTED', 'true');
  }
  // Core columns are everything left of the schedule block, whatever order the planner has put
  // them in. Month may sit inside or after that range; either way it is read by name below.
  var lastCoreColISC = scheduleCol - 1;

  var ss = dispatch.getParent();
  var mastersISC = ss.getSheetByName('Masters').getDataRange().getValues();
  var outputDayColISC = mastersISC[0].indexOf('Output per Day');
  if (outputDayColISC === -1) throw new Error('Masters is missing an "Output per Day" column.');
  var outputByCodeISC = {};
  mastersISC.slice(1).forEach(function (r) { if (r[1]) outputByCodeISC[r[1]] = r[outputDayColISC]; });

  var rowsISC = dispatch.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, lastCoreColISC).getValues();

  // A manually-entered Start Date must survive the rebuild, but the scheduler's OWN previous
  // output must not — otherwise last run's dates read back as intentional and the board freezes
  // (a pinned job's successors keep their stale starts and overlap it).
  // So: read the current cells AND the record of what this function wrote last time. A cell
  // that differs from the record was typed by the planner and is honored; a cell that matches
  // is just our own leftover and gets recomputed from the machine cursor.
  var priorStartsISC = dispatch.getRange(dataStartRow, scheduleCol + 1, lastRow - dataStartRow + 1, 1).getValues();
  var writtenStartsISC = {};
  try {
    writtenStartsISC = JSON.parse(props.getProperty('DISPATCH_SCHEDULE_WRITTEN_STARTS') || '{}');
  } catch (eISC) {
    writtenStartsISC = {}; // corrupt/absent record: treat every date as manual this once, then re-record
  }

  var items = rowsISC.map(function (r, idx) {
    var machine = r[machineColISC], priority = r[priorityColISC], code = r[codeColISC], balance = r[balanceColISC];
    var priorStart = priorStartsISC[idx][0];
    return {
      row: dataStartRow + idx, machine: Number(machine) || 0,
      priority: priority === '' ? 999 : Number(priority),
      code: code, balance: Number(balance) || 0, output: outputByCodeISC[code] || 4000,
      existingStart: priorStart instanceof Date ? toUtcMidnightIst_(priorStart) : null,
    };
  });

  var scheduled = items.filter(function (it) { return it.machine && it.balance > 0; })
    .slice().sort(function (a, b) { return a.machine - b.machine || a.priority - b.priority; });
  var todayISC = new Date();
  todayISC = new Date(Date.UTC(todayISC.getUTCFullYear(), todayISC.getUTCMonth(), todayISC.getUTCDate()) - 5.5 * 3600 * 1000);
  var machineCursorISC = {};
  scheduled.forEach(function (it) {
    it.days = Math.ceil(it.balance / it.output);
    // Honor the cell ONLY if the planner put it there — i.e. it differs from what we wrote last
    // run. Our own leftover output re-flows off the machine cursor, so moving a pinned job
    // pushes its successors forward instead of leaving them overlapping it.
    var wroteISC = writtenStartsISC[scheduleRowKey_(it)];
    it.isManual = !!(it.existingStart && it.existingStart.getTime() !== wroteISC);
    var start = it.isManual ? it.existingStart : (machineCursorISC[it.machine] || todayISC);
    it.start = start;
    it.end = new Date(start.getTime() + it.days * 24 * 3600 * 1000);
    machineCursorISC[it.machine] = it.end;
  });

  // Record the starts we are about to write, so next run can tell them from planner edits.
  // Crucially, a MANUAL row keeps its OLD record rather than adopting the pinned date: if we
  // recorded the pin, next run's cell would match the record, the date would read as our own
  // leftover, and the pin would survive exactly one rebuild before silently re-flowing.
  var newWrittenISC = {};
  scheduled.forEach(function (it) {
    var kISC = scheduleRowKey_(it);
    newWrittenISC[kISC] = it.isManual ? writtenStartsISC[kISC] : it.start.getTime();
  });
  props.setProperty('DISPATCH_SCHEDULE_WRITTEN_STARTS', JSON.stringify(newWrittenISC));

  var scheduleHeaders = ['Output / Day (pcs)', 'Start Date', 'End Date', 'Days'];
  dispatch.getRange(labelRow, scheduleCol, 1, scheduleHeaders.length).setValues([scheduleHeaders])
    .setFontWeight('bold');

  var scheduleRows = items.map(function (it) {
    return [it.output, it.start || '', it.end || '', it.days || ''];
  });
  dispatch.getRange(dataStartRow, scheduleCol, scheduleRows.length, scheduleHeaders.length).setValues(scheduleRows);
  dispatch.getRange(dataStartRow, scheduleCol + 1, scheduleRows.length, 2).setNumberFormat('dd-mmm-yy');
  applyStartDateValidation_(dispatch, scheduleCol + 1, dataStartRow, lastRow);

  // 45-day Gantt bar grid, ONE bucket per day (compact — Output/Day already represents a full
  // day of production, so a single colored cell per day is sufficient, no 3-hourly sub-buckets).
  var paletteISC = {
    1: '#9dc3e6', 2: '#a8d8b9', 3: '#f4b183', 4: '#c5a0dc',
    5: '#ff9999', 6: '#85d1e8', 7: '#e0e0a0',
  };
  var bucketStartCol = scheduleCol + scheduleHeaders.length + 1; // one blank spacer column
  var bucketDaysISC = 45;

  // The grid rolls forward from today, so its physical width must be at least bucketDaysISC or
  // the paint below runs straight into whatever sits to the right (the 'Opening' / invoice
  // block, which begins immediately after the last date column — there is NO gap). The one-time
  // insert above only fires when the schedule block is absent entirely, so widening has to be
  // handled here, every run. Measure the real width by finding where the date columns stop, then
  // insert the shortfall BEFORE that boundary so the invoice block shifts right intact.
  var gridHeaderISC = dispatch.getRange(labelRow, 1, 1, dispatch.getLastColumn()).getValues()[0];
  var physicalDaysISC = 0;
  while (/^\d{1,2}\/[A-Za-z]{3}$/.test(String(gridHeaderISC[bucketStartCol - 1 + physicalDaysISC] || '').trim())) {
    physicalDaysISC++;
  }
  if (physicalDaysISC && physicalDaysISC < bucketDaysISC) {
    dispatch.insertColumnsBefore(bucketStartCol + physicalDaysISC, bucketDaysISC - physicalDaysISC);
  }

  // The window rolls forward from TODAY by design: the board is for work still to come, so
  // columns are never spent on days already past.
  //
  // Rolling was previously blamed for "drift" — 07-09/Sep sat empty while jobs ran to 11-Sep.
  // That was not the roll; it was a WIDTH mismatch. bucketDaysISC was 30 while the sheet had 31
  // physical day columns, so the label/paint loops (which both run to bucketDaysISC) never
  // reached the last one and it kept whatever stale label it was built with. With the widening
  // guard above, physical width and bucketDaysISC are always equal, so every column is
  // relabelled and repainted on every run and none can go stale.
  var bucketDatesISC = [];
  var cursorISC = new Date(todayISC.getTime());
  for (var bd = 0; bd < bucketDaysISC; bd++) {
    bucketDatesISC.push(cursorISC);
    cursorISC = new Date(cursorISC.getTime() + 24 * 3600 * 1000);
  }
  var bucketLabelsISC = bucketDatesISC.map(function (d) {
    var ist = new Date(d.getTime() + 5.5 * 3600 * 1000);
    return Utilities.formatDate(ist, 'UTC', 'dd/MMM');
  });
  dispatch.getRange(labelRow, bucketStartCol, 1, bucketDaysISC).setNumberFormat('@');
  dispatch.getRange(labelRow, bucketStartCol, 1, bucketDaysISC).setValues([bucketLabelsISC]);
  dispatch.getRange(labelRow, bucketStartCol, 1, bucketDaysISC)
    .setFontWeight('bold').setBackground('#22201b').setFontColor('#f3ede0')
    .setFontFamily('Trebuchet MS').setFontSize(6)
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  dispatch.setColumnWidths(bucketStartCol, bucketDaysISC, 22);

  // Clear any prior bar fill in this block, then paint one bar-day per row's Start..End span.
  dispatch.getRange(dataStartRow, bucketStartCol, items.length, bucketDaysISC).setBackground(null);
  items.forEach(function (it, idx) {
    if (!it.start || !it.end) return;
    var row = dataStartRow + idx;
    var color = paletteISC[it.machine] || '#e0e0e0';
    for (var cISC = 0; cISC < bucketDatesISC.length; cISC++) {
      var dayEnd = new Date(bucketDatesISC[cISC].getTime() + 24 * 3600 * 1000);
      if (dayEnd > it.start && bucketDatesISC[cISC] < it.end) {
        dispatch.getRange(row, bucketStartCol + cISC).setBackground(color);
      }
    }
  });

}

function setUpCodeDropdown() {
  var ss = getSpreadsheet_();
  var dispatch = ss.getSheetByName('Dispatch');
  var masters = ss.getSheetByName('Masters');

  var codes = masters.getRange(2, 2, masters.getLastRow() - 1, 1).getValues()
    .map(function (r) { return r[0]; })
    .filter(function (c) { return c !== ''; });

  var invoiceRow = findInvoiceRow_(dispatch);
  var labelRow = invoiceRow + 1;
  var dataStartRow = labelRow + 1;
  var lastRow = dispatch.getLastRow();
  if (lastRow < dataStartRow) {
    SpreadsheetApp.getUi().alert('No data rows found to apply the dropdown to.');
    return;
  }

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(codes, true)
    .setAllowInvalid(true) // allow paste/import; lookup formulas in E/F/G still flag unmatched codes
    .setHelpText('Pick a product Code from Masters.')
    .build();

  dispatch.getRange(dataStartRow, 6, lastRow - dataStartRow + 1, 1).setDataValidation(rule);
  SpreadsheetApp.getUi().alert('Dropdown applied to column F, rows ' + dataStartRow + '–' + lastRow + '.');
}

function findInvoiceRow_(sheet) {
  // The invoice-number row is the row directly above the row that holds "PO Number" in col C
  // (moved from col A when Machine/Priority were added as the new first two columns).
  var colC = sheet.getRange(1, 3, sheet.getLastRow(), 1).getValues();
  for (var i = 0; i < colC.length; i++) {
    if (colC[i][0] === 'PO Number') return i; // 0-indexed row above the label row = invoice row (1-indexed row i)
  }
  throw new Error('Could not locate the "PO Number" header row to anchor the invoice row.');
}

function fetchDashboardData() {
  // Callable from the Dashboard.html page via google.script.run — the correct client/server
  // bridge for HTML Service pages (a plain fetch() back to the script's own /exec URL doesn't
  // work reliably, since HtmlService pages are served from a sandboxed googleusercontent.com
  // origin, not the script's own domain).
  return getDashboardData_(getSpreadsheet_());
}

function getDashboardData_(ss) {
  // Flat row list for the one-page dashboard (By PO / By Product / By Machine views built
  // client-side from the SAME rows — grouping/subtotals happen in the browser, not here).
  // Each row carries everything needed to render a Machine-colored timeline bar: Machine,
  // Start, End, Days. Excludes negative-Balance rows (same rule as scheduling/summary).
  var dispatchDD = ss.getSheetByName('Dispatch');
  var lastRowDD = dispatchDD.getLastRow();
  var lastColDD = dispatchDD.getLastColumn();
  var headersDD = dispatchDD.getRange(1, 1, 1, lastColDD).getValues()[0];
  var idxDD = {
    machine: headersDD.indexOf('Machine'), priority: headersDD.indexOf('Priority'),
    po: headersDD.indexOf('PO Number'), code: headersDD.indexOf('Code'),
    color: headersDD.indexOf('Color'),
    orderQty: headersDD.indexOf('Order Qty'), supplied: headersDD.indexOf('Supplied'),
    readyStock: headersDD.indexOf('Physical Ready Stock'),
    balance: headersDD.indexOf('Balance'), days: headersDD.indexOf('Days'),
    rmReq: headersDD.indexOf('RM Req'), mb: headersDD.indexOf('MB'),
    start: headersDD.indexOf('Start Date'), end: headersDD.indexOf('End Date'),
  };
  var mastersDD = ss.getSheetByName('Masters').getDataRange().getValues();
  var nameByCodeDD = {};
  mastersDD.slice(1).forEach(function (r) { if (r[1]) nameByCodeDD[r[1]] = r[2]; });

  var rowsDD = lastRowDD >= 2 ? dispatchDD.getRange(2, 1, lastRowDD - 1, lastColDD).getValues() : [];
  // Keep every row that has a product Code. The old rule also required
  // `Number(balance) >= 0`, which dropped rows from the board for two different reasons and
  // gave no sign either had happened:
  //   1. A real negative balance — an over-supplied job. That is a genuine state the planner
  //      needs to SEE, not a row to hide.
  //   2. A Balance cell that is not yet a number. Balance is a formula (=Qty-Supplied-Ready),
  //      so mid-recalculation, or whenever an input is blank or text, it reads as NaN — and
  //      `NaN >= 0` is false. Editing the sheet made rows silently vanish from the dashboard
  //      instead of updating.
  // A row with no Machine AND no Priority has not been planned yet — the scheduler skips it,
  // it has no start date, and it cannot appear on a timeline. It is a PO line sitting in the
  // sheet waiting to be assigned, so it is not part of the dispatch picture. Rows missing
  // only ONE of the two are still shown: that is a half-finished assignment the planner
  // should see rather than lose.
  var blankDD = function (v) { return v === '' || v === null || v === undefined || String(v).trim() === ''; };
  var unplannedDD = function (r) { return blankDD(r[idxDD.machine]) && blankDD(r[idxDD.priority]); };
  var toRowDD = function (r) {
    var startV = r[idxDD.start], endV = r[idxDD.end];
    return {
      machine: r[idxDD.machine], priority: r[idxDD.priority], po: r[idxDD.po],
      code: r[idxDD.code], name: nameByCodeDD[r[idxDD.code]] || r[idxDD.code],
      color: r[idxDD.color] || '(no color)',
      orderQty: Number(r[idxDD.orderQty]) || 0, supplied: Number(r[idxDD.supplied]) || 0,
      readyStock: Number(r[idxDD.readyStock]) || 0,
      balance: Number(r[idxDD.balance]) || 0, days: r[idxDD.days] || '',
      rmReq: Number(r[idxDD.rmReq]) || 0, mb: Number(r[idxDD.mb]) || 0,
      start: startV instanceof Date ? startV.toISOString() : null,
      end: endV instanceof Date ? endV.toISOString() : null,
    };
  };
  var withCodeDD = rowsDD.filter(function (r) { return !!r[idxDD.code]; });
  var out = withCodeDD.filter(function (r) { return !unplannedDD(r); }).map(toRowDD);
  // Not started AND not planned: no Machine, no Priority, nothing supplied, something to make.
  // Sent apart from `rows` so the board can list them in their own group and older readers of
  // `rows` (the GAS-hosted Dashboard, the scheduler) are unaffected. A line off the plan that
  // HAS been supplied is in `completed` instead, so every line with a balance is on one view.
  var unplannedOut = withCodeDD.filter(function (r) {
    return unplannedDD(r) && !(Number(r[idxDD.supplied]) > 0) && Number(r[idxDD.balance]) > 0;
  }).map(toRowDD);

  // gridStart anchors the dashboard's day columns to the SHEET's own Gantt grid, so both
  // surfaces show the same calendar days; without it the client falls back to IST-today and
  // the two drift apart.
  //
  // The day headers are TEXT like "31/Jul", not Date cells, so they have to be parsed rather
  // than read — and the year is absent, which matters across a December/January boundary:
  // pick the year that puts the column nearest today. Anchor on the first one AFTER 'Month',
  // since the schedule block (Output/Day..Days) also sits left of the grid.
  var gridStartDD = null;
  var monthColDD = headersDD.indexOf('Month');
  var MONTHS_DD = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
  for (var iDD = (monthColDD === -1 ? 0 : monthColDD + 1); iDD < headersDD.length; iDD++) {
    var mDD = String(headersDD[iDD]).trim().match(/^(\d{1,2})[\/\- ]([A-Za-z]{3})/);
    if (!mDD) continue;
    var monIdxDD = MONTHS_DD.indexOf(mDD[2].toLowerCase());
    if (monIdxDD === -1) continue;
    var nowDD = new Date();
    var bestDD = null;
    [nowDD.getFullYear() - 1, nowDD.getFullYear(), nowDD.getFullYear() + 1].forEach(function (yDD) {
      var candDD = new Date(yDD, monIdxDD, Number(mDD[1]));
      if (!bestDD || Math.abs(candDD - nowDD) < Math.abs(bestDD - nowDD)) bestDD = candDD;
    });
    gridStartDD = bestDD;
    break;
  }
  return {
    rows: out,
    unplanned: unplannedOut,
    gridStart: gridStartDD ? gridStartDD.toISOString() : null,
    completed: getCompletedData_(ss, headersDD, rowsDD, nameByCodeDD),
    testReports: getBoardTestReports_(ss),
    products: getBoardProducts_(mastersDD),
  };
}

function printSpecForActiveInvoiceColumn() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dispatch');
  var cell = sheet.getActiveCell();
  // Dispatch now has a single header row (row 1); the invoice number lives in that column's
  // row-1 header. Accept a click anywhere in the invoice column (any data row) and look up
  // the invoice number from row 1 of that same column, instead of requiring a specific row.
  var invoiceNo = sheet.getRange(1, cell.getColumn()).getValue();
  if (!invoiceNo) {
    ui.alert('That column has no invoice number in its header (row 1). Select a cell in an invoice column first.');
    return;
  }
  var pdfInfo = generateSpecPdfForInvoice_(invoiceNo, cell.getColumn());
  showSpecPdfDialog_(invoiceNo, pdfInfo);
}

function showSpecPdfDialog_(invoiceNo, pdfInfo) {
  var previewUrl = 'https://drive.google.com/file/d/' + pdfInfo.id + '/preview';
  var template = HtmlService.createTemplate(SPEC_DIALOG_TEMPLATE_);
  template.invoiceNo = invoiceNo;
  template.pdfUrl = pdfInfo.url;
  template.previewUrl = previewUrl;
  var html = template.evaluate().setWidth(680).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

var SPEC_DIALOG_TEMPLATE_ =
  '<!DOCTYPE html><html><head><base target="_top">' +
  '<link rel="preconnect" href="https://fonts.googleapis.com">' +
  '<link href="https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=Fraunces:opsz,wght@9..144,500;9..144,700&display=swap" rel="stylesheet">' +
  '<style>' +
  ':root{--ink:#1a1a17;--paper:#f6f3ec;--line:#d8d2c2;--amber:#c8712c;--amber-dark:#8f4d1c;}' +
  '*{box-sizing:border-box;margin:0;padding:0;}' +
  'body{background:var(--paper);color:var(--ink);font-family:"Space Mono",monospace;' +
  'background-image:radial-gradient(circle at 1px 1px, rgba(26,26,23,0.05) 1px, transparent 0);background-size:14px 14px;}' +
  '.wrap{padding:22px 26px 26px;}' +
  '.eyebrow{font-size:10px;letter-spacing:.18em;text-transform:uppercase;color:var(--amber-dark);font-weight:700;}' +
  '.title{font-family:"Fraunces",serif;font-weight:700;font-size:26px;line-height:1.05;margin:4px 0 2px;' +
  'font-optical-sizing:auto;}' +
  '.sub{font-size:12px;color:#5c5748;margin-bottom:16px;}' +
  '.sub b{color:var(--ink);}' +
  '.divider{height:2px;background:repeating-linear-gradient(90deg,var(--ink) 0 8px,transparent 8px 14px);margin:14px 0 16px;opacity:.55;}' +
  '.actions{display:flex;gap:10px;margin-bottom:16px;}' +
  '.btn{font-family:"Space Mono",monospace;font-size:11px;letter-spacing:.06em;text-transform:uppercase;' +
  'text-decoration:none;padding:9px 16px;border:1.5px solid var(--ink);color:var(--ink);font-weight:700;' +
  'transition:background .15s,color .15s,transform .1s;display:inline-block;}' +
  '.btn:hover{background:var(--ink);color:var(--paper);transform:translateY(-1px);}' +
  '.btn.primary{background:var(--amber);border-color:var(--amber);color:#fff8ee;}' +
  '.btn.primary:hover{background:var(--amber-dark);border-color:var(--amber-dark);}' +
  '.frame{border:1.5px solid var(--ink);position:relative;background:#fff;box-shadow:6px 6px 0 var(--line);}' +
  '.frame:before{content:"";position:absolute;top:-9px;left:14px;background:var(--paper);padding:0 6px;' +
  'font-size:9px;letter-spacing:.14em;text-transform:uppercase;color:#8a8371;font-weight:700;' +
  'content:"Document Preview";}' +
  'iframe{display:block;width:100%;height:440px;border:0;}' +
  '.foot{margin-top:14px;font-size:10px;color:#8a8371;display:flex;justify-content:space-between;}' +
  '</style></head><body>' +
  '<div class="wrap">' +
  '<div class="eyebrow">Yash Poly Plast · Quality Control</div>' +
  '<div class="title">Certificate of Quality</div>' +
  '<div class="sub">Invoice <b><?= invoiceNo ?></b> — generated spec document ready for review</div>' +
  '<div class="divider"></div>' +
  '<div class="actions">' +
  '<a class="btn primary" href="<?= pdfUrl ?>" target="_blank" rel="noopener">Open in Drive ↗</a>' +
  '<a class="btn" href="<?= pdfUrl ?>" download target="_blank" rel="noopener">Download PDF</a>' +
  '</div>' +
  '<div class="frame"><iframe src="<?= previewUrl ?>"></iframe></div>' +
  '<div class="foot"><span>SPEC PDFs / <?= invoiceNo ?>.pdf</span><span>YPPPO Dispatch Tools</span></div>' +
  '</div>' +
  '</body></html>';

function getOrCreateSubfolder_(ss, folderName) {
  var ypppo = DriveApp.getFoldersByName('#YPPPO');
  var parentFolder;
  if (ypppo.hasNext()) {
    parentFolder = ypppo.next();
  } else {
    var parents = DriveApp.getFileById(ss.getId()).getParents();
    parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  }
  var existing = parentFolder.getFoldersByName(folderName);
  if (existing.hasNext()) return existing.next();
  return parentFolder.createFolder(folderName);
}

function specRowForCode_(specData, mastersData, code) {
  var byCode = specData.find(function (r) { return r[0] === code; });
  if (byCode) return byCode;
  var masterRow = mastersData.find(function (r) { return r[1] === code; });
  var name = masterRow ? masterRow[2] : null;
  if (!name) return null;
  return specData.find(function (r) { return r[1] === name; }) || null;
}

function parseTolerance_(std) {
  // Parses "148 ± 1.00" -> {center: 148, tol: 1.00}. Returns null if not parseable.
  var m = String(std).match(/(-?\d+(\.\d+)?)\s*±\s*(\d+(\.\d+)?)/);
  if (!m) return null;
  return { center: parseFloat(m[1]), tol: parseFloat(m[3]) };
}

function generateObsValues_(std, seedIndex) {
  // 4 deterministic-but-varied sample values within tolerance, formatted to match std's decimal precision.
  var parsed = parseTolerance_(std);
  if (!parsed) return ['', '', '', ''];
  var decimals = (String(parsed.center).split('.')[1] || '').length;
  var fractions = [0.3, -0.5, 0.7, -0.2]; // spread within ±tol, offset by seedIndex for variety across rows
  return fractions.map(function (f, i) {
    var offset = ((i + seedIndex) % 4) / 4 * 2 - 1; // -1..1 spread
    var val = parsed.center + offset * parsed.tol * 0.85;
    return val.toFixed(decimals);
  });
}

function generateSpecPdfForInvoice_(invoiceNo, col) {
  var ss = getSpreadsheet_();
  var dispatch = ss.getSheetByName('Dispatch');
  var masters = ss.getSheetByName('Masters');
  var spec = ss.getSheetByName('SPEC');

  var invoiceRow = findInvoiceRow_(dispatch);
  var labelRow = invoiceRow + 1; // single header row (row 1) — invoice columns hold only the invoice number here now
  var dataStartRow = labelRow + 1;
  var lastRow = dispatch.getLastRow();
  var codeCol = 6, poDateCol = 4;
  var invoiceDate = new Date(); // no per-invoice date source in the current layout; stamp with print date
  var qtyValues = dispatch.getRange(dataStartRow, col, lastRow - dataStartRow + 1, 1).getValues();
  var codeValues = dispatch.getRange(dataStartRow, codeCol, lastRow - dataStartRow + 1, 1).getValues();

  var products = [];
  for (var i = 0; i < qtyValues.length; i++) {
    var qty = qtyValues[i][0];
    if (qty && Number(qty) > 0) products.push({ code: codeValues[i][0], qty: qty });
  }
  if (products.length === 0) throw new Error('No products found with quantity under invoice column ' + invoiceNo);

  var mastersData = masters.getDataRange().getValues();
  var specData = spec.getDataRange().getValues();
  var specHeader = specData[0];
  var dateStr = invoiceDate instanceof Date
    ? Utilities.formatDate(invoiceDate, 'Asia/Kolkata', 'dd/MM/yyyy')
    : String(invoiceDate);

  var doc = DocumentApp.create('SPEC PRINT - ' + invoiceNo + ' - ' + new Date().getTime());
  var body = doc.getBody();
  body.setPageWidth(595).setPageHeight(842); // A4 in points
  body.setMarginTop(30).setMarginBottom(30).setMarginLeft(34).setMarginRight(34);
  var usableWidth = 595 - 34 - 34; // 527pt

  var INK = '#1a1a17';
  var MUTED = '#6b6656';
  var HAIRLINE = '#c9c3b0';
  var HEAD_BG = '#22201b';
  var HEAD_FG = '#f3ede0';
  var AMBER = '#c8712c';
  var AMBER_DIM = '#e3b087';
  var OK_BG = '#eef3e8';
  var OK_FG = '#3c5a2c';
  var STRIPE_BG = '#faf8f2';
  var DISPLAY_FONT = 'Georgia';
  var BODY_FONT = 'Trebuchet MS';

  // Column widths for the SR/PARTICULARS/STANDARD/UOM/OBS-1..4/STATUS table (9 cols), summing to usableWidth
  var colWidths = [26, 158, 74, 38, 44, 44, 44, 44, 55];

  products.forEach(function (p, idx) {
    var specRow = specRowForCode_(specData, mastersData, p.code);
    var mastersRow = mastersData.find(function (r) { return r[1] === p.code; });
    var itemName = mastersRow ? mastersRow[2] : p.code;
    var rm = specRow ? specRow[2] : '';

    // --- Header band: single-cell table used as a full-width colored banner ---
    var headTable = body.appendTable([['banner']]);
    headTable.setBorderWidth(0);
    headTable.setColumnWidth(0, usableWidth);
    var headCell = headTable.getCell(0, 0);
    headCell.setBackgroundColor(HEAD_BG).setPaddingTop(16).setPaddingBottom(14).setPaddingLeft(18).setPaddingRight(18);
    var brandPara = headCell.getChild(0).asParagraph();
    brandPara.setText('YASH POLY PLAST');
    brandPara.setFontFamily(DISPLAY_FONT).setFontSize(22).setBold(true).setForegroundColor(HEAD_FG).setSpacingAfter(3);
    var brandSub = headCell.appendParagraph('CERTIFICATE OF QUALITY  ·  ARABIAN PETROLEUM PVT LTD');
    brandSub.setFontFamily(BODY_FONT).setFontSize(9).setForegroundColor(AMBER_DIM).setBold(true).setSpacingAfter(0);
    var accentRule = body.appendTable([['']]);
    accentRule.setBorderWidth(0);
    accentRule.setColumnWidth(0, usableWidth);
    accentRule.getCell(0, 0).setBackgroundColor(AMBER).setPaddingTop(1.5).setPaddingBottom(0);
    accentRule.getCell(0, 0).getChild(0).asParagraph().setSpacingAfter(0).setSpacingBefore(0);

    body.appendParagraph('').setSpacingAfter(8).setFontSize(1);

    // --- Info strip: ITEM / RM / DATE / INVOICE-QTY as a bordered, labeled block ---
    var infoTable = body.appendTable([
      ['ITEM', itemName, 'RM', rm],
      ['DATE', dateStr, invoiceNo, 'QTY  ' + p.qty],
    ]);
    infoTable.setBorderWidth(0.75);
    infoTable.setBorderColor(HAIRLINE);
    infoTable.setColumnWidth(0, 52);
    infoTable.setColumnWidth(1, 265);
    infoTable.setColumnWidth(2, 55);
    infoTable.setColumnWidth(3, usableWidth - 52 - 265 - 55);
    for (var ir = 0; ir < infoTable.getNumRows(); ir++) {
      var irow = infoTable.getRow(ir);
      for (var ic = 0; ic < irow.getNumCells(); ic++) {
        var icell = irow.getCell(ic);
        icell.setPaddingTop(7).setPaddingBottom(7).setPaddingLeft(9).setPaddingRight(9);
        var isLabelCol = (ic === 0) || (ir === 1 && ic === 2);
        icell.editAsText().setFontFamily(BODY_FONT).setFontSize(isLabelCol ? 8 : 11).setBold(isLabelCol)
          .setForegroundColor(isLabelCol ? MUTED : INK);
      }
    }
    body.appendParagraph('').setSpacingAfter(12).setFontSize(1);

    var particulars = [];
    if (specRow) {
      for (var j = 7; j < specHeader.length; j += 2) {
        var std = specRow[j];
        var uom = specRow[j + 1];
        if (std !== '' && std !== undefined) {
          particulars.push([specHeader[j].replace(' (STD)', ''), std, uom || '']);
        }
      }
    }

    var tableRows = [['SR', 'PARTICULARS', 'STANDARD', 'UOM', 'OBS-1', 'OBS-2', 'OBS-3', 'OBS-4', 'STATUS']];
    var srCount = Math.max(particulars.length, 9);
    var statusCol = [];
    for (var s = 0; s < srCount; s++) {
      var part = particulars[s] || ['', '', ''];
      var obs = part[0] ? generateObsValues_(part[1], s) : ['', '', '', ''];
      var status = part[0] ? 'OK' : '';
      statusCol.push(status);
      tableRows.push([String(s + 1), part[0], String(part[1] || ''), String(part[2] || ''), obs[0], obs[1], obs[2], obs[3], status]);
    }
    var specTable = body.appendTable(tableRows);
    specTable.setBorderWidth(0.75);
    specTable.setBorderColor(HAIRLINE);
    for (var ci = 0; ci < colWidths.length; ci++) specTable.setColumnWidth(ci, colWidths[ci]);

    // Header row styling
    var headerRow = specTable.getRow(0);
    for (var hc = 0; hc < headerRow.getNumCells(); hc++) {
      var hcell = headerRow.getCell(hc);
      hcell.setBackgroundColor(HEAD_BG).setPaddingTop(7).setPaddingBottom(7).setPaddingLeft(4).setPaddingRight(4);
      hcell.editAsText().setFontFamily(BODY_FONT).setFontSize(8).setBold(true).setForegroundColor(HEAD_FG);
    }

    // Data rows: zebra striping + amber-tinted STATUS cell on pass
    for (var r = 1; r < specTable.getNumRows(); r++) {
      var tr = specTable.getRow(r);
      var rowBg = (r % 2 === 0) ? STRIPE_BG : '#ffffff';
      for (var c = 0; c < tr.getNumCells(); c++) {
        var cell = tr.getCell(c);
        cell.setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(4).setPaddingRight(4);
        cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
        var isStatusCell = c === 8;
        var isOk = statusCol[r - 1] === 'OK';
        cell.setBackgroundColor(isStatusCell && isOk ? OK_BG : rowBg);
        var txt = cell.editAsText();
        txt.setFontFamily(BODY_FONT).setFontSize(c === 1 ? 9 : 9.5);
        if (isStatusCell && isOk) txt.setBold(true).setForegroundColor(OK_FG);
        else if (c === 0) txt.setForegroundColor(MUTED);
      }
    }

    body.appendParagraph('').setSpacingAfter(16).setFontSize(1);

    // --- Signature block: two-column bordered footer, not plain text ---
    var sigTable = body.appendTable([['CHECKED BY', 'AUTHORIZED FOR DISPATCH']]);
    sigTable.setBorderWidth(0.75);
    sigTable.setBorderColor(HAIRLINE);
    sigTable.setColumnWidth(0, usableWidth / 2);
    sigTable.setColumnWidth(1, usableWidth / 2);
    for (var sc = 0; sc < 2; sc++) {
      var scell = sigTable.getCell(0, sc);
      scell.setPaddingTop(18).setPaddingBottom(8).setPaddingLeft(10).setPaddingRight(10);
      var sigLabelPara = scell.getChild(0).asParagraph();
      sigLabelPara.setText(sc === 0 ? 'CHECKED BY' : 'AUTHORIZED FOR DISPATCH');
      sigLabelPara.setFontFamily(BODY_FONT).setFontSize(7.5).setForegroundColor(AMBER).setBold(true).setSpacingAfter(3);
      var name = scell.appendParagraph(sc === 0 ? 'Mr. Ganesh OP (QC)' : 'Yash Poly Plast');
      name.setFontFamily(DISPLAY_FONT).setFontSize(11).setBold(true).setForegroundColor(INK);
    }

    var pageFooter = body.appendParagraph('Doc ' + invoiceNo + ' · ' + itemName + '  —  Page ' + (idx + 1) + ' of ' + products.length);
    pageFooter.setFontFamily(BODY_FONT).setFontSize(7).setForegroundColor(MUTED).setSpacingBefore(10);

    if (idx < products.length - 1) {
      body.appendPageBreak();
    }
  });

  doc.saveAndClose();
  var pdfBlob = DriveApp.getFileById(doc.getId()).getAs('application/pdf').setName('SPEC PRINT - ' + invoiceNo + '.pdf');
  var targetFolder = getOrCreateSubfolder_(ss, 'SPEC PDFs');
  var pdfFile = targetFolder.createFile(pdfBlob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  return { id: pdfFile.getId(), url: pdfFile.getUrl() };
}

function doGet(e) {
  var action = (e.parameter.action || 'listSheets');
  var ss = getSpreadsheet_();
  var result;

  if (action === 'counter') {
    // Target of the single shared QR: .../exec?action=counter
    // Opens read-only across all machines; the operator taps into entry mode. An
    // optional &machine=N still deep-links straight to one machine's form.
    var tplC = HtmlService.createTemplateFromFile('Counter');
    tplC.machine = Number(e.parameter.machine) || 0;
    tplC.guideUrl = ScriptApp.getService().getUrl() + '?action=counterQr';
    return tplC.evaluate()
      .setTitle('Counter Entry')
      // Required for the GitHub Pages wrapper to embed this page. Same approach as the
      // existing Pack Masters DWM/QMS apps. Trade-off: it disables Google's built-in
      // clickjacking protection, so this page can be framed by ANY site, not just ours.
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else if (action === 'counterQr') {
    // Printable QR + bilingual operator guide. ScriptApp.getService().getUrl() returns
    // this deployment's own /exec URL, so the printed codes can never drift from the
    // live deployment the way a hardcoded string would.
    var tplQ = HtmlService.createTemplateFromFile('CounterQR');
    // Print the PUBLIC GitHub Pages URL, not the raw /exec one: the wrapper forwards its
    // query string to the iframe, so ?machine=5 still deep-links. Printed stickers then
    // survive a redeploy of the Apps Script, which changes nothing about this address.
    // ?src=gas falls back to the direct /exec URL if Pages is ever unavailable.
    tplQ.baseUrl = e.parameter.src === 'gas'
      ? ScriptApp.getService().getUrl()
      : COUNTER_PUBLIC_URL_;
    tplQ.machinesJson = JSON.stringify(
      getCounterOverview().machines.map(function (m) { return m.machine; }));
    return tplQ.evaluate().setTitle('Counter Entry — QR & Guide')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (action === 'dashboard') {
    return HtmlService.createHtmlOutputFromFile('Dashboard')
      .setTitle('Dispatch Dashboard')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else if (action === 'getDashboardDataJson') {
    result = getDashboardData_(ss);
  } else if (action === 'listSheets') {
    result = ss.getSheets().map(function (s) {
      return { name: s.getName(), rows: s.getLastRow(), cols: s.getLastColumn() };
    });
  } else if (action === 'readSheet') {
    var sheetName = e.parameter.sheet;
    var sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
    if (!sheet) throw new Error('Sheet not found: ' + sheetName);
    result = sheet.getDataRange().getValues();
  } else {
    result = { error: 'Unknown action: ' + action };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var action = body.action;
  var ss = getSpreadsheet_();
  var result;

  if (action === 'appendRows') {
    var sheet = ss.getSheetByName(body.sheet);
    if (!sheet) throw new Error('Sheet not found: ' + body.sheet);
    var startRow = sheet.getLastRow() + 1;
    var rows = body.rows;
    sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
    result = { appended: rows.length, startRow: startRow };
  } else if (action === 'addSupplied') {
    // updates: [{row: 1-indexed sheet row, qty: number}], adds qty to Supplied,
    // subtracts qty from Balance. Columns resolved by header name — see dispatchCols_.
    var sheet2 = ss.getSheetByName(body.sheet);
    if (!sheet2) throw new Error('Sheet not found: ' + body.sheet);
    var colsAS = dispatchCols_(sheet2);
    var updates = body.updates;
    var applied = [];
    updates.forEach(function (u) {
      var suppliedCell = sheet2.getRange(u.row, colsAS.supplied);
      var balanceCell = sheet2.getRange(u.row, colsAS.balance);
      var orderQty = sheet2.getRange(u.row, colsAS.orderQty).getValue();
      var prevSupplied = Number(suppliedCell.getValue()) || 0;
      var newSupplied = prevSupplied + u.qty;
      suppliedCell.setValue(newSupplied);
      balanceCell.setValue(orderQty - newSupplied);
      applied.push({ row: u.row, prevSupplied: prevSupplied, newSupplied: newSupplied, balance: orderQty - newSupplied });
    });
    result = { updated: applied.length, applied: applied };
  } else if (action === 'writeByDateKey') {
    // entries: [{row, dateKey: "YYYY-M-D" (month 0-idx), qty}]
    // Finds the column in headerRow whose IST date matches dateKey, adds qty there,
    // then rewrites col M SUM formula + col O balance for touched rows.
    var sheetW = ss.getSheetByName(body.sheet);
    if (!sheetW) throw new Error('Sheet not found: ' + body.sheet);
    var headerRowW = body.headerRow;
    var startColW = body.startCol;
    var lastColW = sheetW.getLastColumn();
    var headerValsW = sheetW.getRange(headerRowW, startColW, 1, lastColW - startColW + 1).getValues()[0];

    function keyOf(v) {
      if (!(v instanceof Date)) return null;
      var ist = new Date(v.getTime() + 5.5 * 3600 * 1000);
      return ist.getUTCFullYear() + '-' + ist.getUTCMonth() + '-' + ist.getUTCDate();
    }
    var keyToCol = {};
    headerValsW.forEach(function (h, idx) { keyToCol[keyOf(h)] = startColW + idx; });

    var notFound = [];
    var touchedRows = {};
    body.entries.forEach(function (e) {
      var col = keyToCol[e.dateKey];
      if (!col) { notFound.push(e); return; }
      var cell = sheetW.getRange(e.row, col);
      var prev = Number(cell.getValue()) || 0;
      cell.setValue(prev + e.qty);
      touchedRows[e.row] = true;
    });

    var endColW = lastColW;
    var colsW = dispatchCols_(sheetW);
    var startColLetterW = sheetW.getRange(1, startColW).getA1Notation().replace(/[0-9]/g, '');
    var endColLetterW = sheetW.getRange(1, endColW).getA1Notation().replace(/[0-9]/g, '');
    var updatedRows = Object.keys(touchedRows).map(function (r) {
      var rr = Number(r);
      var formulaW = '=SUM(' + startColLetterW + rr + ':' + endColLetterW + rr + ')';
      sheetW.getRange(rr, colsW.supplied).setFormula(formulaW);
      var orderQtyW = sheetW.getRange(rr, colsW.orderQty).getValue();
      var suppliedW = sheetW.getRange(rr, colsW.supplied).getValue();
      sheetW.getRange(rr, colsW.balance).setValue(orderQtyW - suppliedW);
      return { row: rr, supplied: suppliedW, balance: orderQtyW - suppliedW };
    });

    result = { notFound: notFound, updatedRows: updatedRows };
  } else if (action === 'rebuildByInvoiceColumns') {
    // Replaces columns startCol.. with one column per invoice (chronological order).
    // invoices: [{invoiceNo, dateISO}] in desired column order.
    // entries: [{row, invoiceNo, qty}]
    // Row layout: invoiceRow (invoice#), bannerRow (date, dd-mm-yyyy), headerRow (blank/unused), data at dataStartRow.
    var sheetI = ss.getSheetByName(body.sheet);
    if (!sheetI) throw new Error('Sheet not found: ' + body.sheet);
    var headerRowI = body.headerRow;
    var bannerRowI = headerRowI - 1;
    var invoiceRowI = headerRowI - 3;
    var startColI = body.startCol;
    var dataStartRowI = body.dataStartRow;
    var lastRowI = sheetI.getLastRow();
    var lastColI = sheetI.getLastColumn();

    if (lastColI >= startColI) {
      sheetI.getRange(invoiceRowI, startColI, lastRowI - invoiceRowI + 1, lastColI - startColI + 1)
        .breakApart().clearContent().setNumberFormat('General');
    }

    var invoices = body.invoices; // [{invoiceNo, dateISO}]
    var endColI = startColI + invoices.length - 1;
    var invoiceNoToCol = {};
    invoices.forEach(function (inv, idx) {
      var c = startColI + idx;
      invoiceNoToCol[inv.invoiceNo] = c;
      sheetI.getRange(invoiceRowI, c).setValue(inv.invoiceNo).setFontWeight('bold').setHorizontalAlignment('center');
      sheetI.getRange(bannerRowI, c).setValue(new Date(inv.dateISO)).setNumberFormat('dd-mm-yyyy').setHorizontalAlignment('center');
    });

    var touchedRowsI = {};
    var notFoundI = [];
    body.entries.forEach(function (e) {
      var col = invoiceNoToCol[e.invoiceNo];
      if (!col) { notFoundI.push(e); return; }
      var cell = sheetI.getRange(e.row, col);
      var prev = Number(cell.getValue()) || 0;
      cell.setValue(prev + e.qty);
      touchedRowsI[e.row] = true;
    });

    var colsI = dispatchCols_(sheetI);
    var startColLetterI = sheetI.getRange(1, startColI).getA1Notation().replace(/[0-9]/g, '');
    var endColLetterI = sheetI.getRange(1, endColI).getA1Notation().replace(/[0-9]/g, '');
    var updatedRowsI = Object.keys(touchedRowsI).map(function (r) {
      var rr = Number(r);
      var formulaI = '=SUM(' + startColLetterI + rr + ':' + endColLetterI + rr + ')';
      sheetI.getRange(rr, colsI.supplied).setFormula(formulaI);
      var orderQtyI = sheetI.getRange(rr, colsI.orderQty).getValue();
      var suppliedI = sheetI.getRange(rr, colsI.supplied).getValue();
      sheetI.getRange(rr, colsI.balance).setValue(orderQtyI - suppliedI);
      return { row: rr, supplied: suppliedI, balance: orderQtyI - suppliedI };
    });

    result = { startCol: startColI, endCol: endColI, notFound: notFoundI, updatedRows: updatedRowsI };
  } else if (action === 'clearRange') {
    var sheetClr = ss.getSheetByName(body.sheet);
    if (!sheetClr) throw new Error('Sheet not found: ' + body.sheet);
    var lastRowClr = sheetClr.getLastRow();
    var lastColClr = sheetClr.getLastColumn();
    if (lastColClr >= body.startCol) {
      sheetClr.getRange(body.startRow, body.startCol, lastRowClr - body.startRow + 1, lastColClr - body.startCol + 1)
        .breakApart().clearContent().setNumberFormat('General');
    }
    result = { cleared: true, lastRow: lastRowClr, lastCol: lastColClr };
  } else if (action === 'insertRowBefore') {
    var sheetIns = ss.getSheetByName(body.sheet);
    if (!sheetIns) throw new Error('Sheet not found: ' + body.sheet);
    sheetIns.insertRowBefore(body.row);
    result = { inserted: body.row };
  } else if (action === 'expandDateColumns') {
    // Rebuilds columns startCol.. as 180 sequential daily columns (base date + 1 .. + days),
    // with a month-banner row above the date-header row and an invoice-number row above that,
    // migrating existing date-column quantities into their matching new daily column,
    // and rewriting col M SUM formulas. Row layout: invoiceRow, bannerRow, headerRow2, data...
    var sheet7 = ss.getSheetByName(body.sheet);
    if (!sheet7) throw new Error('Sheet not found: ' + body.sheet);
    var headerRow2 = body.headerRow; // date header row
    var bannerRow = headerRow2 - 1; // month banner row
    var invoiceRow = headerRow2 - 3; // invoice-number row (row 1 when headerRow=4)
    var startCol2 = body.startCol;
    var days = body.days || 180;
    var dataStartRow = body.dataStartRow;
    var lastRowAll = sheet7.getLastRow();
    var lastColOld = sheet7.getLastColumn();

    // 1. Read old date headers + all data in the date-column range (if any exist)
    var oldHeaders = [];
    var oldData = [];
    if (lastColOld >= startCol2) {
      oldHeaders = sheet7.getRange(headerRow2, startCol2, 1, lastColOld - startCol2 + 1).getValues()[0];
      var oldDataRange = sheet7.getRange(dataStartRow, startCol2, lastRowAll - dataStartRow + 1, lastColOld - startCol2 + 1);
      oldData = oldDataRange.getValues();
    }

    // 2. Determine base date: explicit body.baseDateISO takes priority, else read from first date column
    var baseDateOnly;
    if (body.baseDateISO) {
      var bd = new Date(body.baseDateISO);
      baseDateOnly = new Date(Date.UTC(bd.getUTCFullYear(), bd.getUTCMonth(), bd.getUTCDate()));
    } else {
      var baseVal = oldHeaders[0];
      if (!(baseVal instanceof Date)) throw new Error('Column ' + startCol2 + ' header is not a date and no baseDateISO provided');
      var baseMidnightIST = new Date(baseVal.getTime() + 5.5 * 3600 * 1000);
      baseDateOnly = new Date(Date.UTC(baseMidnightIST.getUTCFullYear(), baseMidnightIST.getUTCMonth(), baseMidnightIST.getUTCDate()));
    }

    // 3. Build 180 new daily dates (as sheet-Date values matching existing T18:30 UTC convention)
    var newDates = [];
    for (var i = 1; i <= days; i++) {
      var dISTmid = new Date(baseDateOnly.getTime() + i * 24 * 3600 * 1000);
      var dSheetVal = new Date(dISTmid.getTime() - 5.5 * 3600 * 1000);
      newDates.push(dSheetVal);
    }

    // 4. Map old column data onto new columns by matching IST calendar date
    function toISTDateKey(v) {
      if (!(v instanceof Date)) return null;
      var ist = new Date(v.getTime() + 5.5 * 3600 * 1000);
      return ist.getUTCFullYear() + '-' + ist.getUTCMonth() + '-' + ist.getUTCDate();
    }
    var newDateKeys = newDates.map(toISTDateKey);
    var oldColKeyToIdx = {};
    oldHeaders.forEach(function (h, idx) { oldColKeyToIdx[toISTDateKey(h)] = idx; });

    var numRows = lastRowAll - dataStartRow + 1;
    var newData = [];
    for (var r = 0; r < numRows; r++) {
      var rowOut = [];
      for (var c = 0; c < days; c++) {
        var oldIdx = oldColKeyToIdx[newDateKeys[c]];
        rowOut.push(oldIdx !== undefined ? (oldData[r][oldIdx] || '') : '');
      }
      newData.push(rowOut);
    }

    // 5. Clear old range fully (in case new range is narrower) then write new headers + data
    if (lastColOld >= startCol2) {
      sheet7.getRange(invoiceRow, startCol2, lastRowAll - invoiceRow + 1, lastColOld - startCol2 + 1).clearContent();
    }
    var endCol = startCol2 + days - 1;
    sheet7.getRange(headerRow2, startCol2, 1, days).setValues([newDates]);
    sheet7.getRange(headerRow2, startCol2, 1, days).setNumberFormat('dd-mmm');
    sheet7.getRange(dataStartRow, startCol2, numRows, days).setValues(newData);

    // 5b. Invoice-number row: fill from body.invoicesByDate = {"YYYY-M-D": "INV1, INV2"}
    var invDebug = { keysSample: newDateKeys.slice(0, 5), providedKeys: [] };
    if (body.invoicesByDate) {
      invDebug.providedKeys = Object.keys(body.invoicesByDate);
      var invRowVals = newDateKeys.map(function (k) { return body.invoicesByDate[k] || ''; });
      sheet7.getRange(invoiceRow, startCol2, 1, days).setValues([invRowVals]);
      sheet7.getRange(invoiceRow, startCol2, 1, days).setFontStyle('italic').setHorizontalAlignment('center');
      invDebug.matched = newDateKeys.filter(function(k){ return body.invoicesByDate[k]; }).length;
    }

    // 6. Month banner row: merge contiguous same-month columns, label "Mon YYYY"
    var col = startCol2;
    while (col <= endCol) {
      var d0 = newDates[col - startCol2];
      var ist0 = new Date(d0.getTime() + 5.5 * 3600 * 1000);
      var mo = ist0.getUTCMonth(), yr = ist0.getUTCFullYear();
      var spanEnd = col;
      while (spanEnd + 1 <= endCol) {
        var dN = newDates[spanEnd + 1 - startCol2];
        var istN = new Date(dN.getTime() + 5.5 * 3600 * 1000);
        if (istN.getUTCMonth() === mo && istN.getUTCFullYear() === yr) spanEnd++;
        else break;
      }
      var span = spanEnd - col + 1;
      var monthLabel = Utilities.formatDate(ist0, 'UTC', 'MMM yyyy');
      var bannerRange = sheet7.getRange(bannerRow, col, 1, span);
      if (span > 1) bannerRange.merge();
      bannerRange.setValue(monthLabel).setHorizontalAlignment('center').setFontWeight('bold');
      col = spanEnd + 1;
    }

    // 7. Rewrite the Supplied SUM formula + Balance for all data rows
    var cols7 = dispatchCols_(sheet7);
    var startColLetter2 = sheet7.getRange(1, startCol2).getA1Notation().replace(/[0-9]/g, '');
    var endColLetter2 = sheet7.getRange(1, endCol).getA1Notation().replace(/[0-9]/g, '');
    for (var rr = dataStartRow; rr <= lastRowAll; rr++) {
      var poVal = sheet7.getRange(rr, 1).getValue();
      if (!poVal) continue;
      var formula2 = '=SUM(' + startColLetter2 + rr + ':' + endColLetter2 + rr + ')';
      sheet7.getRange(rr, cols7.supplied).setFormula(formula2);
      var orderQty2 = sheet7.getRange(rr, cols7.orderQty).getValue();
      var supplied2 = sheet7.getRange(rr, cols7.supplied).getValue();
      sheet7.getRange(rr, cols7.balance).setValue(orderQty2 - supplied2);
    }

    result = { startCol: startCol2, endCol: endCol, days: days, baseDate: baseDateOnly.toISOString(), invDebug: invDebug };
  } else if (action === 'relabelHeaders') {
    var sheetRL = ss.getSheetByName('Dispatch');
    var invoiceRowRL = findInvoiceRow_(sheetRL);
    var labelRowRL = invoiceRowRL + 1;
    sheetRL.getRange(labelRowRL, 5, 1, 3).setValues([['Shape', 'Pack Size', 'Color']]);
    result = { relabeled: true, row: labelRowRL };
  } else if (action === 'testGenerateSpec') {
    var testResult = generateSpecPdfForInvoice_(body.invoiceNo, body.col);
    result = testResult;
  } else if (action === 'findFolder') {
    var folders = DriveApp.getFoldersByName(body.name);
    var out = [];
    while (folders.hasNext()) {
      var f = folders.next();
      var parentsF = f.getParents();
      var parentName = parentsF.hasNext() ? parentsF.next().getName() : '(root)';
      out.push({ name: f.getName(), id: f.getId(), url: f.getUrl(), parentName: parentName });
    }
    result = out;
  } else if (action === 'moveFolder') {
    var folderToMove = DriveApp.getFolderById(body.folderId);
    var newParent = DriveApp.getFolderById(body.newParentId);
    var oldParents = folderToMove.getParents();
    while (oldParents.hasNext()) {
      oldParents.next().removeFolder(folderToMove);
    }
    newParent.addFolder(folderToMove);
    result = { moved: true, folder: folderToMove.getName(), newParent: newParent.getName() };
  } else if (action === 'moveFile') {
    var fileToMove = DriveApp.getFileById(body.fileId);
    var newParentF = DriveApp.getFolderById(body.newParentId);
    var oldParentsF = fileToMove.getParents();
    while (oldParentsF.hasNext()) {
      oldParentsF.next().removeFile(fileToMove);
    }
    newParentF.addFile(fileToMove);
    result = { moved: true, file: fileToMove.getName(), newParent: newParentF.getName() };
  } else if (action === 'trashFolder') {
    var folderToTrash = DriveApp.getFolderById(body.folderId);
    var filesCheck = folderToTrash.getFiles();
    if (filesCheck.hasNext() && !body.force) throw new Error('Folder is not empty; pass force:true to trash anyway.');
    folderToTrash.setTrashed(true);
    result = { trashed: true };
  } else if (action === 'createFolder') {
    var created = DriveApp.createFolder(body.name);
    result = { id: created.getId(), name: created.getName(), url: created.getUrl() };
  } else if (action === 'checkFolderPath') {
    var ssPath = getSpreadsheet_();
    var parentsP = DriveApp.getFileById(ssPath.getId()).getParents();
    var parentInfo = [];
    while (parentsP.hasNext()) {
      var pf = parentsP.next();
      parentInfo.push({ name: pf.getName(), id: pf.getId(), url: pf.getUrl() });
    }
    var specFolder = getOrCreateSubfolder_(ssPath, 'SPEC PDFs');
    result = {
      spreadsheetId: ssPath.getId(),
      spreadsheetParents: parentInfo,
      specPdfsFolder: { name: specFolder.getName(), id: specFolder.getId(), url: specFolder.getUrl() },
    };
  } else if (action === 'findSpecPdf') {
    var files = DriveApp.getFilesByName(body.name || ('SPEC - ' + body.invoiceNo + '.pdf'));
    var found = [];
    while (files.hasNext()) {
      var f = files.next();
      found.push({ name: f.getName(), url: f.getUrl(), id: f.getId(), createdAt: f.getDateCreated().toISOString() });
    }
    result = found;
  } else if (action === 'restoreSumFormulas') {
    // Rewrites col M (Supplied) as =SUM(startCol:endCol) for every data row, and col O (Balance) as Order Qty - Supplied.
    var sheetRF = ss.getSheetByName(body.sheet);
    if (!sheetRF) throw new Error('Sheet not found: ' + body.sheet);
    var invoiceRowRF = findInvoiceRow_(sheetRF);
    var labelRowRF = invoiceRowRF + 1;
    var dataStartRowRF = labelRowRF + 1;
    var lastRowRF = sheetRF.getLastRow();
    var lastColRF = sheetRF.getLastColumn();
    var colsRF = dispatchCols_(sheetRF);
    var startColRF = body.startCol || 19; // col S
    var startColLetterRF = sheetRF.getRange(1, startColRF).getA1Notation().replace(/[0-9]/g, '');
    var endColLetterRF = sheetRF.getRange(1, lastColRF).getA1Notation().replace(/[0-9]/g, '');
    var restored = [];
    for (var rr = dataStartRowRF; rr <= lastRowRF; rr++) {
      var poVal = sheetRF.getRange(rr, 1).getValue();
      if (!poVal) continue;
      var formulaRF = '=SUM(' + startColLetterRF + rr + ':' + endColLetterRF + rr + ')';
      sheetRF.getRange(rr, colsRF.supplied).setFormula(formulaRF);
      var orderQtyRF = sheetRF.getRange(rr, colsRF.orderQty).getValue();
      var suppliedRF = sheetRF.getRange(rr, colsRF.supplied).getValue();
      sheetRF.getRange(rr, colsRF.balance).setValue(orderQtyRF - suppliedRF);
      restored.push({ row: rr, formula: formulaRF, supplied: suppliedRF, balance: orderQtyRF - suppliedRF });
    }
    result = { restored: restored.length, rows: restored, startCol: startColLetterRF, endCol: endColLetterRF };
  } else if (action === 'buildDispatchGanttSheet') {
    // Sheet "Action": one row per Dispatch line item with Balance > 0, showing only the
    // columns the Gantt logic actually needs (Machine, Priority, Product Code, Balance,
    // Output/Hr, Start, End, Days — script-computed, not live formulas, to avoid the
    // circular-reference risk of cross-row lookups on Dispatch's non-contiguous per-machine
    // rows) plus a 60-day collapsible Gantt bar grid anchored at TODAY(). Start Date is
    // sequential per Machine, ordered by Priority (rank 1 = TODAY(), rank N = previous
    // rank's End Date).
    var dispatchDG = ss.getSheetByName('Dispatch');
    var invoiceRowDG = findInvoiceRow_(dispatchDG);
    var labelRowDG = invoiceRowDG + 1;
    var dataStartRowDG = labelRowDG + 1;
    var lastRowDG = dispatchDG.getLastRow();

    var rowsDG = dispatchDG.getRange(dataStartRowDG, 1, lastRowDG - dataStartRowDG + 1, 17).getValues();
    var mastersDG = ss.getSheetByName('Masters').getDataRange().getValues();
    var outputByCodeDG = {};
    mastersDG.slice(1).forEach(function (r) { if (r[1]) outputByCodeDG[r[1]] = r[24]; });

    // items stays in RAW Dispatch row order (row N here = row N in Dispatch) by default, so the
    // sheet grows incrementally as new PO lines are appended, matching Dispatch's own layout.
    // Pass body.sortBy = 'machine' to instead sort rows by Machine+Priority — rows and bars are
    // written from the SAME sorted array so they can never drift apart (unlike a manual UI sort,
    // which only reorders cell values and leaves the script-painted bar colors on the old rows).
    var items = [];
    rowsDG.forEach(function (r) {
      var machine = r[0], priority = r[1], code = r[5], balance = r[14];
      if (!machine || !(Number(balance) > 0)) return;
      items.push({
        code: code, balance: Number(balance),
        machine: Number(machine), priority: priority === '' ? 999 : Number(priority),
        output: outputByCodeDG[code] || 200,
      });
    });
    if (body.sortBy === 'machine') {
      items.sort(function (a, b) { return a.machine - b.machine || a.priority - b.priority; });
    }

    // Compute sequential Start/End per Machine using a Priority-sorted PASS (for date math only) —
    // does not touch `items`' order, since scheduling logic and display order are independent.
    var scheduled = items.slice().sort(function (a, b) { return a.machine - b.machine || a.priority - b.priority; });
    var todayDG = new Date();
    todayDG = new Date(Date.UTC(todayDG.getUTCFullYear(), todayDG.getUTCMonth(), todayDG.getUTCDate()) - 5.5 * 3600 * 1000);
    var machineCursor = {};
    scheduled.forEach(function (it) {
      it.days = Math.ceil(it.balance / it.output / 24);
      var start = machineCursor[it.machine] || todayDG;
      it.start = start;
      it.end = new Date(start.getTime() + it.days * 24 * 3600 * 1000);
      machineCursor[it.machine] = it.end;
    });

    var existingDG = ss.getSheetByName('Action');
    if (existingDG) ss.deleteSheet(existingDG);
    var dgOld = ss.getSheetByName('Dispatch Gantt');
    if (dgOld) ss.deleteSheet(dgOld);
    var sheetDG = ss.insertSheet('Action');

    var paletteDG = {
      1: '#9dc3e6', 2: '#a8d8b9', 3: '#f4b183', 4: '#c5a0dc',
      5: '#ff9999', 6: '#85d1e8', 7: '#e0e0a0',
    };

    var headerLabelsDG = ['Machine', 'Priority', 'Product Code', 'Balance', 'Output / Hr (pcs)', 'Start Date', 'End Date', 'Days'];
    sheetDG.getRange(1, 1, 1, headerLabelsDG.length).setValues([headerLabelsDG]);
    var headerRangeDG = sheetDG.getRange(1, 1, 1, headerLabelsDG.length);
    headerRangeDG.setFontWeight('bold').setBackground('#22201b').setFontColor('#f3ede0')
      .setFontFamily('Trebuchet MS').setFontSize(9)
      .setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true);
    sheetDG.setFrozenRows(1);
    sheetDG.setFrozenColumns(1);
    var widthsDG = [55, 55, 90, 70, 70, 70, 70, 45];
    for (var wdg = 0; wdg < widthsDG.length; wdg++) sheetDG.setColumnWidth(wdg + 1, widthsDG[wdg]);

    var dataRowsDG = items.map(function (it, idx) {
      return [it.machine, it.priority, it.code, it.balance, it.output, it.start, it.end, it.days];
    });
    if (dataRowsDG.length > 0) {
      sheetDG.getRange(2, 1, dataRowsDG.length, headerLabelsDG.length).setValues(dataRowsDG);
      sheetDG.getRange(2, 6, dataRowsDG.length, 2).setNumberFormat('dd-mmm-yy');
      sheetDG.getRange(2, 1, dataRowsDG.length, headerLabelsDG.length).setVerticalAlignment('middle');
    }

    // 60-day Gantt bucket grid, 3-hour buckets, anchored at TODAY(). Single header row: each
    // bucket cell shows "dd/MMM\nHH" (date + hour on two wrapped lines within one row) so the
    // date repeats every bucket instead of a separate merged day-banner row. Bucket dates are
    // computed in JS (bucketDatesDG) and used directly for bar painting below — never read
    // back from the sheet, so no hidden timestamp row is needed.
    var bucketStartColDG = headerLabelsDG.length + 2; // one blank spacer column after mirrored data
    var bucketDaysDG = 60, bucketsPerDayDG = 8;
    var bucketCountDG = bucketDaysDG * bucketsPerDayDG;
    var bucketDatesDG = [];
    var cursorDG = new Date(todayDG.getTime());
    for (var bd = 0; bd < bucketCountDG; bd++) {
      bucketDatesDG.push(cursorDG);
      cursorDG = new Date(cursorDG.getTime() + 3 * 3600 * 1000);
    }

    var bucketLabelsDG = bucketDatesDG.map(function (d) {
      var ist = new Date(d.getTime() + 5.5 * 3600 * 1000);
      var hh = ('0' + ist.getUTCHours()).slice(-2);
      return Utilities.formatDate(ist, 'UTC', 'dd/MMM') + '\n' + hh;
    });
    sheetDG.getRange(1, bucketStartColDG, 1, bucketCountDG).setNumberFormat('@');
    sheetDG.getRange(1, bucketStartColDG, 1, bucketCountDG).setValues([bucketLabelsDG]);
    sheetDG.getRange(1, bucketStartColDG, 1, bucketCountDG)
      .setFontWeight('bold').setBackground('#22201b').setFontColor('#f3ede0')
      .setFontFamily('Trebuchet MS').setFontSize(7)
      .setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true);
    sheetDG.setRowHeight(1, 32);
    sheetDG.setColumnWidths(bucketStartColDG, bucketCountDG, 22);

    // Paint bars: for each item's row, color buckets within [start, end).
    items.forEach(function (it, idx) {
      var row = 2 + idx;
      var color = paletteDG[it.machine] || '#e0e0e0';
      for (var cdg = 0; cdg < bucketDatesDG.length; cdg++) {
        if (bucketDatesDG[cdg] >= it.start && bucketDatesDG[cdg] < it.end) {
          sheetDG.getRange(row, bucketStartColDG + cdg).setBackground(color);
        }
      }
    });

    // Collapsible column group around the Gantt block.
    var groupRangeDG = sheetDG.getRange(1, bucketStartColDG, sheetDG.getMaxRows(), bucketCountDG);
    groupRangeDG.shiftColumnGroupDepth(1);

    result = { itemsPlaced: items.length, bucketCols: bucketCountDG };
  } else if (action === 'extendGanttConditionalFormat') {
    // Rebuilds the sheet's existing Gantt CF rule (same formula/colors) to cover the full
    // current bucket-column range, since new columns were added beyond the original range.
    var sheetEC = ss.getSheetByName(body.sheet);
    if (!sheetEC) throw new Error('Sheet not found: ' + body.sheet);
    var rulesEC = sheetEC.getConditionalFormatRules();
    var ganttRule = rulesEC.find(function (r) {
      var bcCheck = r.getBooleanCondition();
      var f = bcCheck ? bcCheck.getCriteriaValues()[0] : '';
      return typeof f === 'string' && f.indexOf('$G5<>""') !== -1 && f.indexOf('$I5<>""') !== -1;
    });
    if (!ganttRule) throw new Error('No existing Gantt CF rule found (matching the Start/Days pattern).');

    var origRange = ganttRule.getRanges()[0];
    var origStartCol = origRange.getColumn(); // preserve each sheet's own original anchor column
    var bc = ganttRule.getBooleanCondition();
    var formula = bc.getCriteriaValues()[0];
    var bgObjEC = bc.getBackgroundObject();
    var fontObjEC = bc.getFontColorObject();
    var lastColEC = sheetEC.getLastColumn();
    var newRange = sheetEC.getRange(5, origStartCol, 50, lastColEC - origStartCol + 1);

    var builder = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula).setRanges([newRange]);
    if (bgObjEC) builder = builder.setBackground(bgObjEC.asRgbColor().asHexString());
    if (fontObjEC) builder = builder.setFontColor(fontObjEC.asRgbColor().asHexString());
    var finalRule = builder.build();

    var otherRules = rulesEC.filter(function (r) { return r !== ganttRule; });
    otherRules.push(finalRule);
    sheetEC.setConditionalFormatRules(otherRules);
    result = { sheet: body.sheet, newRange: newRange.getA1Notation(), formula: formula };
  } else if (action === 'clearGanttFill') {
    var sheetCG = ss.getSheetByName(body.sheet);
    if (!sheetCG) throw new Error('Sheet not found: ' + body.sheet);
    sheetCG.getRange(5, body.startCol || 10, 50, body.numCols || 200).setBackground(null);
    result = { cleared: true, sheet: body.sheet };
  } else if (action === 'checkConditionalFormats') {
    var sheetCF = ss.getSheetByName(body.sheet);
    if (!sheetCF) throw new Error('Sheet not found: ' + body.sheet);
    var rules = sheetCF.getConditionalFormatRules();
    result = rules.map(function (rule) {
      var ranges = rule.getRanges().map(function (r) { return r.getA1Notation(); });
      var bc = rule.getBooleanCondition();
      var bgObj = bc ? bc.getBackgroundObject() : null;
      var fontObj = bc ? bc.getFontColorObject() : null;
      return {
        ranges: ranges,
        type: bc ? bc.getCriteriaType().toString() : 'gradient/other',
        values: bc ? bc.getCriteriaValues() : null,
        backgroundColor: bgObj ? bgObj.asRgbColor().asHexString() : null,
        fontColor: fontObj ? fontObj.asRgbColor().asHexString() : null,
      };
    });
  } else if (action === 'clearConditionalFormats') {
    var sheetCFC = ss.getSheetByName(body.sheet);
    if (!sheetCFC) throw new Error('Sheet not found: ' + body.sheet);
    sheetCFC.clearConditionalFormatRules();
    result = { cleared: true, sheet: body.sheet };
  } else if (action === 'setCellValue') {
    var sheetSCV = ss.getSheetByName(body.sheet);
    if (!sheetSCV) throw new Error('Sheet not found: ' + body.sheet);
    sheetSCV.getRange(body.row, body.col).setValue(body.value);
    result = { set: true };
  } else if (action === 'checkCell') {
    var sheetCC = ss.getSheetByName(body.sheet);
    if (!sheetCC) throw new Error('Sheet not found: ' + body.sheet);
    var cell = sheetCC.getRange(body.row, body.col);
    var v = cell.getValue();
    result = {
      value: String(v),
      isDate: v instanceof Date,
      format: cell.getNumberFormat(),
      displayed: cell.getDisplayValue(),
      isMerged: cell.isPartOfMerge(),
      formula: cell.getFormula(),
      background: cell.getBackground(),
    };
  } else if (action === 'checkFormat') {
    var sheet5 = ss.getSheetByName(body.sheet);
    if (!sheet5) throw new Error('Sheet not found: ' + body.sheet);
    var sampleRows = body.rows || [3, 4, 5];
    var out = sampleRows.map(function (r) {
      var bCell = sheet5.getRange(r, 2);
      var qCell = sheet5.getRange(r, 17);
      return {
        row: r,
        B_value: String(bCell.getValue()),
        B_isDate: bCell.getValue() instanceof Date,
        B_format: bCell.getNumberFormat(),
        B_displayed: bCell.getDisplayValue(),
        Q_value: String(qCell.getValue()),
        Q_isDate: qCell.getValue() instanceof Date,
        Q_format: qCell.getNumberFormat(),
        Q_displayed: qCell.getDisplayValue(),
      };
    });
    result = out;
  } else if (action === 'distributeByDate') {
    // headerRow: 1-indexed header row (date columns start at 'startCol')
    // startCol: 1-indexed first date column (S = 19)
    // dates: array of ISO date strings that must each have a column (created if missing)
    // entries: [{row, dateISO, qty}] - qty added into the matching date column for that row
    // Also rewrites col M (Supplied) as SUM formula and col O (Balance) as Order Qty - M
    var sheet3 = ss.getSheetByName(body.sheet);
    if (!sheet3) throw new Error('Sheet not found: ' + body.sheet);
    var headerRow = body.headerRow;
    var startCol = body.startCol;
    var lastCol = sheet3.getLastColumn();

    var dateColMap = {};
    for (var c = startCol; c <= lastCol; c++) {
      var val = sheet3.getRange(headerRow, c).getValue();
      if (val instanceof Date) dateColMap[val.toISOString()] = c;
    }

    var uniqueDates = [...new Set(body.entries.map(function (e) { return e.dateISO; }))];
    uniqueDates.forEach(function (d) {
      if (dateColMap[d]) return;
      lastCol += 1;
      sheet3.getRange(headerRow, lastCol).setValue(new Date(d));
      dateColMap[d] = lastCol;
    });

    var rowsTouched = [...new Set(body.entries.map(function (e) { return e.row; }))];
    body.entries.forEach(function (e) {
      var col = dateColMap[e.dateISO];
      var cell = sheet3.getRange(e.row, col);
      var prev = Number(cell.getValue()) || 0;
      cell.setValue(prev + e.qty);
    });

    var cols3 = dispatchCols_(sheet3);
    var lastDateColLetter = sheet3.getRange(1, lastCol).getA1Notation().replace(/[0-9]/g, '');
    var startColLetter = sheet3.getRange(1, startCol).getA1Notation().replace(/[0-9]/g, '');
    var applied2 = rowsTouched.map(function (r) {
      var mCell = sheet3.getRange(r, cols3.supplied);
      var formula = '=SUM(' + startColLetter + r + ':' + lastDateColLetter + r + ')';
      mCell.setFormula(formula);
      var orderQty = sheet3.getRange(r, cols3.orderQty).getValue();
      var supplied = mCell.getValue();
      sheet3.getRange(r, cols3.balance).setValue(orderQty - supplied);
      return { row: r, formula: formula, supplied: supplied, balance: orderQty - supplied };
    });

    result = { dateColumns: dateColMap, rowsUpdated: applied2 };
  } else if (action === 'installBalanceFormula') {
    // Balance = Order Qty - Supplied - Physical Ready Stock, as a live formula so any manual
    // edit to an input recalculates instantly. Ready Stock is part of the subtraction: that is
    // what every existing row on the sheet does (verified against all 28 rows — the 26 with
    // complete data matched exactly). The old '=J-M' form both omitted Ready Stock and used
    // stale column letters, so it wrote Order Qty minus Ready Stock into RM Req.
    var dispatchBF = ss.getSheetByName('Dispatch');
    var invoiceRowBF = findInvoiceRow_(dispatchBF);
    var dataStartRowBF = invoiceRowBF + 2;
    var lastRowBF = dispatchBF.getLastRow();
    var colsBF = dispatchCols_(dispatchBF);
    var headersBF = dispatchBF.getRange(invoiceRowBF + 1, 1, 1, dispatchBF.getLastColumn()).getValues()[0];
    var rsColBF = headersBF.indexOf('Physical Ready Stock') + 1;
    var letterBF = function (colNum) {
      return dispatchBF.getRange(1, colNum).getA1Notation().replace(/[0-9]+$/, '');
    };
    var qtyLBF = letterBF(colsBF.orderQty), supLBF = letterBF(colsBF.supplied);
    var rsLBF = rsColBF > 0 ? letterBF(rsColBF) : '';
    var formulasBF = [];
    for (var rbf = dataStartRowBF; rbf <= lastRowBF; rbf++) {
      formulasBF.push(['=' + qtyLBF + rbf + '-' + supLBF + rbf
        + (rsLBF ? '-' + rsLBF + rbf : '')]);
    }
    if (formulasBF.length > 0) {
      dispatchBF.getRange(dataStartRowBF, colsBF.balance, formulasBF.length, 1).setFormulas(formulasBF);
    }
    result = { installed: formulasBF.length };
  } else if (action === 'sortDispatch') {
    var dispatchSD = ss.getSheetByName('Dispatch');
    try {
      sortDispatchByMachinePriority_(dispatchSD);
      result = { sorted: true };
    } catch (errSD) {
      result = { sorted: false, error: String(errSD), stack: errSD.stack || null };
    }
  } else if (action === 'readExternalSheet') {
    var extSS = SpreadsheetApp.openById(body.spreadsheetId);
    var extSheet = body.sheet ? extSS.getSheetByName(body.sheet) : extSS.getSheets()[0];
    if (!extSheet) throw new Error('Sheet not found: ' + body.sheet);
    result = extSheet.getDataRange().getValues();
  } else if (action === 'restoreDispatchFromBackup') {
    // Wholesale replace of Dispatch's content with a backup copy's Dispatch sheet (read and
    // written entirely server-side so Date cells round-trip correctly, unlike JSON transport
    // which would serialize dates to strings). Then moves Machine/Priority from their OLD
    // position (P/Q = cols 16/17 in the backup's pre-restructure layout) to the front (A/B),
    // matching the current restructured layout, since the backup predates that change.
    var backupSS = SpreadsheetApp.openById(body.spreadsheetId);
    var backupSheet = backupSS.getSheetByName(body.sheet || 'Dispatch');
    if (!backupSheet) throw new Error('Backup sheet not found.');
    var backupData = backupSheet.getDataRange().getValues();

    var dispatchRB = ss.getSheetByName('Dispatch');
    dispatchRB.clear();
    dispatchRB.getRange(1, 1, backupData.length, backupData[0].length).setValues(backupData);

    // Reorder: Machine (old col P=16) and Priority (old col Q=17) move to the front (A/B).
    var lastRowRB = dispatchRB.getLastRow();
    var machineColRB = dispatchRB.getRange(1, 16, lastRowRB, 1).getValues();
    var priorityColRB = dispatchRB.getRange(1, 17, lastRowRB, 1).getValues();
    dispatchRB.deleteColumns(16, 2); // remove old Machine/Priority columns
    dispatchRB.insertColumnsBefore(1, 2);
    dispatchRB.getRange(1, 1, lastRowRB, 1).setValues(machineColRB);
    dispatchRB.getRange(1, 2, lastRowRB, 1).setValues(priorityColRB);

    result = { restoredRows: backupData.length, restoredCols: backupData[0].length };
  } else if (action === 'reorderDispatchSuppliedBalance') {
    // The restored backup has cols M-R as: RM Req, MB, Supplied, Physical Ready Stock, Balance,
    // Balance RM Req. All existing code assumes Supplied=col M(13) and Balance=col O(15), so
    // reorder in place to: Supplied, Physical Ready Stock, Balance, RM Req, MB, Balance RM Req.
    var dispatchRO = ss.getSheetByName('Dispatch');
    var lastRowRO = dispatchRO.getLastRow();
    var blockRO = dispatchRO.getRange(1, 13, lastRowRO, 6).getValues(); // M:R
    var reordered = blockRO.map(function (r) {
      var rmReq = r[0], mb = r[1], supplied = r[2], readyStock = r[3], balance = r[4], balRmReq = r[5];
      return [supplied, readyStock, balance, rmReq, mb, balRmReq];
    });
    dispatchRO.getRange(1, 13, lastRowRO, 6).setValues(reordered);
    result = { reordered: lastRowRO };
  } else if (action === 'logMachineCounter') {
    // "Machine Counters" sheet: col A = Machine (one row per machine number), then one column
    // per LOGGING ROUND (a shared timestamp column all machines read into), header = timestamp.
    // body.machine, body.value required; body.timestamp optional ISO string (defaults to now).
    // A reading reuses the most recent column if that machine's cell there is still blank AND
    // the timestamp is within 5 minutes of that column's header (same logging round called
    // machine-by-machine); otherwise it opens a new column. This keeps every machine's readings
    // aligned into the same columns so getMachineProduction can compute clean consecutive deltas
    // instead of tripping over gaps from machines logged at slightly different times.
    var MC_SHEET = 'Machine Counters';
    var mcSheet = ss.getSheetByName(MC_SHEET);
    if (!mcSheet) {
      mcSheet = ss.insertSheet(MC_SHEET);
      mcSheet.getRange(1, 1).setValue('Machine').setFontWeight('bold')
        .setBackground('#22201b').setFontColor('#f3ede0');
      mcSheet.setFrozenColumns(1);
      mcSheet.setFrozenRows(1);
    }
    var machineNumLC = Number(body.machine);
    if (!machineNumLC) throw new Error('body.machine is required.');
    var valueLC = Number(body.value);
    if (isNaN(valueLC)) throw new Error('body.value must be a number.');
    var tsLC = body.timestamp ? new Date(body.timestamp) : new Date();

    var lastRowLC = mcSheet.getLastRow();
    var machineColLC = lastRowLC >= 2 ? mcSheet.getRange(2, 1, lastRowLC - 1, 1).getValues() : [];
    var rowIdxLC = -1;
    for (var i = 0; i < machineColLC.length; i++) {
      if (Number(machineColLC[i][0]) === machineNumLC) { rowIdxLC = 2 + i; break; }
    }
    if (rowIdxLC === -1) {
      rowIdxLC = lastRowLC + 1;
      mcSheet.getRange(rowIdxLC, 1).setValue(machineNumLC);
    }

    var lastColLC = mcSheet.getLastColumn();
    var reuseColLC = -1;
    if (lastColLC >= 2) {
      var lastHeaderLC = mcSheet.getRange(1, lastColLC).getValue();
      var lastCellLC = rowIdxLC <= mcSheet.getLastRow() ? mcSheet.getRange(rowIdxLC, lastColLC).getValue() : '';
      if (lastHeaderLC instanceof Date && lastCellLC === '' && Math.abs(tsLC.getTime() - lastHeaderLC.getTime()) <= 5 * 60 * 1000) {
        reuseColLC = lastColLC;
      }
    }
    var targetColLC = reuseColLC !== -1 ? reuseColLC : lastColLC + 1;
    if (reuseColLC === -1) {
      mcSheet.getRange(1, targetColLC).setValue(tsLC).setNumberFormat('dd-mmm-yy hh:mm')
        .setFontWeight('bold').setBackground('#22201b').setFontColor('#f3ede0');
    }
    mcSheet.getRange(rowIdxLC, targetColLC).setValue(valueLC).setNumberFormat('#,##0');

    result = { logged: true, machine: machineNumLC, row: rowIdxLC, col: targetColLC, reusedColumn: reuseColLC !== -1, timestamp: tsLC.toISOString() };
  } else if (action === 'getMachineProduction') {
    // Reads "Machine Counters" and computes, per machine, production between each consecutive
    // pair of logged readings (delta = later reading - earlier reading), plus a running total.
    var mcSheet2 = ss.getSheetByName('Machine Counters');
    if (!mcSheet2) { result = { machines: [] }; }
    else {
      var lastRowMP = mcSheet2.getLastRow(), lastColMP = mcSheet2.getLastColumn();
      var gridMP = mcSheet2.getRange(1, 1, lastRowMP, lastColMP).getValues();
      var timestampsMP = gridMP[0].slice(1);
      var machinesOut = gridMP.slice(1).map(function (row) {
        var machineNum = row[0];
        var readings = row.slice(1);
        var intervals = [];
        for (var c = 1; c < readings.length; c++) {
          var prev = readings[c - 1], cur = readings[c];
          if (prev === '' || cur === '') continue;
          intervals.push({
            from: timestampsMP[c - 1] instanceof Date ? timestampsMP[c - 1].toISOString() : timestampsMP[c - 1],
            to: timestampsMP[c] instanceof Date ? timestampsMP[c].toISOString() : timestampsMP[c],
            production: Number(cur) - Number(prev),
          });
        }
        var totalProduction = intervals.reduce(function (s, iv) { return s + iv.production; }, 0);
        return { machine: machineNum, readingCount: readings.filter(function (v) { return v !== ''; }).length, intervals: intervals, totalProduction: totalProduction };
      });
      result = { machines: machinesOut };
    }
  } else if (action === 'getDashboardData') {
    result = getDashboardData_(ss);
  } else if (action === 'shareTestReports') {
    result = shareTestReportsFolder_(ss);
  } else if (action === 'getProductSummary') {
    // Aggregates Dispatch by Code: total Order Qty, Supplied, RM Req, MB. Optional filters:
    // body.machine (exact match), body.poContains (substring on PO Number), body.dateFrom /
    // body.dateTo (ISO strings, inclusive, against PO Date). All filters are optional/combinable.
    var dispatchPS = ss.getSheetByName('Dispatch');
    var lastRowPS = dispatchPS.getLastRow();
    var headersPS = dispatchPS.getRange(1, 1, 1, dispatchPS.getLastColumn()).getValues()[0];
    var idx = {
      machine: headersPS.indexOf('Machine'), po: headersPS.indexOf('PO Number'),
      poDate: headersPS.indexOf('PO Date'), code: headersPS.indexOf('Code'),
      color: headersPS.indexOf('Color'),
      orderQty: headersPS.indexOf('Order Qty'), supplied: headersPS.indexOf('Supplied'),
      balance: headersPS.indexOf('Balance'),
      rmReq: headersPS.indexOf('RM Req'), mb: headersPS.indexOf('MB'),
    };
    var rowsPS = lastRowPS >= 2 ? dispatchPS.getRange(2, 1, lastRowPS - 1, dispatchPS.getLastColumn()).getValues() : [];

    var dateFromPS = body.dateFrom ? new Date(body.dateFrom) : null;
    var dateToPS = body.dateTo ? new Date(body.dateTo) : null;
    var filtered = rowsPS.filter(function (r) {
      if (!r[idx.code]) return false;
      if (Number(r[idx.balance]) < 0) return false; // negative Balance rows are excluded, same rule as scheduling
      if (body.machine && String(r[idx.machine]) !== String(body.machine)) return false;
      if (body.poContains && String(r[idx.po]).indexOf(body.poContains) === -1) return false;
      var poDate = r[idx.poDate];
      if (dateFromPS && poDate instanceof Date && poDate < dateFromPS) return false;
      if (dateToPS && poDate instanceof Date && poDate > dateToPS) return false;
      return true;
    });

    var mastersPS = ss.getSheetByName('Masters').getDataRange().getValues();
    var nameByCodePS = {};
    mastersPS.slice(1).forEach(function (r) { if (r[1]) nameByCodePS[r[1]] = r[2]; });

    var byCodePS = {};
    filtered.forEach(function (r) {
      var code = r[idx.code];
      var color = r[idx.color] || '(no color)';
      if (!byCodePS[code]) {
        byCodePS[code] = { code: code, name: nameByCodePS[code] || code, color: color, orderQty: 0, supplied: 0, rmReq: 0, mb: 0, rows: 0 };
      }
      byCodePS[code].orderQty += Number(r[idx.orderQty]) || 0;
      byCodePS[code].supplied += Number(r[idx.supplied]) || 0;
      byCodePS[code].rmReq += Number(r[idx.rmReq]) || 0;
      byCodePS[code].mb += Number(r[idx.mb]) || 0;
      byCodePS[code].rows += 1;
    });

    var summaryRows = Object.keys(byCodePS).map(function (k) { return byCodePS[k]; })
      .sort(function (a, b) { return a.color.localeCompare(b.color) || (b.orderQty - a.orderQty); });
    var totals = summaryRows.reduce(function (acc, r) {
      acc.orderQty += r.orderQty; acc.supplied += r.supplied; acc.rmReq += r.rmReq; acc.mb += r.mb;
      return acc;
    }, { orderQty: 0, supplied: 0, rmReq: 0, mb: 0 });

    var machinesAvailable = Array.from(new Set(rowsPS.map(function (r) { return r[idx.machine]; }).filter(function (m) { return m !== ''; })));

    result = { products: summaryRows, totals: totals, filterOptions: { machines: machinesAvailable }, rowCount: filtered.length };
  } else if (action === 'installOnEditTrigger') {
    // The bare onEdit(e) function is a SIMPLE trigger — it silently stops firing whenever the
    // script's required authorization scopes change (which has happened repeatedly as this
    // project grew: PropertiesService, more SpreadsheetApp calls, etc.), with no error and no
    // log entry. An INSTALLABLE trigger runs under authorization explicitly granted once here
    // (via the web app's own execution context) and keeps working across future code changes.
    // Remove any existing onEdit installable triggers first to avoid firing twice per edit.
    var existingTriggers = ScriptApp.getProjectTriggers();
    var removed = 0;
    existingTriggers.forEach(function (t) {
      if (t.getHandlerFunction() === 'onEdit') { ScriptApp.deleteTrigger(t); removed++; }
    });
    ScriptApp.newTrigger('onEdit').forSpreadsheet(ss).onEdit().create();
    result = { installed: true, removedExisting: removed };
  } else if (action === 'checkOnEditTriggers') {
    var triggersCk = ScriptApp.getProjectTriggers();
    result = triggersCk.map(function (t) {
      return { handler: t.getHandlerFunction(), eventType: t.getEventType().toString(), source: t.getTriggerSource().toString() };
    });
  } else if (action === 'installCodeDropdown') {
    // Same as setUpCodeDropdown() but callable via the web app (no SpreadsheetApp.getUi() calls,
    // which throw outside a bound UI context).
    var dispatchICD = ss.getSheetByName('Dispatch');
    var mastersICD = ss.getSheetByName('Masters');
    var codesICD = mastersICD.getRange(2, 2, mastersICD.getLastRow() - 1, 1).getValues()
      .map(function (r) { return r[0]; }).filter(function (c) { return c !== ''; });
    var invoiceRowICD = findInvoiceRow_(dispatchICD);
    var dataStartRowICD = invoiceRowICD + 2;
    var lastRowICD = dispatchICD.getLastRow();
    if (lastRowICD < dataStartRowICD) { result = { installed: false, reason: 'no data rows' }; }
    else {
      var ruleICD = SpreadsheetApp.newDataValidation()
        .requireValueInList(codesICD, true)
        .setAllowInvalid(true)
        .setHelpText('Pick a product Code from Masters.')
        .build();
      dispatchICD.getRange(dataStartRowICD, 6, lastRowICD - dataStartRowICD + 1, 1).setDataValidation(ruleICD);
      result = { installed: true, rows: lastRowICD - dataStartRowICD + 1 };
    }
  } else if (action === 'clearColumnValidation') {
    var dispatchCV = ss.getSheetByName('Dispatch');
    var lastColCV = dispatchCV.getLastColumn();
    var lastRowCV = dispatchCV.getLastRow();
    var headersCV = dispatchCV.getRange(1, 1, 1, lastColCV).getValues()[0];
    var idxCV = headersCV.indexOf(body.column);
    if (idxCV === -1) throw new Error('Column not found: ' + body.column);
    if (lastRowCV >= 2) {
      dispatchCV.getRange(2, idxCV + 1, lastRowCV - 1, 1).clearDataValidations();
    }
    result = { cleared: body.column, col: idxCV + 1 };
  } else if (action === 'checkDataValidation') {
    var sheetDV = ss.getSheetByName(body.sheet);
    if (!sheetDV) throw new Error('Sheet not found: ' + body.sheet);
    var lastColDV = sheetDV.getLastColumn();
    var headersDV = sheetDV.getRange(1, 1, 1, lastColDV).getValues()[0];
    var colDV = body.col || (headersDV.indexOf(body.column) + 1);
    if (!colDV) throw new Error('Column not found: ' + body.column);
    var lastRowDV = sheetDV.getLastRow();
    var rulesDV = sheetDV.getRange(2, colDV, lastRowDV - 1, 1).getDataValidations();
    var out = [];
    rulesDV.forEach(function (r, i) {
      var rule = r[0];
      if (!rule) return;
      var crit = rule.getCriteriaType();
      var vals = rule.getCriteriaValues();
      out.push({ row: 2 + i, criteria: crit.toString(), values: vals });
    });
    result = { col: colDV, rulesFound: out.length, samples: out.slice(0, 5) };
  } else if (action === 'fixColumnFormat') {
    // Sets a plain number format on a single named column's data rows, by header lookup —
    // never touches values or any column past body.maxCol (safety bound, e.g. to stay clear
    // of real per-invoice dispatch data columns that must not be modified).
    var dispatchCF = ss.getSheetByName('Dispatch');
    var lastColCF = dispatchCF.getLastColumn();
    var lastRowCF = dispatchCF.getLastRow();
    var maxColCF = body.maxCol || lastColCF;
    var headersCF = dispatchCF.getRange(1, 1, 1, lastColCF).getValues()[0];
    var idxCF = headersCF.indexOf(body.column);
    if (idxCF === -1) throw new Error('Column not found: ' + body.column);
    var colCF = idxCF + 1;
    if (colCF > maxColCF) throw new Error('Column ' + body.column + ' is past the allowed maxCol (' + maxColCF + ') — refusing to touch it.');
    if (lastRowCF >= 2) {
      dispatchCF.getRange(2, colCF, lastRowCF - 1, 1).setNumberFormat(body.format || '@');
    }
    result = { fixed: body.column, col: colCF, format: body.format || '@' };
  } else if (action === 'formatDispatchTable') {
    // Applies the same dark-header style already used on the bar-grid header to every core
    // column header (row 1), plus basic number formats (currency for Rate, plain integer for
    // quantity/balance columns, dd-mmm-yy for dates), frozen header row + first column.
    var dispatchFT = ss.getSheetByName('Dispatch');
    var lastColFT = dispatchFT.getLastColumn();
    var lastRowFT = dispatchFT.getLastRow();
    var headersFT = dispatchFT.getRange(1, 1, 1, lastColFT).getValues()[0];

    dispatchFT.getRange(1, 1, 1, lastColFT)
      .setFontWeight('bold').setBackground('#22201b').setFontColor('#f3ede0')
      .setFontFamily('Trebuchet MS').setFontSize(9)
      .setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true);
    dispatchFT.setFrozenRows(1);
    dispatchFT.setFrozenColumns(1);
    dispatchFT.setRowHeight(1, 30);

    var qtyCols = ['Order Qty', 'Gms', 'Supplied', 'Physical Ready Stock', 'Balance', 'RM Req', 'MB', 'MB Ratio', 'Output / Day (pcs)', 'Days'];
    qtyCols.forEach(function (name) {
      var idx = headersFT.indexOf(name);
      if (idx === -1 || lastRowFT < 2) return;
      dispatchFT.getRange(2, idx + 1, lastRowFT - 1, 1).setNumberFormat('#,##0.###');
    });
    var rateIdx = headersFT.indexOf('Rate');
    if (rateIdx !== -1 && lastRowFT >= 2) {
      dispatchFT.getRange(2, rateIdx + 1, lastRowFT - 1, 1).setNumberFormat('#,##0.00');
    }
    var dateCols = ['PO Date', 'Month', 'Start Date', 'End Date'];
    dateCols.forEach(function (name) {
      var idx = headersFT.indexOf(name);
      if (idx === -1 || lastRowFT < 2) return;
      dispatchFT.getRange(2, idx + 1, lastRowFT - 1, 1).setNumberFormat('dd-mmm-yy');
    });

    result = { formatted: true, lastCol: lastColFT, lastRow: lastRowFT };
  } else if (action === 'fillDispatchFromMasters') {
    // Fills Shape, Pack Size, Color and Gms on every Dispatch data row by looking up its Code
    // in Masters. Both sides are resolved by header name: this handler used to read Masters
    // index 6 as 'color', but index 6 is 'MB Ratio' and COLOR is index 8 — so it was writing
    // the MB ratio into the Color column. It also wrote to fixed Dispatch columns G/H/I/L and
    // started at row 2, ignoring the invoice/label rows the rest of the file accounts for.
    var dispatchFM = ss.getSheetByName('Dispatch');
    var lastRowFM = dispatchFM.getLastRow();
    var labelRowFM = findInvoiceRow_(dispatchFM) + 1;
    var dataStartRowFM = labelRowFM + 1;
    var dHdrFM = dispatchFM.getRange(labelRowFM, 1, 1, dispatchFM.getLastColumn()).getValues()[0];
    var dColFM = function (name) {
      var i = dHdrFM.indexOf(name);
      if (i === -1) throw new Error('Dispatch is missing a "' + name + '" column.');
      return i + 1;
    };

    var mastersFM = ss.getSheetByName('Masters').getDataRange().getValues();
    var mHdrFM = mastersFM[0];
    var mColFM = function (name) { return mHdrFM.indexOf(name); };
    var mcCodeFM = mColFM('Product Code'), mcShapeFM = mColFM('SHAPE');
    var mcPackFM = mColFM('PACK SIZE'), mcGmsFM = mColFM('TARE WEIGHT in gms');
    var mcColorFM = mColFM('COLOR');
    if (mcCodeFM === -1) throw new Error('Masters is missing a "Product Code" column.');
    var takeFM = function (r, i) { return i === -1 ? '' : r[i]; };
    var byCodeFM = {};
    mastersFM.slice(1).forEach(function (r) {
      if (!r[mcCodeFM]) return;
      byCodeFM[r[mcCodeFM]] = {
        shape: takeFM(r, mcShapeFM), packSize: takeFM(r, mcPackFM),
        gms: takeFM(r, mcGmsFM), color: takeFM(r, mcColorFM),
      };
    });

    if (lastRowFM < dataStartRowFM) throw new Error('Dispatch has no data rows to fill.');
    var nRowsFM = lastRowFM - dataStartRowFM + 1;
    var codesFM = dispatchFM.getRange(dataStartRowFM, dColFM('Code'), nRowsFM, 1).getValues();
    // Write each field to its own named column rather than assuming Shape/Pack Size/Color are
    // three adjacent cells — they are today, but that is exactly the assumption that rotted here.
    [['Shape', 'shape'], ['Pack Size', 'packSize'], ['Color', 'color'], ['Gms', 'gms']]
      .forEach(function (pairFM) {
        var valsFM = codesFM.map(function (r) {
          var m = byCodeFM[r[0]];
          return [m ? m[pairFM[1]] : ''];
        });
        dispatchFM.getRange(dataStartRowFM, dColFM(pairFM[0]), valsFM.length, 1).setValues(valsFM);
      });
    var fillsFM = codesFM;

    var notFoundFM = codesFM.filter(function (r) { return r[0] && !byCodeFM[r[0]]; }).map(function (r) { return r[0]; });
    result = { filled: fillsFM.length, notFound: notFoundFM };
  } else if (action === 'fillMachineMbFromMasters') {
    // Fills Machine (A) and MB (col given by body.mbCol) on every Dispatch data row by looking
    // up its Code (F) in Masters (MB Ratio=col G idx6, Machine=col H idx7).
    var dispatchMM = ss.getSheetByName('Dispatch');
    var lastRowMM = dispatchMM.getLastRow();
    var mastersMM = ss.getSheetByName('Masters').getDataRange().getValues();
    var mbRatioColMM = mastersMM[0].indexOf('MB Ratio');
    var machineColMM = mastersMM[0].indexOf('Machine ');
    if (mbRatioColMM === -1 || machineColMM === -1) throw new Error('Masters is missing "MB Ratio" or "Machine " column.');
    var byCodeMM = {};
    mastersMM.slice(1).forEach(function (r) {
      if (r[1]) byCodeMM[r[1]] = { mbRatio: r[mbRatioColMM], machine: r[machineColMM] };
    });

    var codesMM = dispatchMM.getRange(2, 6, lastRowMM - 1, 1).getValues();
    var machineFillsMM = codesMM.map(function (r) {
      var m = byCodeMM[r[0]];
      return [m ? m.machine : ''];
    });
    var mbFillsMM = codesMM.map(function (r) {
      var m = byCodeMM[r[0]];
      return [m ? m.mbRatio : ''];
    });
    dispatchMM.getRange(2, 1, machineFillsMM.length, 1).setValues(machineFillsMM); // A
    dispatchMM.getRange(2, body.mbCol, mbFillsMM.length, 1).setValues(mbFillsMM);

    var notFoundMM = codesMM.filter(function (r) { return r[0] && !byCodeMM[r[0]]; }).map(function (r) { return r[0]; });
    result = { filled: machineFillsMM.length, notFound: notFoundMM };
  } else if (action === 'fixMbFormula') {
    // MB = RM Req x MB Ratio / 100. MB Ratio isn't itself a Dispatch column, so this inserts
    // one right after the existing MB column (body.mbCol), fills it from Masters (col G, MB
    // Ratio) by Code lookup, then installs MB = O{row}*<new MB Ratio col>{row}/100 as a live
    // formula on the pre-existing MB column, replacing whatever static value was there before.
    var dispatchFX = ss.getSheetByName('Dispatch');
    var lastRowFX = dispatchFX.getLastRow();
    var mbCol = body.mbCol; // existing MB column (currently holds a stale static value)
    var ratioCol = mbCol + 1;

    dispatchFX.insertColumnsAfter(mbCol, 1);
    dispatchFX.getRange(1, ratioCol).setValue('MB Ratio').setFontWeight('bold');

    var mastersFX = ss.getSheetByName('Masters').getDataRange().getValues();
    var mbRatioColFX = mastersFX[0].indexOf('MB Ratio');
    if (mbRatioColFX === -1) throw new Error('Masters is missing an "MB Ratio" column.');
    var ratioByCodeFX = {};
    mastersFX.slice(1).forEach(function (r) { if (r[1]) ratioByCodeFX[r[1]] = r[mbRatioColFX]; });

    var codesFX = dispatchFX.getRange(2, 6, lastRowFX - 1, 1).getValues();
    var ratioFillsFX = codesFX.map(function (r) {
      var ratio = ratioByCodeFX[r[0]];
      return [ratio === undefined ? '' : ratio];
    });
    dispatchFX.getRange(2, ratioCol, ratioFillsFX.length, 1).setValues(ratioFillsFX);

    var rmReqCol = mbCol - 1; // RM Req sits immediately before MB in the current layout
    var mbFormulasFX = [];
    for (var rfx = 2; rfx <= lastRowFX; rfx++) {
      var rmReqLetter = dispatchFX.getRange(1, rmReqCol).getA1Notation().replace(/[0-9]/g, '');
      var ratioLetter = dispatchFX.getRange(1, ratioCol).getA1Notation().replace(/[0-9]/g, '');
      mbFormulasFX.push(['=' + rmReqLetter + rfx + '*' + ratioLetter + rfx + '/100']);
    }
    dispatchFX.getRange(2, mbCol, mbFormulasFX.length, 1).setFormulas(mbFormulasFX);

    result = { fixed: mbFormulasFX.length, ratioCol: ratioCol };
  } else if (action === 'writeColumn') {
    var sheetWC = ss.getSheetByName(body.sheet);
    if (!sheetWC) throw new Error('Sheet not found: ' + body.sheet);
    var valuesWC = body.values.map(function (v) { return [v]; });
    sheetWC.getRange(body.startRow, body.col, valuesWC.length, 1).setValues(valuesWC);
    result = { written: valuesWC.length };
  } else if (action === 'installDispatchScheduleColumns') {
    var dispatchISC2 = ss.getSheetByName('Dispatch');
    try {
      installDispatchScheduleColumns_(dispatchISC2);
      result = { installed: true };
    } catch (errISC) {
      result = { installed: false, error: String(errISC), stack: errISC.stack || null };
    }
  } else if (action === 'resetScheduleBlockMarker') {
    PropertiesService.getScriptProperties().deleteProperty('DISPATCH_SCHEDULE_BLOCK_INSERTED');
    result = { reset: true };
  } else if (action === 'deleteColumnsFrom') {
    // body.count: number of columns to delete starting at body.startCol. If omitted, deletes
    // through the sheet's last column (old default behavior) — pass count explicitly to avoid
    // accidentally deleting trailing columns that should be preserved.
    var sheetDC = ss.getSheetByName(body.sheet);
    if (!sheetDC) throw new Error('Sheet not found: ' + body.sheet);
    var lastColDC = sheetDC.getLastColumn();
    var countDC = body.count || (lastColDC - body.startCol + 1);
    if (lastColDC >= body.startCol) {
      sheetDC.deleteColumns(body.startCol, Math.min(countDC, lastColDC - body.startCol + 1));
    }
    result = { deletedFrom: body.startCol, count: countDC, wasLastCol: lastColDC };
  } else if (action === 'deleteSheets') {
    var deleted = [];
    (body.names || []).forEach(function (n) {
      var s = ss.getSheetByName(n);
      if (s) { ss.deleteSheet(s); deleted.push(n); }
    });
    result = { deleted: deleted };
  } else if (action === 'ping') {
    // ponytail: quick liveness probe to confirm which deployment is actually serving requests
    result = { ok: true, time: new Date().toISOString() };
  } else {
    result = { error: 'Unknown action: ' + action };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
