/**
 * Chain Start/End dates down each machine.
 *
 * Rule, as specified by the planner:
 *   - Rows are read in SHEET ORDER, top to bottom. Nothing is re-sorted.
 *   - The FIRST row of each machine keeps its own typed Start Date — it is the anchor.
 *   - Every following row on that machine starts where the previous row ended.
 *   - End = Start + Days. Days is read from the sheet's own Days column and never recomputed.
 *   - A new machine value restarts the chain from that machine's own first row.
 *
 * Only Start Date and End Date are written. Days, Balance, Output/Day and the Gantt block are
 * left exactly as they are.
 *
 * Why this is separate from installDispatchScheduleColumns_ rather than a change to it:
 * that function runs from onEdit on every Machine/Priority/Balance/Start edit, recomputes Days
 * from Masters, filters to `balance > 0`, and honours ANY typed date as a pin. Those are
 * different rules, and both writing the same two columns would fight through the
 * DISPATCH_SCHEDULE_WRITTEN_STARTS record it uses to tell its own output from planner edits.
 * This one is manual, previewed, and writes only when the planner clicks Apply.
 */

var CHAIN_DAY_MS_ = 24 * 3600 * 1000;

/** Entry point: Dispatch Tools -> Chain Start/End Dates by Machine. */
function showChainDatesDialog() {
  var html = HtmlService.createHtmlOutputFromFile('ChainDatesDialog')
    .setWidth(720)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Chain Start/End Dates by Machine');
}

/**
 * Resolve the Dispatch columns this feature needs, BY HEADER NAME.
 * The schedule block has moved relative to Month before, and Dispatch's own column order has
 * shifted more than once — a fixed index silently reads the wrong field when that happens.
 */
function chainDatesColumns_(dispatch) {
  var labelRow = findInvoiceRow_(dispatch) + 1;
  var headers = dispatch.getRange(labelRow, 1, 1, dispatch.getLastColumn()).getValues()[0];
  var idx = {
    machine: headers.indexOf('Machine'),
    priority: headers.indexOf('Priority'),
    code: headers.indexOf('Code'),
    days: headers.indexOf('Days'),
    start: headers.indexOf('Start Date'),
    end: headers.indexOf('End Date'),
  };
  var missing = Object.keys(idx).filter(function (k) { return idx[k] === -1; });
  if (missing.length) {
    throw new Error('Dispatch is missing column(s): ' + missing.join(', ') +
      '. Run "Rebuild Schedule Now" first if the schedule block has not been installed.');
  }
  idx.labelRow = labelRow;
  idx.dataStartRow = labelRow + 1;
  return idx;
}

/**
 * Compute the chain and return a preview. Writes NOTHING.
 * Returns { rows: [...], summary: {...} } where each row reports old and new dates so the
 * dialog can show exactly what would change before anything is committed.
 */
function previewChainedDates() {
  var dispatch = getSpreadsheet_().getSheetByName('Dispatch');
  if (!dispatch) throw new Error('No sheet named "Dispatch" found.');

  var idx = chainDatesColumns_(dispatch);
  var lastRow = dispatch.getLastRow();
  if (lastRow < idx.dataStartRow) return { rows: [], summary: emptyChainSummary_() };

  var values = dispatch.getRange(idx.dataStartRow, 1, lastRow - idx.dataStartRow + 1,
    dispatch.getLastColumn()).getValues();

  var cursor = {};          // machine -> end date of the previous row on that machine
  var out = [];
  var today = toUtcMidnightIst_(new Date());

  values.forEach(function (r, i) {
    var machine = Number(r[idx.machine]) || 0;
    var code = r[idx.code];
    // Rows with no machine are unplanned — they have no place in any machine's chain. Skipped
    // rather than defaulted, and counted in the summary so the skip is visible, not silent.
    if (!machine || !code) return;

    var days = Number(r[idx.days]);
    var daysKnown = isFinite(days) && days > 0;
    var oldStart = r[idx.start] instanceof Date ? toUtcMidnightIst_(r[idx.start]) : null;
    var oldEnd = r[idx.end] instanceof Date ? toUtcMidnightIst_(r[idx.end]) : null;

    var isAnchor = !(machine in cursor);
    var start;
    if (isAnchor) {
      // The machine's first row anchors on its OWN typed Start. With none typed there is
      // nothing to anchor to, so fall back to today — skipping would leave the whole machine
      // unscheduled because one cell was blank.
      start = oldStart || today;
    } else {
      start = cursor[machine];
    }
    // Blank or zero Days chains as zero-length rather than being dropped: dropping is what
    // silently hides rows, and the row still needs a position in the sequence. Flagged below.
    var end = new Date(start.getTime() + (daysKnown ? days : 0) * CHAIN_DAY_MS_);
    cursor[machine] = end;

    out.push({
      row: idx.dataStartRow + i,
      machine: machine,
      priority: r[idx.priority] === '' ? '' : r[idx.priority],
      code: String(code),
      days: daysKnown ? days : null,
      isAnchor: isAnchor,
      anchorFromToday: isAnchor && !oldStart,
      oldStart: oldStart ? oldStart.toISOString() : null,
      oldEnd: oldEnd ? oldEnd.toISOString() : null,
      newStart: start.toISOString(),
      newEnd: end.toISOString(),
      changed: !oldStart || !oldEnd ||
        oldStart.getTime() !== start.getTime() || oldEnd.getTime() !== end.getTime(),
    });
  });

  return { rows: out, summary: summariseChain_(out, values.length) };
}

function emptyChainSummary_() {
  return { scanned: 0, chained: 0, changed: 0, unchanged: 0, anchors: 0, noDays: 0, anchorFromToday: 0 };
}

function summariseChain_(out, scanned) {
  return {
    scanned: scanned,
    chained: out.length,
    changed: out.filter(function (r) { return r.changed; }).length,
    unchanged: out.filter(function (r) { return !r.changed; }).length,
    anchors: out.filter(function (r) { return r.isAnchor; }).length,
    noDays: out.filter(function (r) { return r.days === null; }).length,
    anchorFromToday: out.filter(function (r) { return r.anchorFromToday; }).length,
  };
}

/**
 * Recompute the chain server-side and write Start/End for the rows that change.
 *
 * The preview is deliberately NOT trusted as the source of truth for what to write: the sheet
 * can be edited between preview and Apply, and writing a stale plan would overwrite an edit the
 * planner just made. Recomputing here means the write always reflects the sheet as it is now.
 * The row count is compared against the preview so a changed sheet is reported rather than
 * silently applied.
 */
function applyChainedDates(expectedRowCount) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) {
    throw new Error('The sheet is busy with another operation. Try again in a moment.');
  }
  try {
    var dispatch = getSpreadsheet_().getSheetByName('Dispatch');
    if (!dispatch) throw new Error('No sheet named "Dispatch" found.');

    var preview = previewChainedDates();
    if (expectedRowCount != null && preview.rows.length !== expectedRowCount) {
      throw new Error('The sheet changed since the preview (' + preview.rows.length +
        ' schedulable rows now, ' + expectedRowCount + ' when previewed). ' +
        'Nothing was written — reopen the dialog to see the current plan.');
    }

    var idx = chainDatesColumns_(dispatch);
    var toWrite = preview.rows.filter(function (r) { return r.changed; });
    if (!toWrite.length) return { written: 0, summary: preview.summary };

    // Start Date and End Date are adjacent in the schedule block, but that is a layout fact
    // that has changed before — so write them as two ranges resolved by name rather than one
    // two-wide range that would silently straddle the wrong columns if they ever separate.
    // Contiguous runs are batched; setValues per run keeps this to a handful of calls instead
    // of one per row.
    writeChainColumn_(dispatch, idx.start + 1, toWrite, 'newStart');
    writeChainColumn_(dispatch, idx.end + 1, toWrite, 'newEnd');

    // Repaint the Gantt bars from the dates just written. Without this the dates change and the
    // plot does not move: installDispatchScheduleColumns_ paints the bars from ITS OWN computed
    // start/end, never from the sheet's cells, so the only other way to refresh them is
    // "Rebuild Schedule Now" — which would recompute the dates under different rules and undo
    // the chain. This paints from the cells, so what is plotted is what the sheet says.
    var painted = repaintChainGantt_(dispatch, idx, preview.rows);

    return { written: toWrite.length, summary: preview.summary, painted: painted };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Repaint the Dispatch Gantt bars from the chained dates.
 *
 * The bar grid is the block of day-labelled columns ('dd/MMM' text headers) that
 * installDispatchScheduleColumns_ creates after the schedule block. Its position and its day
 * labels are READ from the sheet rather than recomputed: the block has moved before, and its
 * first column is anchored on the day it was last built, not on today. Re-deriving the labels
 * here would silently repaint against a different calendar than the one the header shows.
 *
 * Returns a count of rows painted, and of rows whose span falls outside the visible window —
 * a chained row can start before the grid begins, and a bar that simply is not drawn would
 * otherwise look like a row that was not scheduled.
 */
function repaintChainGantt_(dispatch, idx, rows) {
  var MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  var headers = dispatch.getRange(idx.labelRow, 1, 1, dispatch.getLastColumn()).getValues()[0];

  // Find the run of 'dd/MMM' headers that follows the End Date column.
  var firstCol = -1, dates = [];
  for (var c = idx.end + 1; c < headers.length; c++) {
    var m = String(headers[c]).trim().match(/^(\d{1,2})[\/\- ]([A-Za-z]{3})/);
    if (!m) { if (firstCol !== -1) break; continue; }
    var mon = MONTHS.indexOf(m[2].toLowerCase());
    if (mon === -1) { if (firstCol !== -1) break; continue; }
    if (firstCol === -1) firstCol = c + 1;                 // 1-indexed
    dates.push({ day: Number(m[1]), mon: mon });
  }
  if (firstCol === -1 || !dates.length) return { painted: 0, offGrid: 0, noGrid: true };

  // The headers carry no year. Anchor the first column on the year that puts it nearest today,
  // then walk forward one day per column — the same rule getDashboardData_ uses for gridStart.
  var now = new Date();
  var anchor = null;
  [now.getFullYear() - 1, now.getFullYear(), now.getFullYear() + 1].forEach(function (y) {
    var cand = toUtcMidnightIst_(new Date(y, dates[0].mon, dates[0].day));
    if (!anchor || Math.abs(cand - now) < Math.abs(anchor - now)) anchor = cand;
  });
  var gridStart = anchor.getTime();
  var gridEnd = gridStart + dates.length * CHAIN_DAY_MS_;

  var palette = {
    1: '#9dc3e6', 2: '#a8d8b9', 3: '#f4b183', 4: '#c5a0dc',
    5: '#ff9999', 6: '#85d1e8', 7: '#e0e0a0',
  };

  var lastRow = dispatch.getLastRow();
  var height = lastRow - idx.dataStartRow + 1;
  if (height < 1) return { painted: 0, offGrid: 0, noGrid: false };

  // One read-modify-write over the whole block. Painting cell by cell (as the original does)
  // costs a setBackground call per bar-day — hundreds of round trips, and slow enough to hit
  // the execution limit on a full board.
  var grid = [];
  for (var r = 0; r < height; r++) {
    var blank = [];
    for (var k = 0; k < dates.length; k++) blank.push(null);
    grid.push(blank);
  }

  var painted = 0, offGrid = 0;
  rows.forEach(function (it) {
    var s = new Date(it.newStart).getTime();
    var e = new Date(it.newEnd).getTime();
    var rowIdx = it.row - idx.dataStartRow;
    if (rowIdx < 0 || rowIdx >= height) return;
    if (e <= gridStart || s >= gridEnd) { offGrid++; return; }
    var color = palette[it.machine] || '#e0e0e0';
    var any = false;
    for (var i = 0; i < dates.length; i++) {
      var dayStart = gridStart + i * CHAIN_DAY_MS_;
      var dayEnd = dayStart + CHAIN_DAY_MS_;
      // End Date is EXCLUSIVE — the same convention the scheduler and the dashboard already
      // use (`dayEnd > start && dayStart < end`), so a 1-day job fills exactly one column.
      if (dayEnd > s && dayStart < e) { grid[rowIdx][i] = color; any = true; }
    }
    if (any) painted++;
  });

  dispatch.getRange(idx.dataStartRow, firstCol, height, dates.length).setBackgrounds(grid);
  return { painted: painted, offGrid: offGrid, noGrid: false };
}

/** Write one date field into one column, batching contiguous rows into single setValues calls. */
function writeChainColumn_(dispatch, column1Indexed, rows, field) {
  var i = 0;
  while (i < rows.length) {
    var j = i;
    while (j + 1 < rows.length && rows[j + 1].row === rows[j].row + 1) j++;
    var block = rows.slice(i, j + 1).map(function (r) { return [new Date(r[field])]; });
    dispatch.getRange(rows[i].row, column1Indexed, block.length, 1)
      .setValues(block)
      .setNumberFormat('dd-mmm-yy');
    i = j + 1;
  }
}
