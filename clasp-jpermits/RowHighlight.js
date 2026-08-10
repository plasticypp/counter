/**
 * Neon row highlight for the active cell (background-based).
 *
 * Fills the active row across Dispatch's CORE columns so the current row stays findable while
 * reading sideways. Row only, core columns only, backgrounds only.
 *
 * Why a background and not a border — the reverse of the obvious choice:
 * a border looks safer because it leaves colour alone, but Apps Script can SET borders and cannot
 * READ them (there is no getBorder()). So a border highlight can never restore what it covered:
 * clearing can only strip the previous row's edges to nothing, taking any real rule with it. That
 * was built, deployed, and observed eating the sheet's horizontal rules — first with a column
 * band, then with the row band alone. Backgrounds are readable via getBackgrounds(), so they can
 * be saved before painting and written back afterwards. Readability, not appearance, is what makes
 * a highlight safe.
 *
 * Why core columns only:
 * from column 23 rightwards the cell FILL is load-bearing data — the per-machine Gantt bars are
 * backgrounds, not decoration. Rather than save and restore those and risk corrupting the schedule
 * on a missed restore, the highlight simply stops before the grid begins. The boundary is detected
 * at runtime from the 'dd/MMM' headers rather than hardcoded, because this sheet's column order has
 * shifted repeatedly and a fixed index would silently start painting over bars.
 *
 * Failure mode, stated honestly: the stash is the only record of the covered colours. If an
 * execution dies between painting and restoring, that one row keeps the highlight until something
 * rewrites it. One row, recoverable with clearRowHighlight() — as against a border highlight, where
 * every crossing was unrecoverable.
 */

var ROW_HL_SHEET_ = 'Dispatch';
var ROW_HL_COLOR_ = '#d9ffb0';                 // light neon-green wash: reads against the sheet's
                                               // own fills without hiding the text on top of it
var ROW_HL_PROP_ = 'ROW_HIGHLIGHT_STATE';      // {row, colors[]} of the row currently painted

/**
 * Simple trigger: fires whenever the selection moves.
 *
 * Runs as the active user with no authorization prompt, which is why state lives in USER
 * properties — two people on the sheet at once must not restore each other's rows.
 *
 * Every failure is swallowed: a simple trigger that throws pops an error toast on a plain cursor
 * move, which is more disruptive than a highlight that quietly fails to draw.
 */
function onSelectionChange(e) {
  try {
    var range = e && e.range ? e.range : SpreadsheetApp.getActiveRange();
    if (!range) return;
    var sheet = range.getSheet();
    if (sheet.getName() !== ROW_HL_SHEET_) return;

    var newRow = range.getRow();
    var props = PropertiesService.getUserProperties();
    var state = null;
    try {
      state = JSON.parse(props.getProperty(ROW_HL_PROP_) || 'null');
    } catch (eParse) {
      state = null;                            // corrupt stash: drop it rather than restore junk
    }
    if (state && state.row === newRow) return; // same row: nothing to redraw

    var width = rowHighlightCoreWidth_(sheet);
    if (width < 1) return;

    // Restore FIRST, then stash the new row. Reversing this order means a fast move stashes the
    // highlight colour itself as the row's "original", and the highlight becomes permanent.
    if (state && state.row > 0 && state.row <= sheet.getMaxRows() &&
        state.colors && state.colors.length === width) {
      sheet.getRange(state.row, 1, 1, width).setBackgrounds([state.colors]);
    }

    // Read what is really there now, so the restore puts back the sheet's own colours rather than
    // a guess. One batched read, one batched write — not one call per cell.
    var target = sheet.getRange(newRow, 1, 1, width);
    var original = target.getBackgrounds()[0];
    var paint = [];
    for (var i = 0; i < width; i++) paint.push(ROW_HL_COLOR_);
    target.setBackgrounds([paint]);

    props.setProperty(ROW_HL_PROP_, JSON.stringify({ row: newRow, colors: original }));
  } catch (err) {
    // Deliberately silent — see above.
  }
}

/**
 * Width of the core block: everything left of the Gantt day grid.
 *
 * The grid is found by its 'dd/MMM' header labels rather than a fixed index — Dispatch's column
 * order has changed more than once, and a hardcoded boundary would eventually start painting over
 * the bars. Falls back to the used width when no grid is present.
 */
function rowHighlightCoreWidth_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return 0;
  var labelRow = findInvoiceRow_(sheet) + 1;
  var headers = sheet.getRange(labelRow, 1, 1, lastCol).getValues()[0];
  for (var c = 0; c < headers.length; c++) {
    if (/^\d{1,2}\/[A-Za-z]{3}$/.test(String(headers[c]).trim())) return c; // cols 1..c
  }
  return lastCol;
}

/**
 * Restore the highlighted row and forget the stash.
 *
 * The escape hatch for a highlight stranded by a failed execution, and the thing to run before any
 * rebuild: installDispatchScheduleColumns_ and repaintChainGantt_ both rewrite backgrounds in bulk,
 * and a live highlight sitting under that write would be baked in as though it were real data.
 */
function clearRowHighlight() {
  var sheet = getSpreadsheet_().getSheetByName(ROW_HL_SHEET_);
  if (!sheet) return;
  var props = PropertiesService.getUserProperties();
  var state = null;
  try {
    state = JSON.parse(props.getProperty(ROW_HL_PROP_) || 'null');
  } catch (eParse) {
    state = null;
  }
  if (state && state.row > 0 && state.row <= sheet.getMaxRows() && state.colors) {
    sheet.getRange(state.row, 1, 1, state.colors.length).setBackgrounds([state.colors]);
  }
  props.deleteProperty(ROW_HL_PROP_);
}
