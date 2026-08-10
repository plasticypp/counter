/**
 * Neon cross-hair highlight for the active cell.
 *
 * Draws a thick neon outline around the perimeter of the active cell's ROW and its COLUMN, so the
 * cursor stays findable across Dispatch's 93 columns and ~74 rows.
 *
 * Why a border and not a background:
 * on Dispatch the cell BACKGROUND is load-bearing data — the per-machine Gantt bars, the subtotal
 * bands and the dark day-label strip are all fills. Highlighting with a fill means overwriting
 * those and restoring them afterwards, and any missed restore silently destroys schedule colour.
 * A border writes to a different channel entirely, so no fill is ever touched.
 *
 * Why the perimeter and not per-cell:
 * setBorder() on a whole-row range outlines the RANGE, not each cell in it. So one call draws the
 * band and one clears it, whatever the sheet's width — instead of ~93 calls per keypress.
 *
 * The one unavoidable caveat: Apps Script can SET borders but cannot READ them (there is no
 * getBorder()). So the highlight cannot be saved-and-restored the way a colour can; clearing is
 * necessarily "strip this perimeter to nothing". Any deliberate border sitting exactly on a
 * highlighted row's top or bottom edge is removed with the highlight. Confined to two edges of one
 * row, that is a small blast radius, but it is not zero — this is why the feature is scoped to the
 * one sheet below rather than applied workbook-wide.
 */

var ROW_HL_SHEET_ = 'Dispatch';
var ROW_HL_COLOR_ = '#39FF14';                 // neon green: separates from both the pastel bars
                                               // and the near-black (#22201b) day-label strip
var ROW_HL_PROP_ = 'ROW_HIGHLIGHT_LAST_ROW';
var COL_HL_PROP_ = 'ROW_HIGHLIGHT_LAST_COL';

/**
 * Simple trigger: fires whenever the selection moves.
 *
 * Runs as the active user with no authorization prompt, which is why the previous row is stashed
 * in USER properties — two people on the sheet at once must not clear each other's highlight.
 *
 * Every failure here is swallowed. A simple trigger that throws pops an error toast on a plain
 * cursor move, which is far more disruptive than a highlight that quietly fails to draw.
 */
function onSelectionChange(e) {
  try {
    var range = e && e.range ? e.range : SpreadsheetApp.getActiveRange();
    if (!range) return;
    var sheet = range.getSheet();
    if (sheet.getName() !== ROW_HL_SHEET_) return;

    var props = PropertiesService.getUserProperties();
    var lastRow = Number(props.getProperty(ROW_HL_PROP_)) || 0;
    var lastCol = Number(props.getProperty(COL_HL_PROP_)) || 0;
    var newRow = range.getRow();
    var newCol = range.getColumn();
    if (lastRow === newRow && lastCol === newCol) return;   // nothing moved: no redraw

    // Bound the bands to the range actually in use. getMaxRows()/getMaxColumns() would run the
    // outlines out into the empty grid, which looks broken and costs more to draw.
    var width = Math.max(sheet.getLastColumn(), 1);
    var height = Math.max(sheet.getLastRow(), 1);

    // Clear BOTH old bands before drawing EITHER new one. Clearing after drawing would let the
    // old column's clear cut through the new row band where the two cross, leaving a gap in the
    // fresh highlight — the two bands necessarily intersect, so the phases cannot be interleaved.
    if (lastRow > 0 && lastRow <= sheet.getMaxRows()) {
      rowHighlightSetBorder_(sheet, lastRow, width, null);
    }
    if (lastCol > 0 && lastCol <= sheet.getMaxColumns()) {
      colHighlightSetBorder_(sheet, lastCol, height, null);
    }

    rowHighlightSetBorder_(sheet, newRow, width, ROW_HL_COLOR_);
    colHighlightSetBorder_(sheet, newCol, height, ROW_HL_COLOR_);

    props.setProperty(ROW_HL_PROP_, String(newRow));
    props.setProperty(COL_HL_PROP_, String(newCol));
  } catch (err) {
    // Deliberately silent — see above.
  }
}

/**
 * Outline one row's perimeter. Passing color=null clears it.
 *
 * Only top/left/bottom/right are set; the vertical/horizontal INTERIOR arguments are passed false
 * so the call never touches the gridlines between the cells of the row.
 */
function rowHighlightSetBorder_(sheet, row, width, color) {
  var style = color ? SpreadsheetApp.BorderStyle.SOLID_THICK : null;
  sheet.getRange(row, 1, 1, width)
    .setBorder(true, true, true, true, false, false, color, style);
}

/** Outline one column's perimeter. Passing color=null clears it. */
function colHighlightSetBorder_(sheet, col, height, color) {
  var style = color ? SpreadsheetApp.BorderStyle.SOLID_THICK : null;
  sheet.getRange(1, col, height, 1)
    .setBorder(true, true, true, true, false, false, color, style);
}

/**
 * Clear any lingering highlight and forget the stashed row.
 *
 * Needed because the highlight outlives the things it points at: installDispatchScheduleColumns_
 * and repaintChainGantt_ both rewrite the board without touching borders, so a band can be left
 * outlining a row whose contents have since moved. Also the manual escape hatch if a highlight
 * ever gets stranded.
 */
function clearRowHighlight() {
  var sheet = getSpreadsheet_().getSheetByName(ROW_HL_SHEET_);
  if (!sheet) return;
  var props = PropertiesService.getUserProperties();
  var lastRow = Number(props.getProperty(ROW_HL_PROP_)) || 0;
  var lastCol = Number(props.getProperty(COL_HL_PROP_)) || 0;
  if (lastRow > 0 && lastRow <= sheet.getMaxRows()) {
    rowHighlightSetBorder_(sheet, lastRow, Math.max(sheet.getLastColumn(), 1), null);
  }
  if (lastCol > 0 && lastCol <= sheet.getMaxColumns()) {
    colHighlightSetBorder_(sheet, lastCol, Math.max(sheet.getLastRow(), 1), null);
  }
  props.deleteProperty(ROW_HL_PROP_);
  props.deleteProperty(COL_HL_PROP_);
}
