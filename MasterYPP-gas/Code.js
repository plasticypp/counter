function doGet(e) {
  const fn = (e && e.parameter && e.parameter.fn) || 'inspectTabNames';
  const arg = e && e.parameter && e.parameter.arg;
  let result;
  try {
    if (fn === 'inspectTabNames') result = inspectTabNames();
    else if (fn === 'inspectOneSheet') result = inspectOneSheet(arg);
    else if (fn === 'inspectWorkbook') result = inspectWorkbook();
    else result = { error: 'unknown fn: ' + fn };
  } catch (err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function testPing() {
  Logger.log('pong');
  return 'pong';
}

function inspectTabNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const info = ss.getSheets().map(s => ({
    name: s.getName(),
    lastRow: s.getLastRow(),
    lastCol: s.getLastColumn()
  }));
  Logger.log(JSON.stringify(info));
  return info;
}

function inspectOneSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) { Logger.log('NOT FOUND: ' + name); return; }
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const numRows = Math.min(lastRow, 6);
  const numCols = Math.min(lastCol, 20);
  const sample = sheet.getRange(1, 1, numRows, numCols).getValues();
  Logger.log(JSON.stringify({ name, lastRow, lastCol, sample }));
  return { name, lastRow, lastCol, sample };
}

function inspectSpec() { return inspectOneSheet('Spec'); }
function inspectPrint() { return inspectOneSheet('Print'); }
function inspectM2() { return inspectOneSheet('M2'); }

function inspectWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const report = sheets.map(sheet => {
    const name = sheet.getName();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const numRowsToRead = Math.min(lastRow, 5);
    const numColsToRead = Math.min(lastCol, 40);
    const sample = (numRowsToRead > 0 && numColsToRead > 0)
      ? sheet.getRange(1, 1, numRowsToRead, numColsToRead).getValues()
      : [];
    return {
      name,
      lastRow,
      lastCol,
      sample
    };
  });
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}
