const MASTER_YPP_ID = '1RB_-wCpDvY9IYNMb01IP6mrw1pf9gjxjjeodcJD3ihU';

function inspectTabNames() {
  const ss = SpreadsheetApp.openById(MASTER_YPP_ID);
  const info = ss.getSheets().map(s => ({
    name: s.getName(),
    lastRow: s.getLastRow(),
    lastCol: s.getLastColumn()
  }));
  Logger.log(JSON.stringify(info));
  return info;
}

function inspectOneSheet(name) {
  const ss = SpreadsheetApp.openById(MASTER_YPP_ID);
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

function inspectSpec() { return inspectOneSheet('SPEC'); }
function inspectPrint() { return inspectOneSheet('PRINT'); }

function inspectWorkbook() {
  const ss = SpreadsheetApp.openById(MASTER_YPP_ID);
  const sheets = ss.getSheets();
  const report = sheets.map(sheet => {
    const name = sheet.getName();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const numRowsToRead = Math.min(lastRow, 5);
    const numColsToRead = Math.min(lastCol, 30);
    const sample = (numRowsToRead > 0 && numColsToRead > 0)
      ? sheet.getRange(1, 1, numRowsToRead, numColsToRead).getValues()
      : [];
    return { name, lastRow, lastCol, sample };
  });
  Logger.log(JSON.stringify(report));
  return report;
}
