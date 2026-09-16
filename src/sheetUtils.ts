function getSpreadsheet(): GoogleAppsScript.Spreadsheet.Spreadsheet {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet(name: string): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = getSpreadsheet();
  const existing = ss.getSheetByName(name);
  if (existing) {
    return existing;
  }
  return ss.insertSheet(name);
}

function clearSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  sheet.clear();
  sheet.clearFormats();
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  if (maxRows > 1) {
    sheet.deleteRows(2, maxRows - 1);
  }
  if (maxCols > 1) {
    sheet.deleteColumns(2, maxCols - 1);
  }
}

function normalizeCell(value: unknown): string {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).trim();
}
