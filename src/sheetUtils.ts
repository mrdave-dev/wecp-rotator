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

// sheet.clear() clears every cell's content and formatting across the
// sheet's current dimensions without resizing it, so writers are always
// free to getRange() any row/column they need afterwards. Resizing the
// grid down (e.g. to 1x1) here would make those later getRange() calls
// throw, since Apps Script does not auto-grow a sheet to fit a requested
// range.
function clearSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  sheet.clear();
}

function normalizeCell(value: unknown): string {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).trim();
}

function cellValues(cells: DayBlockCell[]): string[] {
  return cells.map(function (c) {
    return c.value;
  });
}

// --- Timezone-safe date helpers -------------------------------------------
//
// Date#getDay()/getDate()/setDate() operate in whatever timezone the Apps
// Script *project* runs in, which can differ from the spreadsheet's own
// timezone (used everywhere we display dates via Utilities.formatDate). To
// keep weekday math and displayed dates consistent, every calendar
// computation below goes through the spreadsheet's timezone explicitly.

function getSpreadsheetTz(): string {
  return getSpreadsheet().getSpreadsheetTimeZone();
}

// ISO weekday returned by "u" is 1 (Monday) .. 7 (Sunday); convert to this
// project's 0 (Sunday) .. 6 (Saturday) scheme used by DAY_NAMES.
function getWeekdayIndex(date: Date, tz: string): number {
  const iso = Number(Utilities.formatDate(date, tz, "u"));
  return iso % 7;
}

// Adds `days` calendar days to `date` as seen in `tz`, independent of the
// Apps Script project's own timezone. The result is anchored at UTC noon so
// that formatting it back in any real-world timezone still lands on the
// intended calendar day.
function addCalendarDays(date: Date, tz: string, days: number): Date {
  const parts = Utilities.formatDate(date, tz, "yyyy-MM-dd").split("-").map(Number);
  const base = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2], 12, 0, 0));
  base.setUTCDate(base.getUTCDate() + days);
  return base;
}

function formatSheetDate(date: Date): string {
  return Utilities.formatDate(date, getSpreadsheetTz(), "M/d/yyyy");
}
