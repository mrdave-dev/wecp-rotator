// Writes `values[i]` to `cells[i].row` in column `column`, batching
// contiguous runs of rows into a single setValues() call instead of one API
// call per cell.
function writeCellsBatched(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  cells: DayBlockCell[],
  values: unknown[],
  column: number
): void {
  let i = 0;
  while (i < cells.length) {
    let j = i;
    while (j + 1 < cells.length && cells[j + 1].row === cells[j].row + 1) {
      j++;
    }
    const runValues = values.slice(i, j + 1).map(function (v) {
      return [v];
    });
    sheet.getRange(cells[i].row, column, runValues.length, 1).setValues(runValues);
    i = j + 1;
  }
}

// Groups the roster table's rows by day index: each group is the ordered
// list of Name-column cells (with row + raw value) for the rows assigned to
// that day. Rows for different days can be interleaved in the table, so a
// group's rows aren't necessarily contiguous.
function groupRosterCellsByDay(sheet: GoogleAppsScript.Spreadsheet.Sheet): { [dayIndex: number]: DayBlockCell[] } {
  const values = readConfigValues(sheet);
  const rosterHeaderRow = findRosterHeaderRow(values);
  if (rosterHeaderRow === -1) {
    throw new Error(
      "Couldn't find the '" +
        ROSTER_NAME_HEADER +
        "' / '" +
        ROSTER_DAY_HEADER +
        "' header row on the Config sheet."
    );
  }
  const entries = readRosterEntriesFromValues(values, rosterHeaderRow);

  const byDay: { [dayIndex: number]: DayBlockCell[] } = {};
  entries.forEach(function (e) {
    if (!byDay[e.dayIndex]) {
      byDay[e.dayIndex] = [];
    }
    byDay[e.dayIndex].push({ row: e.row, value: e.name, raw: e.rawName });
  });
  return byDay;
}

// Applies `transform` to each day's roster (as raw cell values, so a
// number, date, or other non-text entry keeps its original type instead of
// being coerced to a string) and writes the result back to the Name column
// at the same rows — the Day column is untouched. Shared by rotateRosters
// and unrotateRosters, which differ only in `transform`.
function applyRosterTransform(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  transform: (raw: unknown[]) => unknown[]
): string[] {
  const byDay = groupRosterCellsByDay(sheet);
  const affected: string[] = [];

  Object.keys(byDay).forEach(function (key) {
    const dayIndex = Number(key);
    const cells = byDay[dayIndex];
    if (cells.length < 2) {
      return;
    }
    const rawValues = cells.map(function (c) {
      return c.raw;
    });
    writeCellsBatched(sheet, cells, transform(rawValues), 1);
    affected.push(DAY_NAMES[dayIndex]);
  });

  return affected;
}

// Rotates each day's roster in place: the last person moves to the front,
// and everyone else shifts down one spot.
function rotateRosters(sheet: GoogleAppsScript.Spreadsheet.Sheet): string[] {
  return applyRosterTransform(sheet, function (raw) {
    return [raw[raw.length - 1]].concat(raw.slice(0, raw.length - 1));
  });
}

// The inverse of rotateRosters: the first person moves to the end, and
// everyone else shifts up one spot.
function unrotateRosters(sheet: GoogleAppsScript.Spreadsheet.Sheet): string[] {
  return applyRosterTransform(sheet, function (raw) {
    return raw.slice(1).concat([raw[0]]);
  });
}

// Shifts the Config sheet's start date by `days` (positive to move forward
// a week, negative to move back) and writes it back in place.
function advanceStartDate(sheet: GoogleAppsScript.Spreadsheet.Sheet, days: number): void {
  const values = readConfigValues(sheet);
  const row = findStartDateRow(values);
  if (row === -1) {
    throw new Error(
      "Couldn't find the '" + START_DATE_LABEL + "' row on the Config sheet."
    );
  }
  const current = readStartDateFromValues(values);
  const tz = getSpreadsheetTz();
  const cell = sheet.getRange(row, 2);
  cell.setValue(addCalendarDays(current, tz, days));
  cell.setNumberFormat("M/d/yyyy");
}
