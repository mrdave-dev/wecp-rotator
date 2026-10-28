// Writes `values[i]` to `cells[i].row` in column `column`, batching
// contiguous runs of rows into a single setValues() call instead of one API
// call per cell.
function writeCellsBatched(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  cells: { row: number }[],
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

interface RosterRowCells {
  row: number;
  nameRaw: unknown;
  emailRaw: unknown;
}

// Groups the roster table's rows by day index: each group is the ordered
// list of rows (with row number + raw Name/Email values) assigned to that
// day. Rows for different days can be interleaved in the table, so a
// group's rows aren't necessarily contiguous.
function groupRosterRowsByDay(sheet: GoogleAppsScript.Spreadsheet.Sheet): { [dayIndex: number]: RosterRowCells[] } {
  const values = readConfigValues(sheet);
  const rosterHeaderRow = findRosterHeaderRow(values);
  if (rosterHeaderRow === -1) {
    throw new Error(
      "Couldn't find the '" +
        ROSTER_NAME_HEADER +
        "' / '" +
        ROSTER_DAY_HEADER +
        "' / '" +
        ROSTER_EMAIL_HEADER +
        "' header row on the Config sheet."
    );
  }
  const entries = readRosterEntriesFromValues(values, rosterHeaderRow);

  const byDay: { [dayIndex: number]: RosterRowCells[] } = {};
  entries.forEach(function (e) {
    if (!byDay[e.dayIndex]) {
      byDay[e.dayIndex] = [];
    }
    byDay[e.dayIndex].push({ row: e.row, nameRaw: e.rawName, emailRaw: e.rawEmail });
  });
  return byDay;
}

// Applies `transform` to each day's roster — as raw cell values, so a
// number, date, or other non-text entry keeps its original type instead of
// being coerced to a string — and writes the result back to the Name and
// Email columns at the same rows, moving each row's Email along with its
// Name so the two stay paired. The Day column is untouched. Shared by
// rotateRosters and unrotateRosters, which differ only in `transform`.
function applyRosterTransform(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  transform: (raw: unknown[]) => unknown[]
): string[] {
  const byDay = groupRosterRowsByDay(sheet);
  const affected: string[] = [];

  Object.keys(byDay).forEach(function (key) {
    const dayIndex = Number(key);
    const rows = byDay[dayIndex];
    if (rows.length < 2) {
      return;
    }
    const nameRaws = rows.map(function (r) {
      return r.nameRaw;
    });
    const emailRaws = rows.map(function (r) {
      return r.emailRaw;
    });
    writeCellsBatched(sheet, rows, transform(nameRaws), 1);
    writeCellsBatched(sheet, rows, transform(emailRaws), 3);
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
