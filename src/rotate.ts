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

// Applies `transform` to each included day's roster (as raw cell values, so
// a number, date, or other non-text entry keeps its original type instead
// of being coerced to a string) and writes the result back in place. Shared
// by rotateRosters and unrotateRosters, which differ only in `transform`.
function applyRosterTransform(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  transform: (raw: unknown[]) => unknown[]
): string[] {
  const blocks = findDayBlocks(sheet);
  const affected: string[] = [];

  blocks.forEach(function (block) {
    if (!block.included || block.rosterCells.length < 2) {
      return;
    }
    const rawValues = block.rosterCells.map(function (c) {
      return c.raw;
    });
    writeCellsBatched(sheet, block.rosterCells, transform(rawValues), 2);
    affected.push(block.dayName);
  });

  return affected;
}

// Rotates each included day's roster in place: the last person moves to the
// front, and everyone else shifts down one spot. Blank gaps in the original
// list are left untouched since only the non-blank rows found by
// findDayBlocks are considered.
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
