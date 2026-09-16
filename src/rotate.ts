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

// Rotates each included day's roster in place: the last person moves to
// the front, and everyone else shifts down one spot. Blank gaps in the
// original list are left untouched since only the non-blank rows found by
// findDayBlocks are considered. Cells are rotated by their raw values (not
// the normalized display text) so a number, date, or other non-text entry
// keeps its original type instead of being coerced to a string.
function rotateRosters(sheet: GoogleAppsScript.Spreadsheet.Sheet): string[] {
  const blocks = findDayBlocks(sheet);
  const rotated: string[] = [];

  blocks.forEach(function (block) {
    if (!block.included || block.rosterCells.length < 2) {
      return;
    }
    const rawValues = block.rosterCells.map(function (c) {
      return c.raw;
    });
    const rotatedValues = [rawValues[rawValues.length - 1]].concat(
      rawValues.slice(0, rawValues.length - 1)
    );
    writeCellsBatched(sheet, block.rosterCells, rotatedValues, 2);
    rotated.push(block.dayName);
  });

  return rotated;
}
