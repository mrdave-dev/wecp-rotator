// Rotates each included day's roster in place: the last person moves to
// the front, and everyone else shifts down one spot. Blank gaps in the
// original list are left untouched since only the non-blank rows found by
// findDayBlocks are considered.
function rotateRosters(sheet: GoogleAppsScript.Spreadsheet.Sheet): string[] {
  const blocks = findDayBlocks(sheet);
  const rotated: string[] = [];

  blocks.forEach(function (block) {
    if (!block.included || block.rosterCells.length < 2) {
      return;
    }
    const values = block.rosterCells.map(function (c) {
      return c.value;
    });
    const rotatedValues = [values[values.length - 1]].concat(
      values.slice(0, values.length - 1)
    );
    block.rosterCells.forEach(function (cell, i) {
      sheet.getRange(cell.row, 2).setValue(rotatedValues[i]);
    });
    rotated.push(block.dayName);
  });

  return rotated;
}
