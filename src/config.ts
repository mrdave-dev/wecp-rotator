function findDayNameIndex(text: string): number {
  const normalized = text.trim().toLowerCase();
  for (let i = 0; i < DAY_NAMES.length; i++) {
    if (DAY_NAMES[i].toLowerCase() === normalized) {
      return i;
    }
  }
  return -1;
}

function readStartDate(sheet: GoogleAppsScript.Spreadsheet.Sheet): Date {
  const lastRow = sheet.getLastRow();
  const labelColumn = sheet.getRange(1, 1, Math.max(lastRow, 1), 1).getValues();
  for (let r = 0; r < labelColumn.length; r++) {
    const label = normalizeCell(labelColumn[r][0]);
    if (label === START_DATE_LABEL) {
      const value = sheet.getRange(r + 1, 2).getValue();
      if (value instanceof Date) {
        return value;
      }
      throw new Error(
        "The start date cell doesn't contain a valid date. Please pick a date next to '" +
          START_DATE_LABEL +
          "'."
      );
    }
  }
  throw new Error(
    "Couldn't find the '" + START_DATE_LABEL + "' row on the Config sheet."
  );
}

// Scans column B for day-of-week names to find each day's block, then reads
// the roster/station list cells between that header and the next one (or
// the end of the sheet). Blank rows inside a block are simply skipped, so
// users can leave gaps or insert extra rows without breaking anything.
function findDayBlocks(sheet: GoogleAppsScript.Spreadsheet.Sheet): DayBlock[] {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    return [];
  }
  const columnA = sheet.getRange(1, 1, lastRow, 1).getValues();
  const columnB = sheet.getRange(1, 2, lastRow, 1).getValues();
  const columnC = sheet.getRange(1, 3, lastRow, 1).getValues();

  const headers: { row: number; dayIndex: number }[] = [];
  for (let r = 0; r < lastRow; r++) {
    const dayIndex = findDayNameIndex(normalizeCell(columnB[r][0]));
    if (dayIndex !== -1) {
      headers.push({ row: r + 1, dayIndex });
    }
  }

  const blocks: DayBlock[] = [];
  for (let h = 0; h < headers.length; h++) {
    const headerRow = headers[h].row;
    const dayIndex = headers[h].dayIndex;
    const blockEnd = h + 1 < headers.length ? headers[h + 1].row - 1 : lastRow;
    // Row headerRow+1 is the "Roster / Stations" sub-header; list entries
    // start the row after that.
    const listStart = headerRow + 2;

    const rosterCells: DayBlockCell[] = [];
    const stationCells: DayBlockCell[] = [];
    for (let row = listStart; row <= blockEnd; row++) {
      const rosterValue = normalizeCell(columnB[row - 1][0]);
      if (rosterValue !== "") {
        rosterCells.push({ row, value: rosterValue });
      }
      const stationValue = normalizeCell(columnC[row - 1][0]);
      if (stationValue !== "") {
        stationCells.push({ row, value: stationValue });
      }
    }

    const includeValue = columnA[headerRow - 1][0];
    blocks.push({
      dayIndex,
      dayName: DAY_NAMES[dayIndex],
      headerRow,
      included: includeValue === true,
      rosterCells,
      stationCells,
    });
  }

  return blocks;
}

function readConfig(sheet: GoogleAppsScript.Spreadsheet.Sheet): WorkbookConfig {
  const startDate = readStartDate(sheet);
  const blocks = findDayBlocks(sheet);
  const days: DayConfig[] = blocks.map(function (block) {
    return {
      dayIndex: block.dayIndex,
      dayName: block.dayName,
      included: block.included,
      roster: block.rosterCells.map(function (c) {
        return c.value;
      }),
      stations: block.stationCells.map(function (c) {
        return c.value;
      }),
    };
  });
  return { startDate: startDate, days: days };
}
