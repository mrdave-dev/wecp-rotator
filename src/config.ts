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

// Scans column A for a checkbox (the "include this day" toggle) paired with
// a day-of-week name in column B to find each day's block, then reads the
// roster/station list cells between that header and the next one (or the
// end of the sheet). Requiring an actual checkbox in column A (rather than
// matching the day name anywhere in column B) keeps a roster entry that
// happens to be named after a weekday from being mistaken for a header.
// Blank rows inside a block, and the "Roster / Stations" sub-header itself,
// are simply skipped, so users can leave gaps or insert extra rows without
// breaking anything.
function findDayBlocks(sheet: GoogleAppsScript.Spreadsheet.Sheet): DayBlock[] {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    return [];
  }
  const values = sheet.getRange(1, 1, lastRow, 3).getValues();

  const headers: { row: number; dayIndex: number }[] = [];
  for (let r = 0; r < lastRow; r++) {
    if (typeof values[r][0] !== "boolean") {
      continue;
    }
    const dayIndex = findDayNameIndex(normalizeCell(values[r][1]));
    if (dayIndex !== -1) {
      headers.push({ row: r + 1, dayIndex });
    }
  }

  const seenDayIndexes: { [dayIndex: number]: number } = {};
  headers.forEach(function (h) {
    if (seenDayIndexes[h.dayIndex] !== undefined) {
      throw new Error(
        "Found more than one row for " +
          DAY_NAMES[h.dayIndex] +
          " on the Config sheet (rows " +
          seenDayIndexes[h.dayIndex] +
          " and " +
          h.row +
          "). Please remove the duplicate before continuing."
      );
    }
    seenDayIndexes[h.dayIndex] = h.row;
  });

  const blocks: DayBlock[] = [];
  for (let h = 0; h < headers.length; h++) {
    const headerRow = headers[h].row;
    const dayIndex = headers[h].dayIndex;
    const blockEnd = h + 1 < headers.length ? headers[h + 1].row - 1 : lastRow;
    const listStart = headerRow + 1;

    const rosterCells: DayBlockCell[] = [];
    const stationCells: DayBlockCell[] = [];
    for (let row = listStart; row <= blockEnd; row++) {
      const rosterRaw = values[row - 1][1];
      const rosterValue = normalizeCell(rosterRaw);
      if (rosterValue !== "" && rosterValue !== ROSTER_HEADER_LABEL) {
        rosterCells.push({ row, value: rosterValue, raw: rosterRaw });
      }
      const stationRaw = values[row - 1][2];
      const stationValue = normalizeCell(stationRaw);
      if (stationValue !== "" && stationValue !== STATIONS_HEADER_LABEL) {
        stationCells.push({ row, value: stationValue, raw: stationRaw });
      }
    }

    blocks.push({
      dayIndex,
      dayName: DAY_NAMES[dayIndex],
      headerRow,
      included: values[headerRow - 1][0] === true,
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
      roster: cellValues(block.rosterCells),
      stations: cellValues(block.stationCells),
    };
  });
  return { startDate: startDate, days: days };
}
