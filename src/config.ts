function findDayNameIndex(text: string): number {
  const normalized = text.trim().toLowerCase();
  for (let i = 0; i < DAY_NAMES.length; i++) {
    if (DAY_NAMES[i].toLowerCase() === normalized) {
      return i;
    }
  }
  return -1;
}

// The Config sheet only ever needs columns A (labels / include checkboxes)
// and B (list entries: the start date, the stations list, and each day's
// roster all live in these two columns), so every reader shares one range
// read instead of re-scanning the sheet per section.
function readConfigValues(sheet: GoogleAppsScript.Spreadsheet.Sheet): unknown[][] {
  const lastRow = sheet.getLastRow();
  return lastRow > 0 ? sheet.getRange(1, 1, lastRow, 2).getValues() : [];
}

function readStartDateFromValues(values: unknown[][]): Date {
  for (let r = 0; r < values.length; r++) {
    if (normalizeCell(values[r][0]) === START_DATE_LABEL) {
      const value = values[r][1];
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

// A day header row is a checkbox (the "include this day" toggle) in column
// A paired with a day-of-week name in column B. Requiring an actual
// checkbox (rather than matching the day name anywhere in column B) keeps a
// roster entry that happens to be named after a weekday from being mistaken
// for a header. Throws if the same day appears more than once.
function findDayHeaderRows(values: unknown[][]): { row: number; dayIndex: number }[] {
  const headers: { row: number; dayIndex: number }[] = [];
  for (let r = 0; r < values.length; r++) {
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

  return headers;
}

// Reads the roster list cells between each day header and the next one (or
// the end of the sheet). Blank rows inside a block, and the "Roster"
// sub-header itself, are simply skipped, so users can leave gaps or insert
// extra rows without breaking anything.
function findDayBlocksFromValues(values: unknown[][]): DayBlock[] {
  const headers = findDayHeaderRows(values);
  const lastRow = values.length;

  return headers.map(function (header, h) {
    const blockEnd = h + 1 < headers.length ? headers[h + 1].row - 1 : lastRow;
    const listStart = header.row + 1;

    const rosterCells: DayBlockCell[] = [];
    for (let row = listStart; row <= blockEnd; row++) {
      const raw = values[row - 1][1];
      const value = normalizeCell(raw);
      if (value !== "" && value !== ROSTER_HEADER_LABEL) {
        rosterCells.push({ row, value, raw });
      }
    }

    return {
      dayIndex: header.dayIndex,
      dayName: DAY_NAMES[header.dayIndex],
      headerRow: header.row,
      included: values[header.row - 1][0] === true,
      rosterCells,
    };
  });
}

function findDayBlocks(sheet: GoogleAppsScript.Spreadsheet.Sheet): DayBlock[] {
  return findDayBlocksFromValues(readConfigValues(sheet));
}

// Stations are a single list defined once, read from the row after the
// Stations section label up to (but not including) the first day header —
// so, like the day blocks, users can insert extra rows freely.
function readStationsFromValues(values: unknown[][]): string[] {
  let labelRow = -1;
  for (let r = 0; r < values.length; r++) {
    if (normalizeCell(values[r][0]) === STATIONS_SECTION_LABEL) {
      labelRow = r + 1;
      break;
    }
  }
  if (labelRow === -1) {
    throw new Error(
      "Couldn't find the '" + STATIONS_SECTION_LABEL + "' row on the Config sheet."
    );
  }

  const headers = findDayHeaderRows(values);
  const blockEnd = headers.length > 0 ? headers[0].row - 1 : values.length;

  const stations: string[] = [];
  const seen: { [name: string]: boolean } = {};
  for (let row = labelRow + 1; row <= blockEnd; row++) {
    const value = normalizeCell(values[row - 1][1]);
    if (value === "") {
      continue;
    }
    if (seen[value]) {
      throw new Error(
        "The station list has '" + value + "' listed more than once. Station names must be unique."
      );
    }
    seen[value] = true;
    stations.push(value);
  }
  return stations;
}

function readConfig(sheet: GoogleAppsScript.Spreadsheet.Sheet): WorkbookConfig {
  const values = readConfigValues(sheet);
  const startDate = readStartDateFromValues(values);
  const stations = readStationsFromValues(values);
  const blocks = findDayBlocksFromValues(values);
  const days: DayConfig[] = blocks.map(function (block) {
    return {
      dayIndex: block.dayIndex,
      dayName: block.dayName,
      included: block.included,
      roster: cellValues(block.rosterCells),
    };
  });
  return { startDate: startDate, stations: stations, days: days };
}
