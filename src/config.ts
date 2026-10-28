// The Config sheet only ever needs columns A (names / section labels) and B
// (days / list values / the start date): the start date and stations list
// live in A (label) + B (value), and the roster table lives in A (name) + B
// (day letter), so every reader shares one range read instead of
// re-scanning the sheet per section.
function readConfigValues(sheet: GoogleAppsScript.Spreadsheet.Sheet): unknown[][] {
  const lastRow = sheet.getLastRow();
  return lastRow > 0 ? sheet.getRange(1, 1, lastRow, 2).getValues() : [];
}

function findStartDateRow(values: unknown[][]): number {
  for (let r = 0; r < values.length; r++) {
    if (normalizeCell(values[r][0]) === START_DATE_LABEL) {
      return r + 1;
    }
  }
  return -1;
}

function readStartDateFromValues(values: unknown[][]): Date {
  const row = findStartDateRow(values);
  if (row === -1) {
    throw new Error(
      "Couldn't find the '" + START_DATE_LABEL + "' row on the Config sheet."
    );
  }
  const value = values[row - 1][1];
  if (value instanceof Date) {
    return value;
  }
  throw new Error(
    "The start date cell doesn't contain a valid date. Please pick a date next to '" +
      START_DATE_LABEL +
      "'."
  );
}

// Finds the "Name"/"Day" header row that anchors the Roster table.
function findRosterHeaderRow(values: unknown[][]): number {
  for (let r = 0; r < values.length; r++) {
    if (
      normalizeCell(values[r][0]) === ROSTER_NAME_HEADER &&
      normalizeCell(values[r][1]) === ROSTER_DAY_HEADER
    ) {
      return r + 1;
    }
  }
  return -1;
}

// Stations are a single list defined once, read from the Stations label's
// own row (so a station typed directly next to the label, matching how the
// start date works, isn't dropped) through the row before the Roster
// table's header row — so, like the roster table, users can insert extra
// rows freely.
function readStationsFromValues(values: unknown[][], rosterHeaderRow: number): string[] {
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
  if (rosterHeaderRow !== -1 && labelRow > rosterHeaderRow) {
    throw new Error(
      "The '" +
        STATIONS_SECTION_LABEL +
        "' section must come before the Roster table (the '" +
        ROSTER_NAME_HEADER +
        "' / '" +
        ROSTER_DAY_HEADER +
        "' header row). Please move it above the roster."
    );
  }

  const blockEnd = rosterHeaderRow !== -1 ? rosterHeaderRow - 1 : values.length;

  const stations: string[] = [];
  const seen: { [name: string]: boolean } = {};
  for (let row = labelRow; row <= blockEnd; row++) {
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

interface RosterEntry {
  row: number;
  name: string;
  rawName: unknown;
  dayIndex: number;
}

// Reads the flat roster table from the row after its "Name"/"Day" header
// through the end of the sheet, skipping blank rows. Each non-blank row
// must have both a name and a valid day letter.
function readRosterEntriesFromValues(values: unknown[][], rosterHeaderRow: number): RosterEntry[] {
  const entries: RosterEntry[] = [];
  for (let row = rosterHeaderRow + 1; row <= values.length; row++) {
    const rawName = values[row - 1][0];
    const rawDay = values[row - 1][1];
    const name = normalizeCell(rawName);
    const dayLetter = normalizeCell(rawDay);

    // Defensive: skip a row that duplicates the header text (e.g. the user
    // accidentally pasted it again) rather than trying to parse it as data.
    if (name === ROSTER_NAME_HEADER && dayLetter === ROSTER_DAY_HEADER) {
      continue;
    }
    if (name === "" && dayLetter === "") {
      continue;
    }
    if (name === "" || dayLetter === "") {
      throw new Error(
        "Row " +
          row +
          " of the roster has only a " +
          (name === "" ? "day" : "name") +
          ". Fill in both the name and the day, or leave the row blank."
      );
    }

    const dayIndex = DAY_LETTER_TO_INDEX[dayLetter.toUpperCase()];
    if (dayIndex === undefined) {
      throw new Error(
        "Row " + row + "'s day ('" + dayLetter + "') isn't valid. Use one letter: M, T, W, R, or F."
      );
    }

    entries.push({ row, name, rawName, dayIndex });
  }
  return entries;
}

function readConfig(sheet: GoogleAppsScript.Spreadsheet.Sheet): WorkbookConfig {
  const values = readConfigValues(sheet);
  const startDate = readStartDateFromValues(values);
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
  const stations = readStationsFromValues(values, rosterHeaderRow);
  const entries = readRosterEntriesFromValues(values, rosterHeaderRow);

  const days: DayConfig[] = [1, 2, 3, 4, 5].map(function (dayIndex) {
    const roster = entries
      .filter(function (e) {
        return e.dayIndex === dayIndex;
      })
      .map(function (e) {
        return e.name;
      });
    return {
      dayIndex: dayIndex,
      dayName: DAY_NAMES[dayIndex],
      included: roster.length > 0,
      roster: roster,
    };
  });

  return { startDate: startDate, stations: stations, days: days };
}
