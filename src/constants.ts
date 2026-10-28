const MENU_NAME = "WECP";
const CONFIG_SHEET_NAME = "Config";
const SCHEDULE_SHEET_NAME = "Schedule";
const SCHEDULE_SNAPSHOT_PROPERTY = "SCHEDULE_SNAPSHOT_V2";
// Written by the previous version of this script, which always cleared the
// *entire* Schedule sheet before writing (a bare hash string, not the
// {rows,cols,hash} JSON the current version stores). Its presence means
// nothing but script-owned content could have survived on that sheet, so it
// is used once to safely fall back to a full clear when upgrading, then
// deleted.
const LEGACY_SCHEDULE_SNAPSHOT_PROPERTY = "SCHEDULE_SNAPSHOT_V1";
const SCHEDULE_TITLE = "WECP Pre-K Rotation Schedule";
const START_DATE_LABEL = "Start date (first day of the upcoming week):";
const STATIONS_SECTION_LABEL = "Stations (one per row, top-to-bottom schedule order):";
const ROSTER_SECTION_LABEL = "Roster (one row per person per workday — day is M, T, W, R, or F):";
const ROSTER_NAME_HEADER = "Name";
const ROSTER_DAY_HEADER = "Day (M/T/W/R/F)";

// Sunday = 0 .. Saturday = 6, matching Date#getDay().
const DAY_NAMES: string[] = [
  "Sunday",
  "Monday",
  "Tuesday",
  "Wednesday",
  "Thursday",
  "Friday",
  "Saturday",
];

// Maps the single-letter workday codes used in the Roster section to the
// DAY_NAMES index. R (not T) is Thursday, to disambiguate from Tuesday —
// the standard MTWRF weekday-abbreviation convention. Only Monday-Friday
// are supported as roster input.
const DAY_LETTER_TO_INDEX: { [letter: string]: number } = {
  M: 1,
  T: 2,
  W: 3,
  R: 4,
  F: 5,
};

// Number of blank rows reserved under the Stations section header, and
// under the Roster section's Name/Day header, when the Config sheet is
// first created. Users can insert more rows later; section boundaries are
// found dynamically by scanning for the next section's label, not by these
// constants.
const DEFAULT_STATION_ROWS = 8;
const DEFAULT_ROSTER_ROWS = 20;

// Background colors cycled down the Schedule sheet's station-label column.
const STATION_ROW_COLORS: string[] = [
  "#FF0000",
  "#FF9900",
  "#FFFF00",
  "#00FF00",
  "#00FFFF",
  "#9900FF",
];
