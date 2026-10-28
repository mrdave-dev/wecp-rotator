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
// The "Name"/"Day" header row doubles as the Roster table's anchor: readers
// locate the roster section by finding this exact pair of header cells,
// rather than a separate label row above it.
const ROSTER_NAME_HEADER = "Name";
const ROSTER_DAY_HEADER = "Day";

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

// Sized for a schedule that's meant to be printed on a single portrait
// (vertical) letter-size page and read from a standing distance, not just
// viewed on screen. (Apps Script has no API for print orientation/scaling
// itself — that's still a one-time choice in File > Print — but a large,
// high-contrast grid is what makes "Fit to Width" printing actually
// legible from a few feet away.)
const SCHEDULE_TITLE_FONT_SIZE = 26;
const SCHEDULE_GRID_FONT_SIZE = 16;
const SCHEDULE_TITLE_ROW_HEIGHT = 50;
const SCHEDULE_HEADER_ROW_HEIGHT = 46;
const SCHEDULE_BODY_ROW_HEIGHT = 34;
// Floors for autoResizeColumns, which sizes a column to fit its widest
// *unwrapped* line — enough to guarantee "Wednesday" and a date like
// "11/4/2026" each fit on one line at SCHEDULE_GRID_FONT_SIZE without the
// header cell's day-name/date lines wrapping again mid-word.
const SCHEDULE_MIN_STATION_COLUMN_WIDTH = 160;
const SCHEDULE_MIN_DAY_COLUMN_WIDTH = 130;
