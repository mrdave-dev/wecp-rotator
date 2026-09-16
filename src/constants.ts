const MENU_NAME = "WECP";
const CONFIG_SHEET_NAME = "Config";
const SCHEDULE_SHEET_NAME = "Schedule";
const SCHEDULE_SNAPSHOT_PROPERTY = "SCHEDULE_SNAPSHOT_V2";
const SCHEDULE_TITLE = "WECP Pre-K Rotation Schedule";
const START_DATE_LABEL = "Start date (first day of the upcoming week):";
const STATIONS_SECTION_LABEL = "Stations (one per row, top-to-bottom schedule order):";
const ROSTER_HEADER_LABEL = "Roster (one person per row)";
const INCLUDE_HEADER_LABEL = "Include this day?";

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

// Number of blank rows reserved under each day's "Roster" sub-header, and
// under the Stations section header, when the Config sheet is first
// created. Users can insert more rows inside a block later; block
// boundaries are found dynamically by scanning for the next section header,
// not by these constants.
const DEFAULT_DAY_BLOCK_ROWS = 6;
const DEFAULT_STATION_ROWS = 8;

// Background colors cycled down the Schedule sheet's station-label column.
const STATION_ROW_COLORS: string[] = [
  "#FF0000",
  "#FF9900",
  "#FFFF00",
  "#00FF00",
  "#00FFFF",
  "#9900FF",
];
