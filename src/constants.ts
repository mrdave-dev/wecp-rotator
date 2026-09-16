const CONFIG_SHEET_NAME = "Config";
const SCHEDULE_SHEET_NAME = "Schedule";
const SCHEDULE_SNAPSHOT_PROPERTY = "SCHEDULE_SNAPSHOT_V1";
const START_DATE_LABEL = "Start date (first day of the upcoming week):";
const ROSTER_HEADER_LABEL = "Roster (one person per row)";
const STATIONS_HEADER_LABEL = "Stations (one per row)";
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

// Number of blank rows reserved under each day's "Roster / Stations"
// sub-header when the Config sheet is first created. Users can insert more
// rows inside a day's block later; block boundaries are found dynamically
// by scanning for the next day header, not by this constant.
const DEFAULT_DAY_BLOCK_ROWS = 6;
