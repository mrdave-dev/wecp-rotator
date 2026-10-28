interface DayBlockCell {
  row: number;
  value: string;
  // The raw cell value (before String()-normalization) so callers that
  // write a cell back (rotation) can preserve its original type instead of
  // forcing it to text.
  raw: unknown;
}

interface DayConfig {
  dayIndex: number;
  dayName: string;
  included: boolean;
  roster: string[];
}

interface WorkbookConfig {
  startDate: Date;
  // Stations are defined once for the whole workbook (not per day) and
  // apply, in this order, to every included day.
  stations: string[];
  days: DayConfig[];
}

interface ScheduledDay {
  dayIndex: number;
  dayName: string;
  date: Date;
}

interface ScheduleAssignment {
  orderedDays: ScheduledDay[];
  stationOrder: string[];
  // Keyed by `${station}${dayIndex}` -> assigned person.
  grid: { [key: string]: string };
  errors: string[];
}

// Records the extent (in rows/cols from A1) and a hash of the grid this
// script last wrote to the Schedule sheet, so a later run can tell whether
// that region has been hand-edited since, without treating unrelated
// content elsewhere on the sheet (e.g. announcements below the grid) as
// part of the generated schedule.
interface ScheduleSnapshot {
  rows: number;
  cols: number;
  hash: string;
}
