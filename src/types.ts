interface DayBlockCell {
  row: number;
  value: string;
  // The raw cell value (before String()-normalization) so callers that
  // write a cell back (rotation) can preserve its original type instead of
  // forcing it to text.
  raw: unknown;
}

// A day-of-week's block as located on the Config sheet, with references
// back to the exact rows each list entry came from so callers (rotation)
// can write values back in place.
interface DayBlock {
  dayIndex: number; // 0 = Sunday .. 6 = Saturday
  dayName: string;
  headerRow: number;
  included: boolean;
  rosterCells: DayBlockCell[];
  stationCells: DayBlockCell[];
}

interface DayConfig {
  dayIndex: number;
  dayName: string;
  included: boolean;
  roster: string[];
  stations: string[];
}

interface WorkbookConfig {
  startDate: Date;
  days: DayConfig[];
}

interface ScheduledDay {
  dayIndex: number;
  dayName: string;
  date: Date;
  roster: string[];
  stations: string[];
}

interface ScheduleAssignment {
  orderedDays: ScheduledDay[];
  stationOrder: string[];
  // Keyed by `${station}${dayIndex}` -> assigned person.
  grid: { [key: string]: string };
  errors: string[];
}
