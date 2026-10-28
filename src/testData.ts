// Development/testing helper: quickly fills the Roster table with a batch
// of sample names so Fill Grid, Rotate, etc. can be exercised against a
// nontrivial roster without typing names in by hand.
const TEST_NAME_POOL: string[] = [
  "Alice",
  "Bob",
  "Carol",
  "Dave",
  "Emma",
  "Frank",
  "Grace",
  "Henry",
  "Ivy",
  "Jack",
  "Kate",
  "Leo",
  "Mia",
  "Noah",
  "Olivia",
  "Paul",
];

const TEST_DAY_LETTERS: string[] = ["M", "T", "W", "R", "F"];

// Appends TEST_NAME_POOL to the Roster table, one row per name, cycling
// through the workday letters (M, T, W, R, F, M, ...) so the names spread
// across every day rather than piling onto one. Existing rows are left
// alone — this only adds rows after whatever's already there.
function populateTestNames(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  const values = readConfigValues(sheet);
  const rosterHeaderRow = findRosterHeaderRow(values);
  if (rosterHeaderRow === -1) {
    throw new Error(
      "Couldn't find the '" +
        ROSTER_NAME_HEADER +
        "' / '" +
        ROSTER_DAY_HEADER +
        "' / '" +
        ROSTER_EMAIL_HEADER +
        "' header row on the Config sheet."
    );
  }

  const startRow = Math.max(rosterHeaderRow, sheet.getLastRow()) + 1;
  const rows = TEST_NAME_POOL.map(function (name, i) {
    return [name, TEST_DAY_LETTERS[i % TEST_DAY_LETTERS.length]];
  });
  sheet.getRange(startRow, 1, rows.length, 2).setValues(rows);
}
