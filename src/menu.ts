function onOpen(): void {
  SpreadsheetApp.getUi()
    .createMenu(MENU_NAME)
    .addItem("Set Up Workbook", "menuSetupWorkbook")
    .addSeparator()
    .addItem("Fill Grid for Upcoming Week", "menuFillGrid")
    .addItem("Rotate & Fill for Upcoming Week", "menuRotateAndFill")
    .addSeparator()
    .addItem("Rotate Rosters", "menuRotateRosters")
    .addItem("Unrotate Rosters", "menuUnrotateRosters")
    .addToUi();
}

function requireConfigSheet(
  ui: GoogleAppsScript.Base.Ui
): GoogleAppsScript.Spreadsheet.Sheet | null {
  const sheet = getSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    ui.alert(
      "No Config sheet yet",
      'Use "' + MENU_NAME + ' > Set Up Workbook" first.',
      ui.ButtonSet.OK
    );
    return null;
  }
  return sheet;
}

function showError(ui: GoogleAppsScript.Base.Ui, e: unknown): void {
  const message = e instanceof Error ? e.message : String(e);
  ui.alert("Error", message, ui.ButtonSet.OK);
}

function menuSetupWorkbook(): void {
  const ui = SpreadsheetApp.getUi();
  try {
    setupWorkbook();
  } catch (e) {
    showError(ui, e);
  }
}

// True if it's OK to proceed: either the Schedule grid wasn't hand-edited
// since it was last generated, or the user confirmed overwriting it anyway.
function confirmDirtyOverwrite(
  ui: GoogleAppsScript.Base.Ui,
  scheduleSheet: GoogleAppsScript.Spreadsheet.Sheet
): boolean {
  if (!isScheduleDirty(scheduleSheet)) {
    return true;
  }
  const response = ui.alert(
    "Schedule has manual edits",
    "The Schedule grid looks like it was hand-edited since it was last generated. Filling it now will overwrite those changes. Continue?",
    ui.ButtonSet.YES_NO
  );
  return response === ui.Button.YES;
}

// True if it's OK to proceed: either filling the grid wouldn't grow into
// rows/columns that already hold unrelated content, or the user confirmed
// overwriting it anyway.
function confirmGrowthOverwrite(
  ui: GoogleAppsScript.Base.Ui,
  scheduleSheet: GoogleAppsScript.Spreadsheet.Sheet,
  assignment: ScheduleAssignment
): boolean {
  if (!wouldOverwriteContentOutsideGrid(scheduleSheet, assignment)) {
    return true;
  }
  const response = ui.alert(
    "Content near the grid",
    "This week's grid is larger than last time and would overwrite some rows or columns on the Schedule sheet that already have other content on them (e.g. text below or beside the grid). Continue?",
    ui.ButtonSet.YES_NO
  );
  return response === ui.Button.YES;
}

// True if the assignment can be filled: reports and returns false on either
// a validation error or nothing to fill.
function checkAssignment(ui: GoogleAppsScript.Base.Ui, assignment: ScheduleAssignment): boolean {
  if (assignment.errors.length > 0) {
    ui.alert("Can't fill the grid", assignment.errors.join("\n"), ui.ButtonSet.OK);
    return false;
  }
  if (assignment.orderedDays.length === 0) {
    ui.alert(
      "Nothing to fill",
      "No days are marked as included on the Config sheet. Check at least one day's box and try again.",
      ui.ButtonSet.OK
    );
    return false;
  }
  return true;
}

function menuFillGrid(): void {
  const ui = SpreadsheetApp.getUi();
  try {
    const configSheet = requireConfigSheet(ui);
    if (!configSheet) {
      return;
    }
    const scheduleSheet = getOrCreateSheet(SCHEDULE_SHEET_NAME);
    if (!confirmDirtyOverwrite(ui, scheduleSheet)) {
      return;
    }

    const config = readConfig(configSheet);
    const assignment = computeAssignments(config);
    if (!checkAssignment(ui, assignment)) {
      return;
    }
    if (!confirmGrowthOverwrite(ui, scheduleSheet, assignment)) {
      return;
    }

    writeSchedule(scheduleSheet, assignment);
    getSpreadsheet().setActiveSheet(scheduleSheet);
    ui.alert(
      "Schedule updated",
      "The grid has been filled for the week of " + formatSheetDate(config.startDate) + ".",
      ui.ButtonSet.OK
    );
  } catch (e) {
    showError(ui, e);
  }
}

function menuRotateAndFill(): void {
  const ui = SpreadsheetApp.getUi();
  try {
    const configSheet = requireConfigSheet(ui);
    if (!configSheet) {
      return;
    }
    const scheduleSheet = getOrCreateSheet(SCHEDULE_SHEET_NAME);
    if (!confirmDirtyOverwrite(ui, scheduleSheet)) {
      return;
    }

    // Validate (and get the growth-overwrite confirmation) against the
    // current config *before* rotating anything, so a "No" or a config
    // problem leaves the rosters untouched. Rotation only reorders each
    // day's roster — it never changes counts — so this check's outcome
    // still holds after rotating.
    const preAssignment = computeAssignments(readConfig(configSheet));
    if (!checkAssignment(ui, preAssignment)) {
      return;
    }
    if (!confirmGrowthOverwrite(ui, scheduleSheet, preAssignment)) {
      return;
    }

    rotateRosters(configSheet);

    const config = readConfig(configSheet);
    const assignment = computeAssignments(config);
    writeSchedule(scheduleSheet, assignment);
    getSpreadsheet().setActiveSheet(scheduleSheet);
    ui.alert(
      "Schedule updated",
      "Rosters were rotated and the grid has been filled for the week of " +
        formatSheetDate(config.startDate) +
        ".",
      ui.ButtonSet.OK
    );
  } catch (e) {
    showError(ui, e);
  }
}

// Shared by "Rotate Rosters" and "Unrotate Rosters": applies `action`, then
// reports which days changed (or that there was nothing to do).
function menuRotateOrUnrotate(
  ui: GoogleAppsScript.Base.Ui,
  configSheet: GoogleAppsScript.Spreadsheet.Sheet,
  action: (sheet: GoogleAppsScript.Spreadsheet.Sheet) => string[],
  baseVerb: string,
  pastVerb: string
): void {
  const affectedDays = action(configSheet);
  if (affectedDays.length === 0) {
    ui.alert(
      "Nothing to " + baseVerb,
      "No included day has a roster with more than one person to " + baseVerb + ".",
      ui.ButtonSet.OK
    );
    return;
  }
  ui.alert(
    "Rosters " + pastVerb,
    "Days " +
      pastVerb +
      ": " +
      affectedDays.join(", ") +
      '.\nUse "Fill Grid for Upcoming Week" to apply the new order to the schedule.',
    ui.ButtonSet.OK
  );
}

function menuRotateRosters(): void {
  const ui = SpreadsheetApp.getUi();
  try {
    const configSheet = requireConfigSheet(ui);
    if (!configSheet) {
      return;
    }
    menuRotateOrUnrotate(ui, configSheet, rotateRosters, "rotate", "rotated");
  } catch (e) {
    showError(ui, e);
  }
}

function menuUnrotateRosters(): void {
  const ui = SpreadsheetApp.getUi();
  try {
    const configSheet = requireConfigSheet(ui);
    if (!configSheet) {
      return;
    }
    menuRotateOrUnrotate(ui, configSheet, unrotateRosters, "unrotate", "unrotated");
  } catch (e) {
    showError(ui, e);
  }
}
