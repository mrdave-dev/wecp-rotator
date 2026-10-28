function onOpen(): void {
  SpreadsheetApp.getUi()
    .createMenu(MENU_NAME)
    .addItem("Set Up Workbook", "menuSetupWorkbook")
    .addSeparator()
    .addItem("Fill Grid for Upcoming Week", "menuFillGrid")
    .addSeparator()
    .addItem("Next Week", "menuNextWeek")
    .addItem("Previous Week", "menuPreviousWeek")
    .addSeparator()
    .addItem("Add 16 Test Names", "menuPopulateTestNames")
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
      "No roster rows are filled in on the Config sheet. Add at least one person and day, and try again.",
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

// Shared by "Next Week" and "Previous Week": both advance the Config
// sheet's start date by a week (forward or back), rotate the rosters the
// same direction, and fill the grid for the resulting week — in that
// order, and only once every check has passed, so a "No" or a config
// problem leaves the start date and rosters untouched.
function moveWeek(
  ui: GoogleAppsScript.Base.Ui,
  configSheet: GoogleAppsScript.Spreadsheet.Sheet,
  days: number,
  rotate: (sheet: GoogleAppsScript.Spreadsheet.Sheet) => string[],
  directionLabel: string
): void {
  const scheduleSheet = getOrCreateSheet(SCHEDULE_SHEET_NAME);
  if (!confirmDirtyOverwrite(ui, scheduleSheet)) {
    return;
  }

  // Rotation only reorders each day's roster, and the date shift doesn't
  // change which days/stations are in play, so neither changes the
  // assignment's shape — validating (and getting the growth-overwrite
  // confirmation) against the current config still holds after both.
  const preAssignment = computeAssignments(readConfig(configSheet));
  if (!checkAssignment(ui, preAssignment)) {
    return;
  }
  if (!confirmGrowthOverwrite(ui, scheduleSheet, preAssignment)) {
    return;
  }

  advanceStartDate(configSheet, days);
  rotate(configSheet);

  const config = readConfig(configSheet);
  const assignment = computeAssignments(config);
  writeSchedule(scheduleSheet, assignment);
  getSpreadsheet().setActiveSheet(scheduleSheet);
  ui.alert(
    "Schedule updated",
    "Moved to " +
      directionLabel +
      ": rosters were rotated and the grid has been filled for the week of " +
      formatSheetDate(config.startDate) +
      ".",
    ui.ButtonSet.OK
  );
}

function menuNextWeek(): void {
  const ui = SpreadsheetApp.getUi();
  try {
    const configSheet = requireConfigSheet(ui);
    if (!configSheet) {
      return;
    }
    moveWeek(ui, configSheet, 7, rotateRosters, "next week");
  } catch (e) {
    showError(ui, e);
  }
}

function menuPreviousWeek(): void {
  const ui = SpreadsheetApp.getUi();
  try {
    const configSheet = requireConfigSheet(ui);
    if (!configSheet) {
      return;
    }
    moveWeek(ui, configSheet, -7, unrotateRosters, "previous week");
  } catch (e) {
    showError(ui, e);
  }
}

function menuPopulateTestNames(): void {
  const ui = SpreadsheetApp.getUi();
  try {
    const configSheet = requireConfigSheet(ui);
    if (!configSheet) {
      return;
    }
    populateTestNames(configSheet);
    ui.alert(
      "Test names added",
      "Added " +
        TEST_NAME_POOL.length +
        " test names to the Roster table, spread across M-F. Use \"Fill Grid for Upcoming Week\" to see them in the schedule.",
      ui.ButtonSet.OK
    );
  } catch (e) {
    showError(ui, e);
  }
}
