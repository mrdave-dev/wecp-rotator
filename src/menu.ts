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

// Reads the Config sheet, validates it, and (re)writes the Schedule grid.
// Shared by "Fill Grid" and "Rotate & Fill" so both show the same
// dirty-overwrite confirmation and validation/result alerts. Returns
// whether the grid was actually written.
function fillScheduleFromConfig(
  ui: GoogleAppsScript.Base.Ui,
  configSheet: GoogleAppsScript.Spreadsheet.Sheet
): boolean {
  const scheduleSheet = getOrCreateSheet(SCHEDULE_SHEET_NAME);

  if (isScheduleDirty(scheduleSheet)) {
    const response = ui.alert(
      "Schedule has manual edits",
      "The Schedule grid looks like it was hand-edited since it was last generated. Filling it now will overwrite those changes. Continue?",
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) {
      return false;
    }
  }

  const config = readConfig(configSheet);
  const assignment = computeAssignments(config);

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

  writeSchedule(scheduleSheet, config, assignment);
  getSpreadsheet().setActiveSheet(scheduleSheet);

  ui.alert(
    "Schedule updated",
    "The grid has been filled for the week of " + formatSheetDate(config.startDate) + ".",
    ui.ButtonSet.OK
  );
  return true;
}

function menuFillGrid(): void {
  const ui = SpreadsheetApp.getUi();
  try {
    const configSheet = requireConfigSheet(ui);
    if (!configSheet) {
      return;
    }
    fillScheduleFromConfig(ui, configSheet);
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
    rotateRosters(configSheet);
    fillScheduleFromConfig(ui, configSheet);
  } catch (e) {
    showError(ui, e);
  }
}

function menuRotateRosters(): void {
  const ui = SpreadsheetApp.getUi();
  try {
    const configSheet = requireConfigSheet(ui);
    if (!configSheet) {
      return;
    }
    const rotatedDays = rotateRosters(configSheet);
    if (rotatedDays.length === 0) {
      ui.alert(
        "Nothing rotated",
        "No included day has a roster with more than one person to rotate.",
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        "Rosters rotated",
        "Rotated: " +
          rotatedDays.join(", ") +
          '.\nUse "Fill Grid for Upcoming Week" to apply the new order to the schedule.',
        ui.ButtonSet.OK
      );
    }
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
    const rotatedDays = unrotateRosters(configSheet);
    if (rotatedDays.length === 0) {
      ui.alert(
        "Nothing to unrotate",
        "No included day has a roster with more than one person to unrotate.",
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        "Rosters unrotated",
        "Unrotated: " +
          rotatedDays.join(", ") +
          '.\nUse "Fill Grid for Upcoming Week" to apply the new order to the schedule.',
        ui.ButtonSet.OK
      );
    }
  } catch (e) {
    showError(ui, e);
  }
}
