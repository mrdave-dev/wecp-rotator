function nextMondayOnOrAfter(date: Date, tz: string): Date {
  const weekday = getWeekdayIndex(date, tz); // 0 = Sunday .. 6 = Saturday
  const offset = (1 - weekday + 7) % 7;
  return addCalendarDays(date, tz, offset);
}

function writeDayBlock(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  startRow: number,
  dayIndex: number,
  included: boolean,
  rosterSample: string[] | null,
  stationSample: string[] | null
): number {
  const headerRow = startRow;
  const headerCell = sheet.getRange(headerRow, 1);
  headerCell.insertCheckboxes();
  headerCell.setValue(included);
  sheet.getRange(headerRow, 2).setValue(DAY_NAMES[dayIndex]).setFontWeight("bold");

  const subHeaderRow = headerRow + 1;
  sheet.getRange(subHeaderRow, 2).setValue(ROSTER_HEADER_LABEL).setFontStyle("italic");
  sheet.getRange(subHeaderRow, 3).setValue(STATIONS_HEADER_LABEL).setFontStyle("italic");

  const listStart = subHeaderRow + 1;
  for (let i = 0; i < DEFAULT_DAY_BLOCK_ROWS; i++) {
    const row = listStart + i;
    if (rosterSample && i < rosterSample.length) {
      sheet.getRange(row, 2).setValue(rosterSample[i]);
    }
    if (stationSample && i < stationSample.length) {
      sheet.getRange(row, 3).setValue(stationSample[i]);
    }
  }

  const blockEnd = listStart + DEFAULT_DAY_BLOCK_ROWS - 1;
  return blockEnd + 2; // one blank spacer row before the next block
}

function buildConfigTemplate(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  sheet.getRange(1, 1).setValue("WECP Rotation — Configuration").setFontWeight("bold").setFontSize(12);

  const tz = getSpreadsheetTz();
  sheet.getRange(3, 1).setValue(START_DATE_LABEL).setFontWeight("bold");
  const dateCell = sheet.getRange(3, 2);
  dateCell.setValue(nextMondayOnOrAfter(new Date(), tz));
  dateCell.setNumberFormat("M/d/yyyy");

  sheet
    .getRange(5, 1)
    .setValue(
      "Check the days you use below, then list who works each day and which stations it needs. " +
        "Example: Monday below has Anna, Brenda and Charlie covering Art, Small Room and Sensory."
    )
    .setFontStyle("italic")
    .setWrap(true);

  sheet.getRange(6, 1).setValue(INCLUDE_HEADER_LABEL).setFontStyle("italic");

  let row = 7;
  for (let dayIndex = 0; dayIndex < DAY_NAMES.length; dayIndex++) {
    if (dayIndex === 1) {
      row = writeDayBlock(
        sheet,
        row,
        dayIndex,
        true,
        ["Anna", "Brenda", "Charlie"],
        ["Art", "Small Room", "Sensory"]
      );
    } else {
      row = writeDayBlock(sheet, row, dayIndex, false, null, null);
    }
  }

  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 160);
  sheet.setColumnWidth(3, 160);
}

function setupWorkbook(): void {
  const ui = SpreadsheetApp.getUi();
  const ss = getSpreadsheet();
  const existingConfig = ss.getSheetByName(CONFIG_SHEET_NAME);
  const existingSchedule = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  const alreadySetUp = existingConfig !== null && existingConfig.getLastRow() > 0;

  if (alreadySetUp) {
    const dirty = existingSchedule !== null && isScheduleDirty(existingSchedule);
    const message = dirty
      ? "This workbook already has configuration, and the Schedule sheet has been hand-edited since it was last generated. Running Setup again will erase the Config layout and the current Schedule grid, including those manual edits. Continue?"
      : "This workbook already has a Config sheet. Running Setup again will reset the Config layout and clear the Schedule sheet. Any names or stations you've already typed in will be lost. Continue?";
    const response = ui.alert("Set Up Workbook", message, ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) {
      return;
    }
  }

  const configSheet = getOrCreateSheet(CONFIG_SHEET_NAME);
  clearSheet(configSheet);
  buildConfigTemplate(configSheet);

  const scheduleSheet = getOrCreateSheet(SCHEDULE_SHEET_NAME);
  clearSheet(scheduleSheet);
  scheduleSheet
    .getRange(1, 1)
    .setValue('Fill in the Config sheet, then use the Rotation menu > "Fill Grid for Upcoming Week".');
  PropertiesService.getDocumentProperties().deleteProperty(SCHEDULE_SNAPSHOT_PROPERTY);

  ss.setActiveSheet(configSheet);
  ui.alert(
    "Workbook set up",
    "The Config sheet is ready. Check the days you use, fill in each day's roster and stations, set a start date, then use the Rotation menu to fill the grid.",
    ui.ButtonSet.OK
  );
}
