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
  rosterSample: string[] | null
): number {
  const headerRow = startRow;
  const headerCell = sheet.getRange(headerRow, 1);
  headerCell.insertCheckboxes();
  headerCell.setValue(included);
  sheet.getRange(headerRow, 2).setValue(DAY_NAMES[dayIndex]).setFontWeight("bold");

  const subHeaderRow = headerRow + 1;
  sheet.getRange(subHeaderRow, 2).setValue(ROSTER_HEADER_LABEL).setFontStyle("italic");

  const listStart = subHeaderRow + 1;
  for (let i = 0; i < DEFAULT_DAY_BLOCK_ROWS; i++) {
    if (rosterSample && i < rosterSample.length) {
      sheet.getRange(listStart + i, 2).setValue(rosterSample[i]);
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

  let row = 5;
  sheet.getRange(row, 1).setValue(STATIONS_SECTION_LABEL).setFontWeight("bold").setFontStyle("italic");
  row += 1;
  const sampleStations = ["Art", "Small Room", "Sensory"];
  for (let i = 0; i < DEFAULT_STATION_ROWS; i++) {
    if (i < sampleStations.length) {
      sheet.getRange(row + i, 2).setValue(sampleStations[i]);
    }
  }
  row += DEFAULT_STATION_ROWS + 1;

  sheet
    .getRange(row, 1)
    .setValue(
      "Check the days you use below, then list who works each day. " +
        "Example: Monday below has Anna, Brenda and Charlie."
    )
    .setFontStyle("italic")
    .setWrap(true);
  row += 1;

  sheet.getRange(row, 1).setValue(INCLUDE_HEADER_LABEL).setFontStyle("italic");
  row += 1;

  for (let dayIndex = 0; dayIndex < DAY_NAMES.length; dayIndex++) {
    if (dayIndex === 1) {
      row = writeDayBlock(sheet, row, dayIndex, true, ["Anna", "Brenda", "Charlie"]);
    } else {
      row = writeDayBlock(sheet, row, dayIndex, false, null);
    }
  }

  sheet.setColumnWidth(1, 260);
  sheet.setColumnWidth(2, 160);
}

function setupWorkbook(): void {
  const ui = SpreadsheetApp.getUi();
  const ss = getSpreadsheet();
  const existingConfig = ss.getSheetByName(CONFIG_SHEET_NAME);
  const alreadySetUp = existingConfig !== null && existingConfig.getLastRow() > 0;

  if (alreadySetUp) {
    const response = ui.alert(
      "Set Up Workbook",
      "This workbook already has a Config sheet. Running Setup again will reset its layout. Any stations, days, or rosters you've already entered will be lost. Continue?",
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) {
      return;
    }
  }

  const configSheet = getOrCreateSheet(CONFIG_SHEET_NAME);
  clearSheet(configSheet);
  buildConfigTemplate(configSheet);

  // The Schedule sheet is left alone if it already exists: it may carry
  // content this script doesn't own (e.g. announcements below the grid),
  // and Fill Grid only ever touches its own generated rectangle anyway.
  const scheduleSheet = getOrCreateSheet(SCHEDULE_SHEET_NAME);
  if (scheduleSheet.getLastRow() === 0) {
    scheduleSheet
      .getRange(1, 1)
      .setValue('Fill in the Config sheet, then use the ' + MENU_NAME + ' menu > "Fill Grid for Upcoming Week".');
  }

  ss.setActiveSheet(configSheet);
  ui.alert(
    "Workbook set up",
    "The Config sheet is ready. List your stations, check the days you use, fill in each day's roster, set a start date, then use the " +
      MENU_NAME +
      " menu to fill the grid.",
    ui.ButtonSet.OK
  );
}
