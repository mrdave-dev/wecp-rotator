function nextMondayOnOrAfter(date: Date, tz: string): Date {
  const weekday = getWeekdayIndex(date, tz); // 0 = Sunday .. 6 = Saturday
  const offset = (1 - weekday + 7) % 7;
  return addCalendarDays(date, tz, offset);
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
      "List who works each day below: one row per person per workday. Day is one letter " +
        "— M, T, W, R, or F — so someone working multiple days gets one row per day. " +
        "Email is optional; list more than one for the same person with a semicolon, " +
        "e.g. abc@example.com;def@example.com."
    )
    .setFontStyle("italic")
    .setWrap(true);
  row += 1;
  sheet.getRange(row, 1).setValue(ROSTER_NAME_HEADER).setFontWeight("bold");
  sheet.getRange(row, 2).setValue(ROSTER_DAY_HEADER).setFontWeight("bold");
  sheet.getRange(row, 3).setValue(ROSTER_EMAIL_HEADER).setFontWeight("bold");
  row += 1;

  const sampleRoster: [string, string, string][] = [
    ["Adam", "M", "adam@example.com"],
    ["Brenda", "T", "brenda@example.com;brenda2@example.com"],
    ["Charlie", "R", ""],
    ["Diana", "T", "diana@example.com"],
  ];
  for (let i = 0; i < DEFAULT_ROSTER_ROWS; i++) {
    if (i < sampleRoster.length) {
      sheet.getRange(row + i, 1).setValue(sampleRoster[i][0]);
      sheet.getRange(row + i, 2).setValue(sampleRoster[i][1]);
      if (sampleRoster[i][2] !== "") {
        sheet.getRange(row + i, 3).setValue(sampleRoster[i][2]);
      }
    }
  }

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 260);
}

function setupWorkbook(): void {
  const ui = SpreadsheetApp.getUi();
  const ss = getSpreadsheet();
  const existingConfig = ss.getSheetByName(CONFIG_SHEET_NAME);
  const alreadySetUp = existingConfig !== null && existingConfig.getLastRow() > 0;

  if (alreadySetUp) {
    const response = ui.alert(
      "Set Up Workbook",
      "This workbook already has a Config sheet. Running Setup again will reset its layout. Any stations or roster rows you've already entered will be lost. Continue?",
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
    "The Config sheet is ready. List your stations, set a start date, then list who works each day (one row per person per workday) using the " +
      MENU_NAME +
      " menu to fill the grid.",
    ui.ButtonSet.OK
  );
}
