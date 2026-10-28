// Exports a single sheet as a portrait, fit-to-width PDF via the
// spreadsheet's export URL. Apps Script has no dedicated "export this sheet
// as PDF" API; UrlFetchApp + the script's own OAuth token is the standard
// way to fetch it. gridlines/titles/sheet-name/page-numbers are all turned
// off so the PDF is just the formatted grid.
function exportSheetAsPdf(sheet: GoogleAppsScript.Spreadsheet.Sheet): GoogleAppsScript.Base.Blob {
  const ss = getSpreadsheet();
  const url =
    "https://docs.google.com/spreadsheets/d/" +
    ss.getId() +
    "/export" +
    "?format=pdf" +
    "&gid=" +
    sheet.getSheetId() +
    "&size=letter" +
    "&portrait=true" +
    "&fitw=true" +
    "&gridlines=false" +
    "&printtitle=false" +
    "&sheetnames=false" +
    "&pagenumbers=false" +
    "&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5";
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
  });
  return response.getBlob().setName(SCHEDULE_TITLE + ".pdf");
}

// Every email address on file (config.emailsByName) for everyone actually
// assigned a station somewhere in this week's grid, deduped. People in the
// roster who aren't scheduled this week (e.g. their day wasn't included, or
// there weren't enough stations to reach them) are left out.
function collectRecipientEmails(
  assignment: ScheduleAssignment,
  emailsByName: { [name: string]: string[] }
): string[] {
  const assignedNames: { [name: string]: boolean } = {};
  Object.keys(assignment.grid).forEach(function (key) {
    const name = assignment.grid[key];
    if (name) {
      assignedNames[name] = true;
    }
  });

  const emails: string[] = [];
  Object.keys(assignedNames).forEach(function (name) {
    (emailsByName[name] || []).forEach(function (email) {
      if (emails.indexOf(email) === -1) {
        emails.push(email);
      }
    });
  });
  return emails;
}

function isValidEmail(email: string): boolean {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

// Builds and saves a Gmail draft (not sent) addressed Bcc to every email on
// file for someone scheduled this week, with the Schedule sheet attached as
// a PDF. Throws if nobody scheduled this week has an email on file, or if
// any of their addresses don't look valid.
function createScheduleEmailDraft(
  scheduleSheet: GoogleAppsScript.Spreadsheet.Sheet,
  config: WorkbookConfig,
  assignment: ScheduleAssignment
): GoogleAppsScript.Gmail.GmailDraft {
  const emails = collectRecipientEmails(assignment, config.emailsByName);
  if (emails.length === 0) {
    throw new Error(
      "No email addresses are on file for anyone scheduled this week. Add emails in the Roster table's '" +
        ROSTER_EMAIL_HEADER +
        "' column."
    );
  }
  const badEmails = emails.filter(function (e) {
    return !isValidEmail(e);
  });
  if (badEmails.length > 0) {
    throw new Error(
      "These email addresses don't look valid: " +
        badEmails.join(", ") +
        ". Fix them in the Roster table's '" +
        ROSTER_EMAIL_HEADER +
        "' column."
    );
  }

  const pdf = exportSheetAsPdf(scheduleSheet);
  const subject = SCHEDULE_TITLE + " — Week of " + formatSheetDate(config.startDate);
  const body = "Attached is the rotation schedule for the week of " + formatSheetDate(config.startDate) + ".";

  return GmailApp.createDraft("", subject, body, {
    bcc: emails.join(","),
    attachments: [pdf],
  });
}

// Deep link to open the draft directly in Gmail's web UI. This URL shape
// isn't formally documented by Google, but is the long-standing, widely
// used pattern for linking straight to a specific draft by message ID.
function draftLinkUrl(draft: GoogleAppsScript.Gmail.GmailDraft): string {
  return "https://mail.google.com/mail/u/0/#drafts?compose=" + draft.getMessageId();
}
