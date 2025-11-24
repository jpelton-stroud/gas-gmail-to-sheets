interface AppConfig {
  SHEET_NAME: string;
  FILE_TYPE: ".xlsx" | ".csv";
  CONVERT_TO_SHEET: boolean;
  QUERY_SUBJECT: string;
  SENDER: string;
  GMAIL_QUERY(): string;
}

const CONFIG: AppConfig = {
  SHEET_NAME: "Data",
  SENDER: "",
  FILE_TYPE: ".csv",
  QUERY_SUBJECT: "",
  CONVERT_TO_SHEET: true,
  GMAIL_QUERY() {
    return `subject:(${this.QUERY_SUBJECT} from:(${this.SENDER}) has:attachment newer_than:1d)`;
  },
};

const runScript = () => {
  const msgs: GoogleAppsScript.Gmail.GmailMessage[] = fetchMessages(
    CONFIG.GMAIL_QUERY()
  );
  Logger.log(msgs.length + " messages fetched");

  const attachments: GoogleAppsScript.Gmail.GmailAttachment[] =
    fetchAttachments(msgs);
  Logger.log(attachments.length + " attachments fetched");

  if (attachments.length > 0) {
    const targetSheet: GoogleAppsScript.Spreadsheet.Sheet | null =
      SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);

    if (!targetSheet) {
      Logger.log(CONFIG.SHEET_NAME + " not found.");
      return;
    }

    const tempFileRefs: GoogleAppsScript.Drive.File[] = createTempFiles(
      attachments,
      CONFIG.FILE_TYPE,
      CONFIG.CONVERT_TO_SHEET
    );

    tempFileRefs.forEach((file: any) => {
      Logger.log("Updating Sheet");
      const newData = SpreadsheetApp.openById(file.id)
        .getSheets()[0]
        .getDataRange()
        .getValues();

      overwriteSheet(targetSheet, newData);
      // appendToSheet(targetSheet, newData);
    });

    Logger.log("Updates complete, exiting script");
    return;
  } else {
    Logger.log("No attachments to process, exiting script");
    return;
  }
};

const fetchMessages = (
  query: string
): GoogleAppsScript.Gmail.GmailMessage[] => {
  Logger.log("Fetching Messages");
  const threads: GoogleAppsScript.Gmail.GmailThread[] = GmailApp.search(query);
  return threads.flatMap((thread) => thread.getMessages());
};

const fetchAttachments = (
  msgs: GoogleAppsScript.Gmail.GmailMessage[]
): GoogleAppsScript.Gmail.GmailAttachment[] => {
  Logger.log("Fetching Attachments");
  return msgs.flatMap((msg) => {
    return msg.getAttachments({ includeInlineImages: false });
  });
};

const createTempFiles = (
  files: GoogleAppsScript.Gmail.GmailAttachment[],
  ext?: string,
  convert?: boolean
): GoogleAppsScript.Drive.File[] => {
  Logger.log("Creating Temp Files");
  if (convert === undefined) convert = true;
  if (ext === undefined) ext = ".csv";
  return files
    .filter((file) => file.getName().endsWith(ext))
    .map((file) =>
      (Drive.Files as any).insert(
        { title: file.getName(), labels: { trashed: true } },
        file,
        { convert: convert }
      )
    ) as GoogleAppsScript.Drive.File[];
};

const overwriteSheet = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  data: any[][]
) => {
  Logger.log("Overwriting data to sheet");
  sheet
    .clearContents()
    .getRange(1, 1, data.length, data[0].length)
    .setValues(data);
  Logger.log("Done.");
};

const appendToSheet = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  data: any[][]
) => {
  Logger.log("Appending data to sheet");
  data.shift(); // Drop header row
  const currentLastRow = sheet.getLastRow();
  sheet
    .getRange(currentLastRow + 1, 1, data.length, data[0].length)
    .setValues(data);
  Logger.log("Done.");
};
