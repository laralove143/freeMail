/* exported HEADER_CELL_VALUE, BANDING_THEME, getSenderName, getTemplate, getSubject */

const HEADER_CELL_VALUE = {
  email: "Email",
  placeholder: "{YOUR_PLACEHOLDER}",
  subject: "Subject",
  templateSubject: "Template Subject",
};

const CONFIG_SHEET_NAME = "FRee Mail";
const RANGE = {
  email: "A2:A",
  header: {
    all: "1:1",
    email: "A1",
    placeholder: "B1:Z1",
    subject: `${CONFIG_SHEET_NAME}!B1`,
    templateSubject: `${CONFIG_SHEET_NAME}!A1`,
  },
  placeholder: "B2:Z",
  subject: `${CONFIG_SHEET_NAME}!B2`,
  templateSubject: `${CONFIG_SHEET_NAME}!A2`,
};

const BANDING_THEME = SpreadsheetApp.BandingTheme.GREEN;

const getSenderName = (
  ui: Readonly<GoogleAppsScript.Base.Ui> | null
): string => {
  if (!ui) {
    return "Lara Kayaalp";
  }

  const response = ui.prompt("Sender Name:", ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.OK) {
    return response.getResponseText();
  }

  throw new Error("Sender name is required!");
};

const getTemplate = (
  ss: Readonly<GoogleAppsScript.Spreadsheet.Spreadsheet>,
  senderName: string
): string => {
  const subject = ss.getRange(RANGE.templateSubject).getValue() as unknown;

  if (typeof subject !== "string") {
    throw new Error("Failed to get template subject.");
  }

  const threads = GmailApp.search(`subject:"${subject}"`);

  const body = threads[0]
    ?.getMessages()[0]
    ?.getBody()
    .replaceAll("{SENDER}", senderName);

  if (body !== "string") {
    throw new Error(`No mail with subject "${subject}" found!`);
  }

  return body;
};

const getSubject = (
  ss: Readonly<GoogleAppsScript.Spreadsheet.Spreadsheet>
): string => {
  const subject = ss.getRange(RANGE.subject).getValue() as unknown;

  if (typeof subject !== "string") {
    throw new Error("Failed to get subject.");
  }

  return subject;
};
