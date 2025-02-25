/* exported main, Sheet, Spreadsheet */

interface Sheet {
  readonly getRange: (range: string) => GoogleAppsScript.Spreadsheet.Range;
}

type Spreadsheet = Readonly<GoogleAppsScript.Spreadsheet.Spreadsheet>;

const getUi = (): GoogleAppsScript.Base.Ui | null => {
  try {
    return SpreadsheetApp.getUi();
  } catch {
    return null;
  }
};

const main = (): void => {
  const ss = SpreadsheetApp.getActive();
  const ui = getUi();

  validate(ss);

  const senderName = getSenderName(ui);
  const template = getTemplate(ss, senderName);
  const subject = getSubject(ss);

  for (const sheet of ss.getSheets()) {
    if (sheet.getName() === CONFIG_SHEET_NAME) {
      continue;
    }

    processSheet(sheet, template, subject);
  }
};
