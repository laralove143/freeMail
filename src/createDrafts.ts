/* exported createDrafts */

const getUi = (): GoogleAppsScript.Base.Ui | null => {
  try {
    return SpreadsheetApp.getUi();
  } catch {
    return null;
  }
};

const createDrafts = (): void => {
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
