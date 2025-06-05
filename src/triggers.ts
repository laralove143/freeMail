/* exported onInstall */

const onOpen = (): void => {
  const ui = SpreadsheetApp.getUi();

  ui.createAddonMenu()
    .addItem("Set up spreadsheet", "onInstall")
    .addItem("Create drafts", "createDrafts")
    .addToUi();
};

const onInstall = (): void => {
  const ss = SpreadsheetApp.getActive();

  insertConfigSheet(ss);
  insertSampleSheet(ss);

  onOpen();
};
