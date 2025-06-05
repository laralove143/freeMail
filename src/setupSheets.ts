/* exported insertConfigSheet, insertSampleSheet */

const applyFormat = (
  sheet: Readonly<GoogleAppsScript.Spreadsheet.Sheet>
): void => {
  const maxRange = sheet.getRange(
    1,
    1,
    sheet.getMaxRows(),
    sheet.getMaxColumns()
  );

  sheet
    .getRange(RANGE.header.all)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sheet.setColumnWidths(1, maxRange.getLastColumn(), 200);
  maxRange.applyRowBanding(BANDING_THEME);
  sheet.setFrozenRows(1);
};

const insertConfigSheet = (
  ss: Readonly<GoogleAppsScript.Spreadsheet.Spreadsheet>
): void => {
  const configSheet = ss.insertSheet(CONFIG_SHEET_NAME, 0);
  configSheet.deleteColumns(3, 24);
  configSheet.deleteRows(3, 998);

  configSheet
    .getRange(RANGE.header.templateSubject)
    .setValue(HEADER_CELL_VALUE.templateSubject);

  configSheet
    .getRange(RANGE.header.subject)
    .setValue(HEADER_CELL_VALUE.subject);

  applyFormat(configSheet);
};

const insertSampleSheet = (
  ss: Readonly<GoogleAppsScript.Spreadsheet.Spreadsheet>
): void => {
  const sampleSheet = ss.insertSheet("Sample Sheet", 1);

  sampleSheet.getRange(RANGE.header.email).setValue(HEADER_CELL_VALUE.email);

  sampleSheet
    .getRange(RANGE.header.placeholder)
    .setValue(HEADER_CELL_VALUE.placeholder);

  applyFormat(sampleSheet);
};
