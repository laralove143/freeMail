/* exported insertConfigSheet, insertSampleSheet */

const insertConfigSheet = (ss: Spreadsheet): void => {
  const configSheet = ss.insertSheet(CONFIG_SHEET_NAME, 0);

  configSheet
    .getRange(RANGE.header.templateSubject)
    .setValue(HEADER_CELL_VALUE.templateSubject)
    .setBackground(COLOR.header);

  configSheet
    .getRange(RANGE.header.subject)
    .setValue(HEADER_CELL_VALUE.subject)
    .setBackground(COLOR.header);

  configSheet.getRange(RANGE.templateSubject).setBackground(COLOR.cell);
  configSheet.getRange(RANGE.subject).setBackground(COLOR.cell);
};

const insertSampleSheet = (ss: Spreadsheet): void => {
  const sampleSheet = ss.insertSheet("Sample Sheet", 1);

  sampleSheet
    .getRange(RANGE.header.email)
    .setValue(HEADER_CELL_VALUE.email)
    .setBackground(COLOR.header);

  sampleSheet
    .getRange(RANGE.header.placeholder)
    .setValue(HEADER_CELL_VALUE.placeholder)
    .setBackground(COLOR.header);

  sampleSheet.getRange(RANGE.email).setBackground(COLOR.cell);
  sampleSheet.getRange(RANGE.placeholder).setBackground(COLOR.cell);
};
