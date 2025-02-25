/* exported validate */

const validateCellEq = (sheet: Sheet, range: string, text: string): void => {
  const rule = SpreadsheetApp.newDataValidation()
    .requireTextEqualTo(text)
    .setAllowInvalid(false)
    .setHelpText(`This cell's value must be "${text}".`)
    .build();

  sheet.getRange(range).setDataValidation(rule);
};

const validateCellNotEmpty = (sheet: Sheet, range: string): void => {
  const rule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied("=LEN(INDIRECT(ADDRESS(ROW(),COLUMN())))>0")
    .setAllowInvalid(false)
    .setHelpText("This cell cannot be empty.")
    .build();

  sheet.getRange(range).setDataValidation(rule);
};

const validateEmail = (sheet: Sheet, range: string): void => {
  const emailRule = SpreadsheetApp.newDataValidation()
    .requireTextIsEmail()
    .setAllowInvalid(false)
    .setHelpText("This cell must be an email.")
    .build();

  sheet.getRange(range).offset(1, 0).setDataValidation(emailRule);
};

const validatePlaceholderFormat = (sheet: Sheet, range: string): void => {
  const formula =
    '=AND(LEFT(INDIRECT(ADDRESS(ROW(),COLUMN())),1)="{" , RIGHT(INDIRECT(ADDRESS(ROW(),COLUMN())),1)="}")';

  const placeholderRule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied(formula)
    .setAllowInvalid(false)
    .setHelpText('Placeholder name must be wrapped in "{ and "}.')
    .build();

  sheet.getRange(range).setDataValidation(placeholderRule);
};

const validateConfig = (ss: Spreadsheet): void => {
  if (
    !ss
      .getSheets()
      .map((sheet: Readonly<GoogleAppsScript.Spreadsheet.Sheet>) =>
        sheet.getName()
      )
      .includes(CONFIG_SHEET_NAME)
  ) {
    throw new Error(`No sheet called ${CONFIG_SHEET_NAME} found!`);
  }

  validateCellEq(
    ss,
    RANGE.header.templateSubject,
    HEADER_CELL_VALUE.templateSubject
  );
  validateCellEq(ss, RANGE.header.subject, HEADER_CELL_VALUE.subject);
  validateCellNotEmpty(ss, RANGE.templateSubject);
  validateCellNotEmpty(ss, RANGE.subject);
};

const validate = (ss: Spreadsheet): void => {
  validateConfig(ss);

  for (const sheet of ss.getSheets()) {
    if (sheet.getName() === CONFIG_SHEET_NAME) {
      continue;
    }

    validatePlaceholderFormat(sheet, RANGE.header.placeholder);
    validateCellEq(sheet, RANGE.header.email, HEADER_CELL_VALUE.email);
    validateEmail(sheet, RANGE.email);
    validateCellNotEmpty(sheet, RANGE.placeholder);
  }
};
