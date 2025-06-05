/* exported Sheet, Spreadsheet */

interface Sheet {
  readonly getRange: (range: string) => GoogleAppsScript.Spreadsheet.Range;
}

type Spreadsheet = Readonly<GoogleAppsScript.Spreadsheet.Spreadsheet>;
