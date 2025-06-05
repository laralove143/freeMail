/* exported processSheet */

interface Placeholder {
  readonly name: string;
  readonly values: readonly string[];
}

const getEmails = (
  sheet: Readonly<GoogleAppsScript.Spreadsheet.Sheet>
): string[] => {
  const emails = sheet.getRange(RANGE.email).getValues().flat();

  if (!emails.every((email) => typeof email === "string")) {
    throw new Error("Failed to get emails.");
  }

  return emails;
};

const getPlaceholders = (
  sheet: Readonly<GoogleAppsScript.Spreadsheet.Sheet>
): Placeholder[] => {
  const err = new Error("Failed to get placeholders.");

  const is2DArrayString = (
    value: readonly (readonly unknown[])[]
  ): value is string[][] =>
    value.every((array) => array.every((item) => typeof item === "string"));

  const placeholders = sheet.getRange(RANGE.placeholder).getValues();

  if (!is2DArrayString(placeholders)) {
    throw err;
  }

  const [placeholderNames] = placeholders;

  if (!placeholderNames) {
    throw err;
  }

  const placeholdersMapped = placeholderNames.map((name, idx) => ({
    name,
    values: placeholders.slice(1).map((row: readonly string[]) => {
      const value = row[idx];

      if (typeof value !== "string") {
        throw new Error("Failed to get placeholder values.");
      }

      return value;
    }),
  }));

  return placeholdersMapped;
};

const processRow = (params: {
  readonly rowIdx: number;
  readonly template: string;
  readonly subject: string;
  readonly emails: readonly string[];
  readonly placeholders: readonly Placeholder[];
}): void => {
  const email = params.emails[params.rowIdx];

  if (typeof email !== "string") {
    throw new Error("Failed to get email.");
  }

  let body = params.template;
  for (const placeholder of params.placeholders) {
    const placeholderValue = placeholder.values[params.rowIdx];

    if (typeof placeholderValue !== "string") {
      throw new Error("Failed to get placeholder value.");
    }

    body = body.replaceAll(placeholder.name, placeholderValue);
  }

  GmailApp.createDraft(email, params.subject, "", { htmlBody: body });
};

const processSheet = (
  sheet: Readonly<GoogleAppsScript.Spreadsheet.Sheet>,
  template: string,
  subject: string
): void => {
  const emails = getEmails(sheet);
  const placeholders = getPlaceholders(sheet);

  for (let rowIdx = 1; rowIdx < emails.length; rowIdx += 1) {
    processRow({
      emails,
      placeholders,
      rowIdx,
      subject,
      template,
    });
  }
};
