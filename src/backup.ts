type Properties = {
  [key: string]: string;
  SOURCE_SPREADSHEET_ID: string;
  SOURCE_SHEET_NAME: string;
  TARGET_SPREADSHEET_ID: string;
  TARGET_SHEET_NAME: string;
};

export function execPreviousMonthDataBackup() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const properties = scriptProperties.getProperties() as Properties;

  const now = new Date();
  const previousMonthUnixTime = new Date(
    now.getFullYear(),
    now.getMonth() - 1
  ).getTime();

  const sourceSpreadsheet = SpreadsheetApp.openById(
    properties.SOURCE_SPREADSHEET_ID
  );
  const sourceSheet = sourceSpreadsheet.getSheetByName(
    properties.SOURCE_SHEET_NAME
  );

  if (sourceSheet === null) throw new Error("sourceSheet is not found");

  const sourceRange = sourceSheet.getDataRange();
  const sourceValues = sourceRange.getValues();

  const previousMonthValues = sourceValues.filter(row => {
    const rowUnixTime = new Date(row[12]).getTime();
    return previousMonthUnixTime === rowUnixTime;
  });

  const targetSpreadsheet = SpreadsheetApp.openById(
    properties.TARGET_SPREADSHEET_ID
  );
  const targetSheet = targetSpreadsheet.getSheetByName(
    properties.TARGET_SHEET_NAME
  );

  if (targetSheet === null) throw new Error("targetSheet is not found");

  const targetRange = targetSheet.getDataRange();

  targetSheet
    .getRange(
      targetRange.getLastRow() + 1,
      1,
      previousMonthValues.length,
      previousMonthValues[0].length
    )
    .setValues(previousMonthValues);
}
