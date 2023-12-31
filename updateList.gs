/**
 * Updates a list by appending new entries from a source sheet to a destination sheet.
 * This can be used for various data types like emails, names, student IDs, etc.
 *
 * @param {string} sourceSheetId The ID of the source spreadsheet.
 * @param {string} sourceSheetName The name of the source sheet.
 * @param {string} sourceRange A1 notation for the range in the source sheet.
 * @param {string} destSheetName The name of the destination sheet in the active spreadsheet.
 * @param {string} destRange A1 notation for the range in the destination sheet.
 */
function updateList(sourceSheetId, sourceSheetName, sourceRange, destSheetName, destRange) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const destinationSheet = ss.getSheetByName(destSheetName);
  const emailSource = SpreadsheetApp.openById(sourceSheetId)
                                    .getSheetByName(sourceSheetName)
                                    .getRange(sourceRange)
                                    .getValues();

  let newData = [], rows = 0;
  let oldData = destinationSheet.getRange(destRange).getValues().toString();

  let sourceData = emailSource.map(function(el) { return [el]; });

  for (let i = 0; i < sourceData.length; i++) {
    if (oldData.indexOf(sourceData[i]) == -1) {
      newData[rows] = sourceData[i];
      rows++;
    }
  }

  if (rows == 0) {
    SpreadsheetApp.getActive().toast('No new entries found', 'Status', 1);
    Utilities.sleep(1000);
    return;
  }

  let lastFilledRow = getLastFilledRow(destinationSheet, destRange);

  if (newData.length == 0) { return; }

  let range = destinationSheet.getRange(lastFilledRow + 1, 1, rows, 1);
  range.setValues(newData);

  let message = rows == 1 ? '1 entry added' : rows + ' entries added';
  SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Status', 1);
  Utilities.sleep(1000);
}

/**
 * Finds the last filled row in a given column range.
 *
 * @param {Sheet} sheet The sheet to search in.
 * @param {string} columnRange A1 notation for the column range.
 * @return {number} The last filled row number.
 */
function getLastFilledRow(sheet, columnRange) {
  var column = sheet.getRange(columnRange);
  var values = column.getValues();
  var row = values.length - 1;
  while (row >= 0 && values[row][0] == '') {
    row--;
  }
  return row + 1;
}
