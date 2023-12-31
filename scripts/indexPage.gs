/**This will generate a nice index page, including a descriptions column. The descriptions come from
 * Notes which are added to cell A1 on each tab.
 * 
 * The script is designed to enable the user to decide exactly where on their Index page they want the
 * table to start, and also to choose their own colour theme.
 * 
 * With thanks to ChatGPT for building the basic function which I then expanded on. Intended to be run
 * from a scripts library so it is easy to use across various files. When paired with an onOpen() that
 * also calls this script, it means the index page remains up to date.
 */

function updateIndexSheet(startCell, bandingTheme) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  var indexSheet = ss.getSheetByName("Index");
  if (!indexSheet) {
    indexSheet = ss.insertSheet("Index", 0);
  }

  // Add in some instructions for the user which will appear on the generated index page.
  indexSheet.clear();
  indexSheet.getRange("A1").setNote("Description/Instructions can be added to each sheet by adding a Note to cell A1.\nTo update this Index, go to the Extensions Menu, then AppsScript, then run the OnOpen Script and allow all the permissions.\nThis must be done by each editor of this file and will ensure the Index is updated each time the file is opened.")

  // Pull in the list of sheets and their associated notes.
  var data = sheets.map(function(sheet) {
    var sheetName = sheet.getName();
    var note = sheet.getRange('A1').getNote(); // Get the note from cell A1
    var link = '=HYPERLINK("#gid=' + sheet.getSheetId() + '", "' + sheetName + '")';
    return [link, note]; // Include the hyperlink and note in the data array
  });

  // Parse the starting cell coordinates
  var startCoords = parseCellCoordinates(startCell);
  var headerRow = startCoords.row;
  var dataStartRow = headerRow + 1;
  indexSheet.setFrozenRows(headerRow);

  // Set header and apply formatting
  var headerRange = indexSheet.getRange(headerRow, startCoords.column, 1, 2);
  headerRange.setValues([["Sheet Name", "Description/Instructions"]])
              .setFontWeight("bold");

  // Set data
  var range = indexSheet.getRange(dataStartRow, startCoords.column, data.length, 2);
  range.setValues(data);

  // Apply banding
  applyBanding(indexSheet, headerRow, data.length, startCoords.column, bandingTheme);

}

// Helper function to parse cell coordinates
function parseCellCoordinates(cellRef) {
  var match = cellRef.match(/^([A-Z]+)(\d+)$/);
  return {
    column: columnToNumber(match[1]),
    row: parseInt(match[2])
  };
}

// Helper function to convert column letter to number
function columnToNumber(col) {
  var column = 0, length = col.length;
  for (var i = 0; i < length; i++) {
    column += (col.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

// Helper function to apply banding & apply some nice formatting
function applyBanding(sheet, startRow, numRows, startColumn, theme) {
  // Remove existing bandings in the range to avoid overlaps
  var bandings = sheet.getBandings();
  bandings.forEach(function(banding) {
    banding.remove();
  });

  var bandingRange = sheet.getRange(startRow, startColumn, numRows + 1, 2);
  bandingRange.applyRowBanding(SpreadsheetApp.BandingTheme[theme], true, false);

  // Set up some nice formatting
  bandingRange.setVerticalAlignment("middle").setWrap(true);

}
