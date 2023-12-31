// Script for managing non-teaching staff attendance and calendar events.
// It compares attendance data from a Google Sheet with events in a Google Calendar,
// creates new calendar events for new absences, and updates the sheet with these events.

const CALENDAR_ID = ""; // Define the Google Calendar ID

function updateCalendar() {
  // Access the Google Sheets for source data and event tracking
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Non-teaching Staff Attendance");
  const eventsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CalendaredEvents");
  
  // Get access to the specified Google Calendar
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);

  // Retrieve and process existing event data from the "CalendaredEvents" sheet
  let originalData = eventsSheet.getDataRange()
                                .getValues()
                                .map(row => row.slice(0, 3)) // Keep only the first 3 columns
                                .map(row => row.concat(row.join(''))); // Concatenate these columns for comparison

  // Retrieve and modify data from the "Non-teaching Staff Attendance" sheet
  let sourceData = sourceSheet.getDataRange().getValues();
  sourceData.shift(); // Remove the first two header rows
  sourceData.shift();
  let processedSourceData = sourceData.map(row => row.slice(0, 2).concat(row.slice(11))); // Remove unnecessary columns

  // Filter to select rows with non-empty values in specific columns
  let newAbsences = [];
  for (let i = 1; i < processedSourceData.length; i++) {
    let row = processedSourceData[i];
    for (let j = 2; j < row.length; j++) {
      let element = row[j];
      if (element !== "") {
        // Construct and add a new row to the newAbsences array
        newAbsences.push([row[0], row[1], processedSourceData[0][j], element]);
      }
    }
  }

  // Identify new absences to add to the calendar by comparing with original data
  var absencesToAdd = newAbsences.map(row => row.concat([row[0], row[2], row[3]].join('')))
                                 .filter(row => !originalData.some(existingRow => row[4] === existingRow[3]))
                                 .map(row => row.slice(0, 4));

  // Create new calendar events for each identified absence
  let eventsArray = [];
  for (let i = 0; i < absencesToAdd.length; i++) {
    let [email, name, date, absenceType] = absencesToAdd[i];
    let eventTitle = (typeof absenceType === "number") ? `${name} (VL ${absenceType} day)` : `${name} (Sick Leave)`;
    let event = cal.createAllDayEvent(eventTitle, new Date(date));
    event.setColor("5"); // Set the event color
    eventsArray.push([email, date, absenceType, eventTitle, event.getId()]);
  }

  // Append new event details to the "CalendaredEvents" sheet
  eventsSheet.getRange(eventsSheet.getLastRow() + 1, 1, eventsArray.length, 5).setValues(eventsArray);

  // Notify the user of the operation's completion
  let plural = eventsArray.length > 1 ? 's' : '';
  SpreadsheetApp.getActiveSpreadsheet().toast(`${eventsArray.length} event${plural} added to calendar`, 'Status', 2);
}
