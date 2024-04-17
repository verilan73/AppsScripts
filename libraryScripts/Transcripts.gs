/** This collection of functions populates individual transcript sheets with the relevant details for each student.
 *  Note that the script cannot handle when a student leaves after Quarter reports but before the end of the semester. In this case,
 *  please review the student's individual profile in ManageBac and update the transcript manually with the relevant Quarter data
 *  if so required.
 * 
 *  Note that we're writing to specific cell locations, so do NOT change the template without also updating the script accordingly.
 * 
 *  With ENORMOUS thanks to ChatGPT 4 who generated the code for this.
 * 
 *  I am considering whether it would make sense to update my BQ query to combine the demographic & transcript data so it can be pulled
 *  via a single request.
 * 
 * */ 

/*****************************************************************************************************************************************
 *  This first script works through a list of students and generates or updates transcripts, depending on whether they already existed. We 
 *  also remind the user to run the script to update the index page. However, note that this should also run automatically every time the
 *  file is opened, regardless.
 ******************************************************************************************************************************************/
function processAllStudents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("Source Data"); // Adjust the name as per your setup
  var dataRange = sourceSheet.getDataRange();
  var data = dataRange.getValues();
  
  data.shift(); // Remove header row
  
  data.forEach(function(row) {
    var studentId = row[0];
    var classOf = row[1];
    var sheetNamePrefix = 'C' + classOf + '-' + studentId;
    var sheet = findOrCreateSheet(sheetNamePrefix, studentId);
    
    importStudentData(studentId, sheetNamePrefix, classOf); // Call modified import function
    
    // Update the "Source Data" sheet with status
    var statusColumnIndex = 3; // Assuming status is in the third column
    sourceSheet.getRange(dataRange.getRowIndex()+1 + data.indexOf(row), statusColumnIndex).setValue(sheet ? "Updating previous version" : "Creating new transcript");
  });
  
  SpreadsheetApp.getActiveSpreadsheet().toast('Data update complete. Please remember to update the index page.', 'Update Complete', 10);
}
/***********************************************************************************************
 * This is the master function that places data in the appropriate locations in each transcript.
 ***********************************************************************************************/
function importStudentData(studentId, sheetName, classOf) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error("Sheet not found: " + sheetName);
  }

  //Clear any old data that may be in place
  sheet.getRange("B14:H27").clearContent();
  sheet.getRange("B30:H50").clearContent();

  // Set the graduation year in cell C9
  sheet.getRange('C9').setValue(classOf);

  // Fetch and set the student's personal information
  fetchAndSetStudentInfo(studentId, sheet);

  //Calculate academic years based on "Class of" year, then place them appropriately in the transcript sheet.
  var academicYears = [
    (classOf - 1) +" - " + classOf + " Academic Year",
    (classOf - 2) +" - " + (classOf - 1) + " Academic Year",
    (classOf - 3) +" - " + (classOf - 2) + " Academic Year",
    (classOf - 4) +" - " + (classOf - 3) + " Academic Year"
  ];
  sheet.getRange("C12").setValue(academicYears[0]);
  sheet.getRange("G12").setValue(academicYears[1]);
  sheet.getRange("C28").setValue(academicYears[2] );
  sheet.getRange("G28").setValue(academicYears[3]);
  
  // By also outputting the column headers, we're able to adjust whether the grades are reported as 
  // "Final Grades" or "Semester 1", depending on whether we have whole year data, or just data from
  // the first semester.
  //  These cells are within the Header row for each academic year.
  var outputCells = ['B14', 'F14', 'B30', 'F30'];
  
  academicYears.forEach(function(year, index) {
    var data = queryBigQueryData(studentId, year, sheetName); // Query BigQuery for each academic year
    if (data.length > 0) {
      sheet.getRange(outputCells[index]).offset(0, 0, data.length, data[0].length).setValues(data); // Paste the array into the sheet
    }

  });
}
/****************************************************************************************************************
 *  We check if a particular transcript sheet already exists. If yes, we will update it, otherwise we create one.
 ****************************************************************************************************************/
function findOrCreateSheet(sheetNamePrefix, studentId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = ss.getSheets();
  var sheet = allSheets.find(function(s) { return s.getName().indexOf(sheetNamePrefix) === 0; });
  
  if (!sheet) {
    // Copy from MASTER and rename
    var masterSheet = ss.getSheetByName("MASTER");
    sheet = masterSheet.copyTo(ss).setName(sheetNamePrefix); // This is the filename, originally the plan was to add the student name, but no longer.
  }
  return sheet;
}

/******************************************************
 * * Here we pull basic demographic data from BigQuery. 
 ******************************************************/
function fetchAndSetStudentInfo(studentId, sheet) {
  var ui = SpreadsheetApp.getUi(); 
  var projectId = 'PROJECT ID'; //Big Query project ID
  var request = {
    query: 'SELECT CONCAT(first_name, " ", last_name) AS name, birthday, gender, attendance_start_date, graduated_on, withdrawn_on ' +
           'FROM `BIGQUERY TABLE ID` WHERE student_id = "' + studentId + '"',
    useLegacySql: false
  };
  
  try{
    var queryResults = BigQuery.Jobs.query(request, projectId);
    var jobId = queryResults.jobReference.jobId;
  
    // Check on status of the Query Job
    var sleepTimeMs = 500;
    while (!queryResults.jobComplete) {
      Utilities.sleep(sleepTimeMs);
      sleepTimeMs *= 2;
      queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId);
    }
  
  // Assuming there's only one row returned
  if (queryResults.rows && queryResults.rows.length > 0) {
    var row = queryResults.rows[0].f;

    sheet.getRange('C8').setValue(row[0].v); // Student's full name
    sheet.getRange("A1").setNote(row[0].v);
    sheet.getRange('C10').setValue(row[3].v); // Attendance start date

    sheet.getRange('G8').setValue(row[2].v); // Gender
    sheet.getRange('G9').setValue(row[1].v); // Birthday
    // Conditional logic for Graduated or Withdrawn
    if (row[4].v !== null) { // If graduated_on is not null
      sheet.getRange('F10').setValue('Graduation Date:');
      sheet.getRange('G10').setValue(row[4].v); // Graduated on
      //Add in the details at the end of the file and some formatting
        sheet.getRange("B57").setValue("Diploma Achieved");
        sheet.getRange("D57").setValue("International Baccalaureate");
        sheet.getRange("F57").setValue("YOUR SCHOOL NAME");
        sheet.getRange("G57").setValue("Korean High School");
        sheet.getRange("B59").setValue("Date");
        sheet.getRange("D59").setFormula('=date(C9,7,1)');
        sheet.getRange("F59").setFormula('=edate(D59,-2)');
        sheet.getRange("G59").setFormula('=F59');
        sheet.getRange("B57:G57").setBorder(null, null, true, null, null, null, '#00669f',SpreadsheetApp.BorderStyle.SOLID);

    } else if (row[5].v !== null) { // If withdrawn_on is not null (and graduated_on is null)
      sheet.getRange('F10').setValue('Withdrawal Date:');
      sheet.getRange('G10').setValue(row[5].v); // Withdrawn on
    } else { // If both are null
      sheet.getRange('F10').setValue('Withdrawal Date:');
      sheet.getRange('G10').setValue(''); // Clear any existing value
    }
  } else {
    ui.alert('No data found for the provided Student ID.');
  } 
  } catch (error) {
    //Display a toast message with the error
    SpreadsheetApp.getActiveSpreadsheet().toast('Failed to fetch student information. Please check the logs for more details.', 'Query Failed', -1);
    // Optionally log the error for more detailed debugging
    Logger.log('Error fetching student information: ' + error.toString());
    }
  }

/*******************************************************************************************************************************************
 * Here is where we pull the data from BiqQuery. If any results come through incorrectly, go into BigQuery, review the data, and review the
 *  query 'Transcripts - 2 Create Transcript Details'
 *******************************************************************************************************************************************/
function queryBigQueryData(studentId, academicYear, sheetName) {
  var projectId = 'BIG QUERY PROJECT ID'; // Replace with your project ID
  var request = {
    query: 'SELECT class_name, term_grade, credits, assessment_period FROM `BIG QUERY TABLE ID` ' +
           'WHERE mb_id = "' + studentId + '" AND academic_year = "' + academicYear + '"',
    useLegacySql: false
  };
  
  var queryResults = BigQuery.Jobs.query(request, projectId);
  var jobId = queryResults.jobReference.jobId;
  
  // Check on status of the Query Job
  var sleepTimeMs = 500;
  while (!queryResults.jobComplete) {
    Utilities.sleep(sleepTimeMs);
    sleepTimeMs *= 2;
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId);
  }
  
  // Initialize data with a default header
  var data = [['Class', 'Final Grade', 'Credits']];
  
  if (queryResults.rows) {
    var allSemester1 = true;
    var allHalfCredits = true;
    
    var rowData = queryResults.rows.map(function(row) {
      // Check if any entry is not from Semester 1
      if (row.f[3].v !== 'Semester 1') {
        allSemester1 = false;
      }
      
      // Check if any credits are not 0.5
      if (parseFloat(row.f[2].v) !== 0.5) {
        allHalfCredits = false;
      }
      
      return [row.f[0].v, row.f[1].v, row.f[2].v];
    });
    
    // If all data meets the criteria, adjust the header accordingly
    if (allSemester1 || allHalfCredits) {
      data[0][1] = 'Semester 1';
    }
    
    // Append the rows to the data array
    data = data.concat(rowData);
  }
  
  return data;
}

/*******************************************************************************************************************************************
 * This enables us to create, or update, the transcript for just one student at a time. We can also do this using the Source Data sheet, but 
 * we might not want to mess with what is there.
 *******************************************************************************************************************************************/
function updateSingleStudent() {
  var ui = SpreadsheetApp.getUi(); // Get the UI environment to use alerts and prompts.
  
  // Prompt for Student ID
  var studentIdResponse = ui.prompt('Update Student Record', 'Please enter the Student ID:', ui.ButtonSet.OK_CANCEL);
  
  // Check if the user clicked "OK"
  if (studentIdResponse.getSelectedButton() == ui.Button.OK) {
    var studentId = studentIdResponse.getResponseText().trim();
    
    // Prompt for Class Of
    var classOfYearResponse = ui.prompt('Update Student Record', 'Please enter the "Class Of" Year:', ui.ButtonSet.OK_CANCEL);
    
    if (classOfYearResponse.getSelectedButton() == ui.Button.OK) {
      var classOfYear = classOfYearResponse.getResponseText().trim();

      //We've got what we need, now let's get the sheet and the data and create the transcript
      var sheetName = 'C' + classOfYear + '-' + studentId;
      var sheet = findOrCreateSheet(sheetName, studentId);
      importStudentData(studentId, sheet.getName(), classOfYear);
    }
  }

}
