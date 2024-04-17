/***********************************************************************************************************************
 * Add this code to your spreadsheet file, connecting it to your script library where you have the Transcripts.gs code
 * You will also want the indexPage.gs script
 ***********************************************************************************************************************/

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Script')
      .addItem('UpdateIndex', 'indexPage')
      .addItem('Update a single student','updateSingleStudent')
      .addItem('Update transcripts from Source Data','processAllStudents')
      .addToUi();

   YourLibraryName.updateIndexSheet("C3","GREEN")   
}

function indexPage() {
  YourLibraryName.updateIndexSheet("C3","GREEN")
}

function updateSingleStudent(){
  YourLibraryName.updateSingleStudent();
}

function processAllStudents(){
  YourLibraryName.processAllStudents();
}



