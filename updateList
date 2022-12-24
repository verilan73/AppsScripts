/************************************************************************************************************
 * This script updates the list of staff emails by appending them to the end of the list.
 * This ensures each row has a unique reference, ensuring manually added data remains tied to the staff member
 * **********************************************************************************************************/
function updateStaffList(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const destinationSheet = ss.getSheetByName("SheetName");
  const emailSource = SpreadsheetApp.openById("SheetID")
                                    .getSheetByName("Employed Staff")
                                    .getRange("B2:B")
                                    .getValues();
  
  var newData=[],rows=0;
  let oldData = destinationSheet.getRange("A4:A").getValues().toString();
  
  var sourceData = emailSource
      .map(function(el){
        return [el];
      });
  for (var i=0;i<sourceData.length;i++){
    if(oldData.indexOf(sourceData[i])== -1){
      newData[rows]=sourceData[i];
      rows++;
    }
  }
  
  if(rows==0){
    SpreadsheetApp.getActive().toast('No new staff','Status',1);
    Utilities.sleep(1000);
    return;
  }

  // get lastFilledRow in the specified column
  var column = destinationSheet.getRange('A:A');
  var value = ''
  const max = destinationSheet.getMaxRows();
  var values = column.getValues();
  values = [].concat.apply([], values);
  for (row = max - 1; row > 0; row--) {
    value = values[row];
    if (value != '') { break }
  }
  var lastFilledRow = row + 1;

  //Add in the list of new emails, appending them to the end of the previous set
  var range = destinationSheet.getRange(lastFilledRow+1,1,rows,1);
  if(newData.length==0){return;}

  range.setValues(newData);

  if(rows==1){
    SpreadsheetApp.getActiveSpreadsheet().toast('1 staff member added', 'Status',1);
    Utilities.sleep(1000);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast(rows + ' staff members added', 'Status',1);
    Utilities.sleep(1000);
  }
  
}
