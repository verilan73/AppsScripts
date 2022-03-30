/* 
Need to copy an entire folder from one Google Drive account to another Google Drive account? 
1. Right-click on original folder in Google Drive
2. Share with the destination Google account
3. Go into destination account's Google Drive
4. Find the shared folder under "Shared with me"
5. Select all the files (Ctrl-A / Cmd-A)
6. Right-click, Make a copy
7. Create a New Folder in destination account's Google Drive
8. Go to Recent
9. Select all recently copied files and move to the new folder
10. Great, except all the filenames begin with "Copy of"xxxxxx.txt
*/

function fileRename() {
  var ui = DocumentApp.getUi();
  
  var folderName = ui.prompt("FolderName?");
 if(folderName.getSelectedButton() == ui.Button.OK){
  var folders = DriveApp.getFoldersByName(folderName.getResponseText());
  var folder = folders.next();
  var files = folder.getFiles();
  
  while(files.hasNext()){
    var file = files.next()
    var fileName = file.getName();
    if (fileName.indexOf('Copy of ') > -1) {
        fileName= fileName.split('Copy of ')[1];
        file.setName(fileName);
    };
  };

}
}
function onOpen(){
var ui = DocumentApp.getUi();
  ui.createMenu('Scripts')
    .addItem('Remove Copy of prefix','fileRename')
    .addToUi();
}
