 function saveSubFolderLinksToExcel() {
  var folderId = '1JxieFM8ZB58POLoO9WTFPNJ6xHqGBGqT'; // Replace with the ID of the folder containing the subfolders
  var folder = DriveApp.getFolderById(folderId);
  var subfolders = folder.getFolders();
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = 1;
  
  while (subfolders.hasNext()) {
    var subfolder = subfolders.next();
    sheet.getRange(row, 1).setValue(subfolder.getName());
    sheet.getRange(row, 2).setValue(subfolder.getUrl());
    row++;
  }
}
