/*
** function that creates new menu item in top ribbon in Google Sheets to run the other functions
*/

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('NEW MONTH')
    .addItem('Duplicate Calendar', 'makeDuplicate')
    .addItem('Download as PDF', 'exportSheetAsPDF')
    .addItem('Clear Values', 'resetBoxes')
    .addToUi();
}

/**
 * function to create the duplicated calendar spreadsheet and send to appropriate folder
 */
 
function makeDuplicate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetName = ss.getName();
  //Logger.log(spreadsheetName);
  
  var dupeCalendar = DriveApp.getFilesByName(spreadsheetName);
  Logger.log(dupeCalendar.hasNext());
  
  var destFolder = DriveApp.getFoldersByName('backup calendars').next();
  DriveApp.getFileById(ss.getId()).makeCopy('Backup: ' + spreadsheetName, destFolder);
}
  
/**
 * function to download specific range of the spreadsheet to send to accounts
 */

function exportSheetAsPDF() {
  var ogSS = SpreadsheetApp.getActive(); // grab original spreadsheet
  var sheet = ogSS.getActiveSheet();
  
  var sourceRange = sheet.getRange('A1:I36').activate(); // this is where i select columns i want to export from sheet
  var sourceValues = sourceRange.getValues();
  Logger.log(sourceValues);
  
  var sheetName = sheet.getName() + ' Content Calendar';
  var folder = DriveApp.getFoldersByName('2019').next();
  
  // temporary spreadsheet to put range in for download
  var destSS = SpreadsheetApp.open(DriveApp.getFileById(ogSS.getId()).makeCopy('tmp_pdf_ss'));
  var destSheet = destSS.getSheets()[0];
  
  var destRange = destSheet.getRange('A1:I36');
  destRange.setValues(sourceValues);
  Logger.log(destRange);
  
  var blob = destSS.getBlob().getAs('application/pdf').setName(sheetName);
  var newFile = folder.createFile(blob);
  
  DriveApp.getFileById(destSS.getId()).setTrashed(true);
}

/**
 * function to clear checked boxes and posts after duplicate is made
 */
 
function resetBoxes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  // cell ranges containing info on FB/IG/TW posts
  var dataRange = sheet.getRangeList(['A12:G14','A16:G18','A20:G22','A24:G26','A28:G30', 'A32:G34']);
  dataRange.activate();
  dataRange.clearContent();
  
  // cell ranges with checked boxes
  var boxRange = sheet.getRange(12, 10, 23, 1);
  boxRange.activate();
  boxRange.clearContent();
}
