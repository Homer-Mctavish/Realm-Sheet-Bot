function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('ouoohhthebutton')
    .setTitle('Installable Trigger-O-Matic');
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showSidebar(html);
    authorizeItemImport();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Realm Custom Scripts')
    .addItem('Add Update Trigger', 'showSidebar')
    .addToUi();
}


function findRow(searchVal) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let data = sheet.getDataRange().getValues();
  let columnCount = sheet.getDataRange().getLastColumn();

  let i = data.flat().indexOf(searchVal);
  let columnIndex = i % columnCount;
  let rowIndex = ((i - columnIndex) / columnCount);

  Logger.log({columnIndex, rowIndex }); // zero based row and column indexes of searchVal

  return i >= 0 ? rowIndex + 1 : "searchVal not found";
}

function updeeat(e) {
  //let rangp = findRow(e.value);
  SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1FSyukC97LQ8MCEvbTlrZojJF-UiFdQaZM7PGr5Ky9dQ/edit#gid=1147878197").getSheetByName("Copy of Item Import").getRange("F"+findRow(e.value)).setValue(e.value);
}

function createSpreadsheetEditTrigger() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('updeeat')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
}