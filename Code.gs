function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('dataform')
    .setTitle('Estimator Robot');
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showSidebar(html);
    authorizeItemImport();
}

function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .createMenu('Realm Custom Scripts')
    // .addSubMenu(SpreadsheetApp.getUi().createMenu('Add New Connection').addItem('Mysql', 'createMysqlPrompt').addItem('SQL Server','createMssqlPrompt'))
    .addItem('Show Estimator sidebar', 'showSidebar')
    // .addItem('Refresh', 'refreshPrompt')
    .addToUi();
}

function authorizeItemImport(){

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Item Import");
//authroize the data import cell
  var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
      addImportrangePermission_(ssId,'1xz9Y9EgLcui3ekKkLic-3BC3Z8RS1s4qWvz5NFu6EM4');

//SpreadsheetApp.getUi().alert("running import");

  sheet.getRange(1, 1).setFormula("");
  sheet.getRange(1, 1).setFormula("=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1xz9Y9EgLcui3ekKkLic-3BC3Z8RS1s4qWvz5NFu6EM4/edit#gid=0\",\"items!A1:D\")");

}

function addImportrangePermission_(fileId, donorId) {
  // adding permission by fetching this url
  var url = 'https://docs.google.com/spreadsheets/d/' +
    fileId +
    '/externaldata/addimportrangepermissions?donorDocId=' +
    donorId;
  var token = ScriptApp.getOAuthToken();
  var params = {
    method: 'post',
    headers: {
      Authorization: 'Bearer ' + token,
    },
    muteHttpExceptions: true
  };
  UrlFetchApp.fetch(url, params);
}

//added by SS
function doGet(e) {
  if(!e.parameter.page){
        return HtmlService.createTemplateFromFile('itemform.html')
        .evaluate() // evaluate MUST come before setting the Sandbox mode
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
  } 
  return HtmlService.createTemplateFromFile('sqlconn.html').evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function runRealmItemAdd() {
      //this will create and show our html file
   var t = HtmlService.createTemplateFromFile('itemform');
    var htmlOutput = t.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setHeight(600).setWidth(900);
    var doc = SpreadsheetApp.getActive();
    doc.show(htmlOutput);
  
}


/*
function onEdit(event) {
  var sheet = SpreadsheetApp.getActiveSheet();

  if( sheet.getName() == "Internal" ) { //checks that we're on the correct sheet
    var r = event.range;
    if( r.getColumn() == 3 ) { //checks the column
  var activeRow=  r.getRow();
  var itemcodeSelected = sheet.getRange(r.getRow(), 3).getValue();
  r.offset(0, 2).setFormula("=IF(C" + activeRow + " = \"\",\"\",VLOOKUP(C" + activeRow + ",'Item Import'!$A$2:D,2,0))");
 //     r.offset(0, 2).setFormula("=IF(C" + activeRow + " = \"\",\"\",QUERY('Item Import'!A2:D,\"SELECT B WHERE A = ''\"; 0))");
      r.offset(0, 3).setFormula("=ROUND(L" + activeRow + ",-1)");
      r.offset(0, 4).setFormula("=IF(C" + activeRow + "=\"\",\"\",SUMIF(VLOOKUP(Internal!C" + activeRow + ",'Item Import'!$A$2:D,3,0),\"<>#N/A\"))");
      r.offset(0, 5).setFormula("=G" + activeRow + "*D" + activeRow + "");
      r.offset(0, 6).setFormula("=IF(C" + activeRow + "=\"\",\"\",SUMIF(VLOOKUP(Internal!C" + activeRow + ",'Item Import'!$A$2:D,4,0),\"<>#N/A\"))");  
      r.offset(0, 7).setFormula("=I" + activeRow + "*D" + activeRow + "");
      r.offset(0, 8).setFormula("=I" + activeRow + "*(1-'Project Calcs'!$C$10)"); 
      r.offset(0, 9).setFormula("=K" + activeRow + "*D" + activeRow + ""); 
      r.offset(0, 10).setFormula("=J" + activeRow + "*'Project Calcs'!$C$6"); 
      r.offset(0, 11).setFormula("=J" + activeRow + "*'Project Calcs'!$C$3");  
      r.offset(0, 12).setFormula("=N" + activeRow + "*'Project Calcs'!$C$4"); 
      r.offset(0, 13).setFormula("=N" + activeRow + "*'Project Calcs'!$C$5*(1-'Project Calcs'!$C$9)"); 
      r.offset(0, 14).setFormula("=L" + activeRow + "-H" + activeRow + ""); 
      r.offset(0, 15).setFormula("=P" + activeRow + "-O" + activeRow + "");
      r.offset(0, 16).setFormula("=M" + activeRow + "*'Project Calcs'!$C$7");  
      r.offset(0, 17).setFormula("=R" + activeRow + "+Q" + activeRow + "+S" + activeRow);  
      r.offset(0, 18).setFormula("=IF(Internal!B"+ activeRow +"=\"\",\"\",Internal!I"+ activeRow +"*'Project Calcs'!$C$8)");     
  }
}

if(sheet.getName() == "Hardware Ordering"){ 
  var r = event.range;
  var col =  r.getColumn();


  if( col == 3 ) { 
      var watchCol = [3], 
          userCol = [4],
          stampCol = [5],
          ind = watchCol.indexOf(event.range.columnStart);
          row = event.range.getRow();
      if ( ind == -1 || event.range.rowStart < 2) return;
      
      if(sheet.getRange(row,3).isChecked()==false){
                sheet.getRange(row, stampCol[ind]).setValue("");
                sheet.getRange(row, userCol[ind]).setValue("");
      } else {
                sheet.getRange(row, stampCol[ind]).setValue(event.value ? new Date() : null);
                sheet.getRange(row, userCol[ind]).setValue(Session.getEffectiveUser().getUsername());
          }

    }

  if( col == 6 ) { 
    
      var watchCol = [6], 
          userCol = [8],
          stampCol = [7],
          ind = watchCol.indexOf(event.range.columnStart);
          row = event.range.getRow();
      if ( ind == -1 || event.range.rowStart < 2) return;
      
      if(sheet.getRange(row,6).isChecked()==false){
                sheet.getRange(row, stampCol[ind]).setValue("");
               // sheet.getRange(row, userCol[ind]).setValue("");
      } else {
                sheet.getRange(row, stampCol[ind]).setValue(event.value ? new Date() : null);
               // sheet.getRange(row, userCol[ind]).setValue(Session.getEffectiveUser().getUsername());
          }

    }




}

}
*/
//added by SS
var ssApp = {
  activeSpreadSheet: SpreadsheetApp.getActiveSpreadsheet(),
  activeSheet: SpreadsheetApp.getActiveSheet(),
  namedSheet: function(name){ activeSpreadSheet.getSheetByName(name)},
  idSheet: function(id){activeSpreadSheet.openByID},
  setA1Val: function(range, value){
    activeSpreadSheet.getRange(range).setValue(value);
  }
}
const activeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
const activeSheet = SpreadsheetApp.getActiveSheet();
//todo: instantiate this variable when the proper spreadsheet is active
// var cellulor = activeSpreadSheet.getSelection().getActiveRangeList().getRanges();

function protection(rabge){
  var protection = rabge.protect();
  var me = Session.getEffectiveUser();
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}



//can be turned into an onEdit solution where rather than asynchronusly adding some cells via highlight, whatever added cells just have the range coppied to.
/**
 * sets highlighted number of rows as number to be added, preserves formulas contained within all of them. 
 * Only works (and only should work) when rows are highlighted entirely across and add rows before is chosen as the method of addition.
 */
function addRow(){
  var sheet = ssApp.activeSheet;
  var range = sheet.getActiveRange();
    try {
    let fill = sheet.getRange("2:2");
    SpreadsheetApp.flush();
    sheet.insertRowsBefore(sheet.getActiveCell().getRow(), range.getValues().length);
    SpreadsheetApp.flush();
    fill.copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  } catch (err){
    Logger.log('Failed with an error %s', + err.message)
  }
}


function getItemList() {
    var sheet = ssApp.activeSpreadSheet.getSheetByName("Item Import");
    try {
    let data = sheet.getDataRange().getValues()
    let array = [];
    data.forEach(function(row){array.push([row[0],row[1]]); });
   // Logger.log(array);
    return array;
  } catch (err){
	Logger.log('Failed with an error %s', + err.message)
  }
}
//end add

//edited by SS
/**
 * adds single item to the internal sheet. Active cell (and it can only be one cell) must be the Item Name column 
 * and in the row you wish to add the details into.
 * 
 * @param {String}  selectedItemToPaste:  the item you wish to insert by name
 * @param {String}  itemQty:  the number of item to add
 * @param {String}  itemRoom:   name of room item will be added to.
 */
function addItems(selectedItemToPaste,itemQty,itemRoom){ 
  let sheet = ssApp.activeSheet;
  let srow = sheet.getActiveRange().getRow();
  let scolumn = sheet.getActiveRange().getColumn();

  //change scolumn to letter, change s
  let scolumnlet1 = getLetter(scolumn-2);
  let scolumnlet2 = getLetter(scolumn+1);
  ssApp.activeSpreadSheet.getRange(scolumnlet1+srow).setValue(itemRoom.toUpperCase());
  SpreadsheetApp.getActiveRange().setValue(selectedItemToPaste);
  ssApp.activeSpreadSheet.getRange(scolumnlet2+srow).setValue(itemQty);
}


function getBOMList() {
  const ss = SpreadsheetApp.openById("1xz9Y9EgLcui3ekKkLic-3BC3Z8RS1s4qWvz5NFu6EM4"); 
  var roomTypeSheet = ss.getSheetByName("BOM");
  var getLastRow = roomTypeSheet.getLastRow();
  var data = roomTypeSheet.getRange(2, 1, getLastRow - 1, 2).getValues();
 // Logger.log(data);
  return data;
}
//end edit

function addBOMtoTemplate() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt("Please input BOM Type");
  var bomName = result.getResponseText();
 
  //make sure BOM Type doesn't already exist
  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bomSheet = ss.getSheetByName("BOM");
  var bomSheetLastRow = bomSheet.getLastRow()+1;
  var bomSheetValue = bomSheet.getRange("A2:" + "ZZ" + bomSheetLastRow).getValues();
  let ohNoUserBadInput = false;

  for (var i = 0; i < (bomSheetLastRow - 1); i++) {
    var bomType = bomSheetValue[i][0];
    if (bomName===""){
      ohNoUserBadInput = false;
    }
    else if (bomName === bomType ){
      ohNoUserBadInput = true;
    }else{
      ohNoUserBadInput=false;
    }

  }
  if (ohNoUserBadInput){
    ui.alert("BOM TYPE ALREADY EXISTS. Bye Felicia.");
    return
  } else {
    var range = '';
    var selectedValues = SpreadsheetApp.getActiveSheet().getActiveRange().getValues();
    var selectedValuesArrayCount = selectedValues.length;
    bomSheet.getRange(bomSheetLastRow, 1).setValue(bomName);

    var item = 3;
    var qty = 4;

    for (var i = 0; i < selectedValuesArrayCount; i ++){
      bomSheet.getRange(bomSheetLastRow, item).setValue(selectedValues[i][0]);
      bomSheet.getRange(bomSheetLastRow, qty).setValue(selectedValues[i][1]);
      item += 2;
      qty+=2;
    }
  }
}
//end add
function openDialog() {
  var html = HtmlService.createTemplateFromFile('dataform')
    .evaluate();

    SpreadsheetApp.getUi() 
    .showModalDialog(html, 'Dialog title');
}

function include(File) {
  return HtmlService.createHtmlOutputFromFile(File).getContent();
};

function insertItems(selectedRoomNameInput, selectedBomType) {
    const ss = SpreadsheetApp.openById("1xz9Y9EgLcui3ekKkLic-3BC3Z8RS1s4qWvz5NFu6EM4"); 
    var selectedRoomNames = [];
    selectedRoomNames = selectedRoomNameInput.split(",");
    var selectedBomType = selectedBomType;
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var internalSheet = sheet.getSheetByName("Internal");

    var internalLastRow =  getLastDataRow(internalSheet) +2;

  // if user wants to just copy current selected items we will loop through selected and copy paste that. 
  if (selectedBomType == "Selected Text") {
    var range = '';
    var selectedValues = SpreadsheetApp.getActiveSheet().getActiveRange().getValues();

    //should use this information to verify we are only highlighting columns C and D
    var sel = SpreadsheetApp.getActive().getSelection().getActiveRangeList().getRanges();
    for (var i = 0; i < sel.length; i++) {
      var rangeStart = sel[i].getA1Notation();
      range += sel[i].getA1Notation() + ', ';
    }


    selectedRoomNames.forEach(function (selectedRoomName) {
      if ((selectedRoomName.length > 0) && (selectedRoomName !== " ")) {
        var selectedValuesArrayCount = selectedValues.length;
        for (var i = 0; i < selectedValuesArrayCount; i ++){
        //insert Room name
        internalSheet.getRange(internalLastRow, 1).setValue(selectedRoomName.toUpperCase());
        // insert item number
        internalSheet.getRange(internalLastRow, 3).setValue(selectedValues[i][0]);
        //insert item qty
        internalSheet.getRange(internalLastRow, 4).setValue(selectedValues[i][1]);
        internalLastRow = internalLastRow + 1;
      //  SpreadsheetApp.getUi().alert(internalLastRow);
        }
      }
      
         internalLastRow = internalLastRow + 1;
       //  SpreadsheetApp.getUi().alert(internalLastRow);
    });
  }

    // designate necessary information to modify and read from sheet.
    var bomSheet = ss.getSheetByName("BOM");
    var bomSheetLastRow = getFirstEmptyBOMRowWholeRow(bomSheet);
    var bomSheetLastColum = bomSheet.getLastColumn() + 1;
    var bomSheetValue = bomSheet.getRange("A2:" + "BQ" + bomSheetLastRow).getValues();
    var insertData = "";

    selectedRoomNames.forEach(function (selectedRoomName) {
      if ((selectedRoomName.length > 0) && (selectedRoomName !== " ")) {

        //loop through all the rows of the BOM sheet.
        for (var i = 0; i < (bomSheetLastRow - 1); i++) {
          var bomType = bomSheetValue[i][0];
          if (selectedBomType == bomType) {
            // we assume that price is even and item number is odd
            for (var ii = 2; ii < (bomSheetLastColum - 1); ii++) {
              if (bomSheetValue[i][ii]) {
                //insert Room name
                internalSheet.getRange(internalLastRow, 1).setValue(selectedRoomName.toUpperCase());
                insertData = bomSheetValue[i][ii];
                // insert item number
                if (!isOdd(ii)) {
                  internalSheet.getRange(internalLastRow, 3).setValue(insertData);
                  //insert item qty
                } else {
                  internalSheet.getRange(internalLastRow, 4).setValue(insertData);
                  internalLastRow = internalLastRow + 1;
                }
              };
            };
            var bomType = "";
          };
        };
        internalLastRow = internalLastRow + 1;
      }
    });
  }

//added by SS
/**
 * inserts an item, discription cost and name to the named spreadsheet
 * 
 * @param {String}  item:  the item you wish to insert by name. inserted at the first column of the named sheet
 * @param {String}  desc:  the description one should give to the item. inserted at the second column of the sheet
 * @param {String}  itemRoom:   cost of the item. inserted at the third column of the sheet.
 * @param {String}  name:    name of the sheet. to be obtained from the UI. case sensitive.
 */
function sheetInsertion(item, desc, cost, name){
  let mSheet = activeSpreadSheet.getSheetByName(name); 
  let inserto = getLastDataRow(mSheet)+1;
  try {
  mSheet.getRange("A"+inserto).setValue(item);
  mSheet.getRange("B"+inserto).setValue(desc);
  mSheet.getRange("C"+inserto).setValue(cost);
  } catch (err){
    Logger.log('Failed with an error %s', + err.message)
  }
}

function rowLastVal(range, firstRow) {
  // range is passed as an array of values from the indicated spreadsheet cells.
  for (var i = range.length - 1;  i >= 0;  -- i) {
    if (range[i] != "")  return i + firstRow;
  }
  return range.length;
}

function itemDesc(){
  sheet=activeSpreadSheet.getSheetByName("Internal");
  if(sheet.getRange("C2:C")===""){
    vLookup();
  }
}

function testingfd(){
  var j=activeSpreadSheet.getSheetByName("Sheet37");
  j.getRange("A19:A").setValue(j.getRange("$A$2:D").getValues());
}

//VLOOKUP(C2,'Item Import'!$A$2:D,2,0)
//C2=value(value is the cell you want to search for of I.E. C2), 'Item Import'!$A$2:D= sheet searchrange, where $A$2:D= is searchrange and 'Item Import'! is sheet, and 2=place, or the location to return the value of.
//grabit is the column of data you wish to get

/**
 * Searches a range for a passed cell value from active sheet in another sheet. 
 * if found sets the specified adjacent cell value from that other sheet into the active sheet's cell.
 * 
 * @param {String}  sheet:  the sheet you want to search
 * @param {String}  itemQty:  the number of item to add
 * @param {String}  itemRoom:   name of room item will be added to.
 */
function vLookup(sheet, value, searchRange, grabit, place){
  var s = activeSpreadSheet.getActiveSheet();     
  var data = activeSpreadSheet.getSheetByName(sheet);
  var searchValue = s.getRange(value).getValue();
  var dataValues = data.getRange(searchRange).getValues();
  var dataList = dataValues.join("ღ").split("ღ");
  var index = dataList.indexOf(searchValue);
  if (index === -1) {
      throw new Error('Value not found')
  } else {
      var foundValue = data.getRange(grabit+(index+2)).getValue();
      s.getRange(place).setValue(foundValue);
  }
}

function importList(linktoimport, startingrowindex, startingcolumnindex, sheetName) {
  //get values to be imported from the linked sheet
  var s = SpreadsheetApp.openByUrl(linktoimport);
  var rowstocopy = getLastDataRow(s);
  var colstocopy = getLastDataCol(s);     
  var values = s.getSheetValues(startingrowindex, startingcolumnindex, rowstocopy, colstocopy);
  var sheetimportto=activeSpreadSheet.getSheetByName(sheetName);
  //set  values imported   
  sheetimportto.getRange(1,1,values.length,values[0].length).setValues(values);
}

function doimp(){
  importList("https://docs.google.com/spreadsheets/d/1xz9Y9EgLcui3ekKkLic-3BC3Z8RS1s4qWvz5NFu6EM4/edit#gid=0", 1, 1, "Custom Sheet");
  importList("https://docs.google.com/spreadsheets/d/1xz9Y9EgLcui3ekKkLic-3BC3Z8RS1s4qWvz5NFu6EM4/edit#gid=0", 1, 1, "Master Sheet");
}

// function onEdit(e) {
//   const row = e.range.getRow();
//   const col = e.range.getColumn();
//   var scolumnlet2 = getLetter(col);
//   // if(e.source.getActiveSheet().getName()==="Sheet37"&& col >= 1 && row=== 3 && e.value=== 'TRUE'){
//   //   alerto ="original cell changed to: "+e.source.getActiveSheet().getRange(scolumnlet2+row).getValue();
//   // }
//   // e.source.getActiveSheet().getRange("C1").setValue(e.source.getActiveSheet().getRange(scolumnlet2+row).getValue());
//   return e.source.getActiveSheet().getRange(scolumnlet2+row).getValue();
// }

// function sendNotification(sheet, changingarea) {
//   var arrg = [];
  // var cellScan = sheet.getRange(changingarea).forEach(cell=>{
  //   cell.getActiveCell().getA1Notation().getValue().toString();
  // });
//   cellScan.forEach(function onEdit(e){
//   const range = e.range;
//   var alert = "";
//   range.forEach(cell =>{
//     alert ="original changed to: "+cell.getValue()
//     arrg.push(alert);
//   });
// });
// return arrg;
// };

function getLastDataCol(sheet) {
  var lastCol = sheet.getLastColumn();
  var colval = getLetter(lastCol);
  var range = sheet.getRange(colval+"1");
  if (range.getValue() !== "") {
    return lastCol;
  } else {
    return range.getNextDataCell(SpreadsheetApp.Direction.NEXT).getColumn();
  }              
}


function getLastDataRow(sheet) {
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange("A" + lastRow);
  if (range.getValue() !== "") {
    return lastRow;
  } else {
    return range.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  }              
}

//end add
//optimised by SS
  function isOdd(num) { return num & 1; };
//end optimzation

  function getFirstEmptyRowWholeRow() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var values = sheet.getRange("D1:E" + sheet.getLastRow()).getValues();
    var row = 0;
    for (var row = 0; row < values.length; row++) {
      if (!values[row].join("")) break;
    }
    Logger.log(row);
    return (row + 1);
  }

function getFirstEmptyBOMRowWholeRow(sheet) {
    var values = sheet.getRange("A1:E" + sheet.getLastRow()).getValues();
    var row = 0;
    for (var row = 0; row < values.length; row++) {
      if (!values[row].join("")) break;
    }
    Logger.log(row);
    return (row + 1);
  }

      function getLetter(num){
      var letter = String.fromCharCode(num + 64);
      return letter;
    }
