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
    .addItem('Show Estimator sidebar', 'showSidebar')
    .addToUi();
}

function authorizeItemImport(){

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Item Import");
//authroize the data import cell
  var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
      addImportrangePermission_(ssId,'1xz9Y9EgLcui3ekKkLic-3BC3Z8RS1s4qWvz5NFu6EM4');

//SpreadsheetApp.getUi().alert("running imoprt");

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
function doGet() {
    return HtmlService.createTemplateFromFile('itemform.html')
        .evaluate() // evaluate MUST come before setting the Sandbox mode
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function runRealmItemAdd() {
      //this will create and show our html file
   var t = HtmlService.createTemplateFromFile('itemform');
    var htmlOutput = t.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setHeight(600).setWidth(900);
    var doc = SpreadsheetApp.getActive();
    doc.show(htmlOutput);
  
  
  /*  var html = HtmlService.createHtmlOutputFromFile('ClientMusicForm')
      .setTitle('Client Service Request')
      .setWidth(450);
       SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html); */
  
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

function addForumlas(){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var sheetName = ss.getActiveSheet().getName();


  if (sheetName == "Internal") {     
    var activeColumn = sheet.getActiveCell().getColumn();

    // when the itemcode  changes we look up the descriptino and cost
    if (activeColumn == 3) {
      

      var activeRow = sheet.getActiveCell().getRow(); 

      var itemcodeSelected = sheet.getRange(activeRow, 3).getValue();
      //SpreadsheetApp.getUi().alert(itemcodeSelected);

      sheet.getRange(activeRow,5).setFormula("=IF(C" + activeRow + " = \"\",\"\",QUERY('Item Import'!A2:D,\"SELECT B WHERE A = '" + itemcodeSelected + "'\"; 0))");
      sheet.getRange(activeRow,6).setFormula("=ROUND(L" + activeRow + ",-1)");
      sheet.getRange(activeRow,7).setFormula("=IF(C" + activeRow + "=\"\",\"\",SUMIF(VLOOKUP(Internal!C" + activeRow + ",'Item Import'!A2:D,3,0),\"<>#N/A\"))");
      sheet.getRange(activeRow,8).setFormula("=G" + activeRow + "*D" + activeRow + "");
      sheet.getRange(activeRow,9).setFormula("=IF(C" + activeRow + "=\"\",\"\",SUMIF(VLOOKUP(Internal!C" + activeRow + ",'Item Import'!A2:D,4,0),\"<>#N/A\"))");  
      sheet.getRange(activeRow,10).setFormula("=I" + activeRow + "*D" + activeRow + "");
      sheet.getRange(activeRow,11).setFormula("=I" + activeRow + "*(1-'Project Calcs'!$C$10)"); 
      sheet.getRange(activeRow,12).setFormula("=K" + activeRow + "*D" + activeRow + ""); 
      sheet.getRange(activeRow,13).setFormula("=J" + activeRow + "*'Project Calcs'!$C$6"); 
      sheet.getRange(activeRow,14).setFormula("=J" + activeRow + "*'Project Calcs'!$C$3");  
      sheet.getRange(activeRow,15).setFormula("=N" + activeRow + "*'Project Calcs'!$C$4"); 
      sheet.getRange(activeRow,16).setFormula("=N" + activeRow + "*'Project Calcs'!$C$5*(1-'Project Calcs'!$C$9)"); 
      sheet.getRange(activeRow,17).setFormula("=L" + activeRow + "-H" + activeRow + ""); 
      sheet.getRange(activeRow,18).setFormula("=P" + activeRow + "-O" + activeRow + "");
      sheet.getRange(activeRow,19).setFormula("=M" + activeRow + "*'Project Calcs'!$C$7");  
      sheet.getRange(activeRow,20).setFormula("=R" + activeRow + "+Q" + activeRow + "+S" + activeRow);  
      sheet.getRange(activeRow,21).setFormula("=IF(Internal!B"+ activeRow +"=\"\",\"\",Internal!I"+ activeRow +"*'Project Calcs'!$C$8)");

    }
  };
}
*/
//added by SS
const activeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
const activeSheet = SpreadsheetApp.getActiveSheet();


function getItemList() {
    var sheet = activeSpreadSheet.getSheetByName("Item Import");
    var data = sheet.getDataRange().getValues()
    var array = [];
    data.forEach(function(row){array.push([row[0],row[1]]); });
   // Logger.log(array);
    return array;
}
//edited by SS
function addItems(selectedItemToPaste,itemQty,itemRoom){ 
  let sheet = activeSpreadSheet.getActiveSheet();
  srow = sheet.getActiveRange().getRow();
  scolumn = sheet.getActiveRange().getColumn();
  
  //SpreadsheetApp.getUi().alert(srow + " " + scolumn );
  activeSpreadSheet.getRange(srow,scolumn-2).setValue(itemRoom);
  SpreadsheetApp.getActiveRange().setValue(selectedItemToPaste);
  activeSpreadSheet.getRange(srow,scolumn+1).setValue(itemQty);
 // addForumlas();
  sheet.setActiveRange(sheet.getRange(srow+1,scolumn));
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //sheet.getRange(e.range.getRow(), col).setValue(selectedItemToPaste);

}

//added by SS
/*
function removeItems(itemQty, itemRoom){
  let sheet = activeSpreadSheet.getActiveSheet();
  srow = sheet.getActiveRange().getRow();
  scolumn = sheet.getActiveRange().getColumn();
  
  activeSpreadSheet.getRange(srow,scolumn-2).setValue(itemRoom);
  SpreadsheetApp.getActiveRange.setValue("");
  activeSpreadSheet.getRange(srow, scolumn);
  sheet.setActiveRange(sheet.getRange(srow+1, scolumn));
 }
*/

function getBOMList() {
  var ss = SpreadsheetApp.openById("1xz9Y9EgLcui3ekKkLic-3BC3Z8RS1s4qWvz5NFu6EM4"); 
  var roomTypeSheet = ss.getSheetByName("BOM");
  var getLastRow = roomTypeSheet.getLastRow();
  var data = roomTypeSheet.getRange(2, 1, getLastRow - 1, 2).getValues();
 // Logger.log(data);
  return data;
}



function addBOMtoTemplate() {
  var ui = SpreadsheetApp.getUi();
  var reAdlt = ui.prompt("Please BOM Type");
  var bomName = result.getResponseText();
 
  //we will make sure BOM Type doesn't already exist or that would not be good
  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ss = SpreadsheetApp.openById("1xz9Y9EgLcui3ekKkLic-3BC3Z8RS1s4qWvz5NFu6EM4"); 
  var bomSheet = ss.getSheetByName("BOM");
  //var bomSheetLastRow = getFirstEmptyRowWholeRow();
  var bomSheetLastRow = bomSheet.getLastRow()+1;
  var bomSheetValue = bomSheet.getRange("A2:" + "ZZ" + bomSheetLastRow).getValues();
  let ohNoUserBadInput = false;

  for (var i = 0; i < (bomSheetLastRow - 1); i++) {
    var bomType = bomSheetValue[i][0];
    if (bomName === bomType ){
      ohNoUserBadInput = true;
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

function openDialog() {
  var html = HtmlService.createTemplateFromFile('dataform')
    .evaluate();

    SpreadsheetApp.getUi() 
    .showModalDialog(html, 'Dialog title');
}

function include(File) {
  return HtmlService.createHtmlOutputFromFile(File).getContent();
};

// function newItemAddition(){
//   let ui = 


// }

function insertItems(selectedRoomNameInput, selectedBomType) {

    var selectedRoomNames = [];
    selectedRoomNames = selectedRoomNameInput.split(",");
    var selectedBomType = selectedBomType;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var bs = SpreadsheetApp.openById("1xz9Y9EgLcui3ekKkLic-3BC3Z8RS1s4qWvz5NFu6EM4"); 
    var internalSheet = ss.getSheetByName("Internal");

    var internalLastRow =  getLastDataRow(internalSheet) +2;
    //var internalLastRow = getFirstEmptyRowWholeRow();

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
    var bomSheet = bs.getSheetByName("BOM");
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
function masterSheet(item, desc, cost){
  let mSheet = activeSpreadSheet.getSheetByName("Master Sheet"); 
  // let insertCell=rowLastVal(mSheet.getRange("A2:A"), 2);
  let inserto = mSheet.getActiveRange().getValues().length;
  mSheet.getRange("A",inserto).setValue(item);
  mSheet.getRange("B",inserto).setValue(desc);
  mSheet.getRange("C",inserto).setValue(cost);
}

function customSheet(item, desc, cost){
  let mSheet = activeSpreadSheet.getSheetByName("Custom Sheet") 
  let inserto = mSheet.getActiveRange().getValues().length;
  mSheet.getRange("A",inserto).setValue(item);
  mSheet.getRange("B",inserto).setValue(desc);
  mSheet.getRange("C",inserto).setValue(cost);
}


// used like so: rowWithLastValue("A2:A", 2)
// function rowWithLastValue(sheet, firstRow) {
//   // range is passed as an array of values from the indicated spreadsheet cells.
//   var arrayb=[];
//   var lastRow = sheet.getLastRow();
//   var data = sheet.getRange(1, 1, lastRow, 1).getValues(); //getRange(starting Row, starting column, number of rows, number of columns)
//   for(var i=0;i<(lastRow-1);i++)
//     {
//       arrayb.push(data[0][i]);
//     }
//   for (var i = arrayb.length - 1;  i >= 0;  -- i) {
//     if (arrayb[i] != "")  return i + firstRow;
//   }
//   return firstRow;
// }

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

function vLookup(){
  var s = activeSpreadSheet.getActiveSheet();     
  var data = s.getSheetByName("Item Import");
  var searchValue = s.getRange("B2").getValue();
  var dataValues = data.getRange("B2:B").getValues();
  var dataList = dataValues.join("ღ").split("ღ");
  var index = dataList.indexOf(searchValue);
  if (index === -1) {
      throw new Error('Value not found')
  } else {
      var row = index + 3;
      var foundValue = data.getRange("E"+row).getValue();
      s.getRange("E2:E").setValue(foundValue);
  }
}

function sendNotification(row, col, changeCol) {
  var sheet = activeSpreadSheet.getSheetByName("Custom Sheet");
  var cellScan = sheet.getRange(row).forEach(cell=>{
    cell.getActiveCell().getA1Notation().getValue().toString();
  })
  // var message = '';
  if(cellScan.indexOf(col)!=-1){ 
    message = sheet.getRange(changeCol+ sheet.getActiveCell().getRowIndex()).getValue()
  }
  return message;
};

  function isOdd(num) { return num & 1; };

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

function getLastDataRow(sheet) {
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange("A" + lastRow);
  if (range.getValue() !== "") {
    return lastRow;
  } else {
    return range.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  }              
}
