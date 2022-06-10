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
//todo: instantiate this variable when the proper spreadsheet is active
// var cellulor = activeSpreadSheet.getSelection().getActiveRangeList().getRanges();

//use the cellulor instead of the getRange to grab all the cells to copy. figure out how the 
function addRow() {
  var ui = SpreadsheetApp.getUi();
  var sheet = activeSpreadSheet.getActiveSheet(); 
  var range = sheet.getActiveRange(); 
  if (range.getFormulas())
  {
    sheet.insertRowsBefore(rowe, rownum);
    let stringo = range.getA1Notation();
    let st = stringo.split(":")[0]
    let ri = stringo.split(":")[1]
    let nu = parseInt(st); 
    let mb = parseInt(ri);
    let er = nu+rownum;
    let is = mb+rownum;
    //range in getRange is one to copy
    range.setValues(sheet.getRange(er+":"+is).getFormulas());
  }else{
  ui.alert("uh, no formulas to copy");
  }
}

function getItemList() {
    var sheet = activeSpreadSheet.getSheetByName("Item Import");
    var data = sheet.getDataRange().getValues()
    var array = [];
    data.forEach(function(row){array.push([row[0],row[1]]); });
   // Logger.log(array);
    return array;
}
//end add

//edited by SS
function addItems(selectedItemToPaste,itemQty,itemRoom){ 
  let sheet = activeSpreadSheet.getActiveSheet();
  let srow = sheet.getActiveRange().getRow();
  let scolumn = sheet.getActiveRange().getColumn();
  //change scolumn to letter, change s
  let scolumnlet1 = getLetter(scolumn-2);
  let scolumnlet2 = getLetter(scolumn+1);
  activeSpreadSheet.getRange(scolumnlet1+srow).setValue(itemRoom.toUpperCase());
  SpreadsheetApp.getActiveRange().setValue(selectedItemToPaste);
  activeSpreadSheet.getRange(scolumnlet2+srow).setValue(itemQty);
}

//added by SS
// function removeItems(itemQty, itemRoom){
//   let sheet = activeSpreadSheet.getActiveSheet();
//   srow = sheet.getActiveRange().getRow();
//   scolumn = sheet.getActiveRange().getColumn();
  
//   activeSpreadSheet.getRange(srow,scolumn-2).setValue(itemRoom);
//   SpreadsheetApp.getActiveRange.setValue("");
//   activeSpreadSheet.getRange(srow, scolumn);
//   sheet.setActiveRange(sheet.getRange(srow+1, scolumn));
//  }


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
  var result = ui.prompt("Please input BOM Type");
  var bomName = result.getResponseText();
 
  //make sure BOM Type doesn't already exist
  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ss = SpreadsheetApp.openById("1xz9Y9EgLcui3ekKkLic-3BC3Z8RS1s4qWvz5NFu6EM4"); 
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

    var selectedRoomNames = [];
    selectedRoomNames = selectedRoomNameInput.split(",");
    var selectedBomType = selectedBomType;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var bs = SpreadsheetApp.openById("1xz9Y9EgLcui3ekKkLic-3BC3Z8RS1s4qWvz5NFu6EM4"); 
    var internalSheet = ss.getSheetByName("Internal");

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
  let inserto = getLastDataRow(mSheet)+1;
  mSheet.getRange("A"+inserto).setValue(item);
  mSheet.getRange("B"+inserto).setValue(desc);
  mSheet.getRange("C"+inserto).setValue(cost);
}

function customSheet(item, desc, cost){
  let mSheet = activeSpreadSheet.getSheetByName("Custom Sheet") 
  let inserto = getLastDataRow(mSheet)+1;
  mSheet.getRange("A"+inserto).setValue(item);
  mSheet.getRange("B"+inserto).setValue(desc);
  mSheet.getRange("C"+inserto).setValue(cost);
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


//change to general formula maintainer using function(sheet, formula)
//also use the insert items foreach loop to do each formula in one pass
//in future make the function take only these things, and the range parameters will be some implementation of a method for future activespreadsheet object
//set the value of the range as protected after filling with formula
// function keepvLookup(start, end){
//   var spreader= activeSpreadSheet.getSheetByName("Sheet37");
//   spreader.getRange(start, end).setValue(vLookup("Sheet37", "C2", ""))

// }

function testingfd(){
  var j=activeSpreadSheet.getSheetByName("Sheet37");
  j.getRange("A19:A").setValue(j.getRange("$A$2:D").getValues());
}
//VLOOKUP(C2,'Item Import'!$A$2:D,2,0)
//C2=start(start is the cell you want to grab the value of I.E. C2), 'Item Import'!$A$2:D= start:end, range of items you want to vlookup, and 2=index, or the location to start searching
//
function vLookup(sheet, start, end){
  var s = activeSpreadSheet.getActiveSheet();     
  var data = s.getSheetByName(sheet);
  var searchValue = s.getRange(start).getValue();
  var dataValues = data.getRange(end).getValues();
  var dataList = dataValues.join("ღ").split("ღ");
  var index = dataList.indexOf(searchValue);
  if (index === -1) {
      throw new Error('Value not found')
  } else {
      var row = index + 3;
      var foundValue = data.getRange(inputend+row).getValue();
      s.getRange(inputstart+":"+inputend).setValue(foundValue);
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

      function getLetter(num){
      var letter = String.fromCharCode(num + 64);
      return letter;
    }
//end add


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
