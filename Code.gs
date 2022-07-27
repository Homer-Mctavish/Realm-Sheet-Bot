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
    setImportRange("https://docs.google.com/spreadsheets/d/1xz9Y9EgLcui3ekKkLic-3BC3Z8RS1s4qWvz5NFu6EM4/edit#gid=0", "1gZvsAIcLfsbiCG0--cDgkl5wqx4X-8L2Zpf-6FGFyWY","Master Sheet");
    setImportRange("https://docs.google.com/spreadsheets/d/1xz9Y9EgLcui3ekKkLic-3BC3Z8RS1s4qWvz5NFu6EM4/edit#gid=0", "1gZvsAIcLfsbiCG0--cDgkl5wqx4X-8L2Zpf-6FGFyWY","Custom Sheet");

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

//added by SS
const activeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
const activeSheet = SpreadsheetApp.getActiveSheet();



Array.prototype.find = function(regex) {
  const arr = this;
  const matches = arr.filter( function(e) { return regex.test(e); } );
  return matches.map(function(e) { return arr.indexOf(e); } );
};

const deepGet = (obj, keys) =>
  keys.reduce(
    (xs, x) => (xs && xs[x] !== null && xs[x] !== undefined ? xs[x] : null),
    obj
  );

//sheetId, sheetName, queryString
function queryASpreadsheet(sheetId, sheetName, queryString) {
 var url = 'https://docs.google.com/spreadsheets/d/'+sheetId+'/gviz/tq?'+
            'sheet='+sheetName+
            '&tq=' + encodeURIComponent(queryString);
  var params = {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  };
  var ret  = UrlFetchApp.fetch(url, params).getContentText();
  var k = JSON.parse(ret.replace("/*O_o*/", "").replace("google.visualization.Query.setResponse(", "").slice(0, -2));
  var depp = deepGet(k, ['table','rows']);
  var arr = [];
  depp.forEach(column=>{
    arr.push(JSON.stringify(column['c'][0].v))
  });
  return arr;
}

function returneo(){
  let ss = SpreadsheetApp.getActiveSpreadsheet().getId();
  let query = "SELECT A WHERE A IS NOT NULL ";
  let hur = queryASpreadsheet("1xz9Y9EgLcui3ekKkLic-3BC3Z8RS1s4qWvz5NFu6EM4", "BOM", query);
  const dur = queryASpreadsheet(ss, "BOM", query); 
  const hurdur = hur.concat(dur);
  return hurdur;
}

function setImportRange(url, iD, sheetName){
  let sheeto = activeSpreadSheet.getSheetByName(sheetName);
  const importation = SpreadsheetApp.openById(iD); 
  let lastRow = getLastDataRow(importation);
  var isit = sheeto.getRange("A1").setFormula('=IMPORTRANGE("'+url+'","items!A1:E'+lastRow+'")');
  }

function onlyNumbers(array) {
  return array.every(element => {
    return typeof element === 'number';
  });
}

function addRow(){
  var sheet = activeSheet;
  var range = sheet.getActiveRange();
    try {
    let fill = sheet.getRange("2:2");
    SpreadsheetApp.flush();
    sheet.insertRowsAfter(range.getLastRow(), range.getValues().length);
    let row1=range.getRow()
    let row2=range.getLastRow();
    let dist = row2-row1;
    let nrow1 = row1+dist+1;
    let nrow2 = row2+dist+1;
    let stringi = nrow1+":"+nrow2;
    let itbe = sheet.getRange(stringi);
    SpreadsheetApp.flush();
    fill.copyTo(itbe, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
    return stringi;
  } catch (err){
    Logger.log('Failed with an error %s', + err.message)
  }
}

function setAllFormulas(){
  let sheet = activeSpreadSheet.getSheetByName("Internal");
  try{
    
    let fill = sheet.getRange("F2:U2");
    let all = sheet.getRange("F3:U"+getLastDataRow(sheet));
    fill.copyTo(all, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  }catch(err){
    console.log('Failed with stupid error %s', + err.message);
  }
}

let porkoso = 0;

function onEdit(){
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Copy of Item Import");
  let vealue = sheet.getRange("B2:B").getValues().flat();
  let teb = sheet.getRange("F2:F").getValues().flat();
  let boba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Internal");
  let chonk = boba.getRange("C3:C").getValues().flat();
    for(i in chonk){
      if(vealue.indexOf(chonk[i])!==-1 && teb.length>porkoso &&i<=chonk.length){
        let koe = vealue.indexOf(chonk[i]+2);
        SpreadsheetApp.getActiveSpreadsheet().toast("price of "+chonk[i]+" has changed: "+sheet.getRange("F"+koe).getValue()+"new price: "+sheet.getRange("C"+koe).getValue());
      }
    }
  SpreadsheetApp.getActiveSpreadsheet().toast(vealue.length);
}

function idiob(){
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Copy of Item Import");
  let vealue = sheet.getRange("A2:A").getValues().flat();
  let boba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Internal");
  let chonk = boba.getRange("C3:C").getValues().flat();
  let kf = [];
  for(i in chonk){
    if(vealue.indexOf(chonk[i])!==-1){
      kf.push(vealue.indexOf(chonk[i]+2));
    }
  }
  return kf;
}

// function bolbina(){
//   let sheet = activeSpreadSheet.getSheetByName("Internal")
//   let protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
//   let range = sheet.getRange("F2:U2");
//     //Iterate through all sheet's protections
//   for(let i = 0; i < protections.length; i++) {
//       let protection = protections[i];
//       if (protection.getRange().getA1Notation() == range.getA1Notation()) {
//         protection.remove();
//       }
//     }
//   }

function getItemList() {
    var sheet = activeSpreadSheet.getSheetByName("Item Import");
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
  let sheet = activeSheet;
  let srow = sheet.getActiveRange().getRow();
  let scolumn = sheet.getActiveRange().getColumn();

  //change scolumn to letter, change s
  let scolumnlet1 = getLetter(scolumn-2);
  let scolumnlet2 = getLetter(scolumn+1);
  activeSpreadSheet.getRange(scolumnlet1+srow).setValue(itemRoom.toUpperCase());
  SpreadsheetApp.getActiveRange().setValue(selectedItemToPaste);
  activeSpreadSheet.getRange(scolumnlet2+srow).setValue(itemQty);
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
  var result = ui.prompt("Please input BOM Type Name");
  var bomName = result.getResponseText();
  //make sure BOM Type doesn't already exist
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bomSheet = ss.getSheetByName("BOM");
  var bomSheetLastRow = bomSheet.getLastRow()+1;
  var bomSheetValue = bomSheet.getRange("A2:" + "ZZ" + bomSheetLastRow).getValues();
  let ohNoUserBadInput = false;

  for (var i = 0; i < (bomSheetLastRow - 1); i++) {
    var bomType = bomSheetValue[i][0];
    if (bomName===""){
      ohNoUserBadInput = false;
      return;
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
function sheetInsertion(item, desc, cost, msrp, name){
  let mSheet = activeSpreadSheet.getSheetByName(name); 
  let inserto = getLastDataRow(mSheet)+1;
  try {
  mSheet.getRange("A"+inserto).setValue(item);
  mSheet.getRange("B"+inserto).setValue(desc);
  mSheet.getRange("C"+inserto).setValue(cost);
  mSheet.getRange("D"+inserto).setValue(msrp);
  } catch (err){
    Logger.log('Failed with an error %s', + err.message)
  }
}

function itemDesc(){
  sheet=activeSpreadSheet.getSheetByName("Internal");
  if(sheet.getRange("C2:C")===""){
    vLookup();
  }
}

function querySearch(v, scol, wcol, sheetName){
    let iD = activeSpreadSheet.getId();
    let val = "'"+v+"'";
    const query = "SELECT "+scol+" WHERE "+wcol+" MATCHES "+val;
    let vlook = queryASpreadsheet(iD, sheetName, query);
    return vlook[0];
}


/**
 * Searches a column for a passed columns' values from active sheet in another sheet. 
 * if found sets the specified value from that other sheet into active sheet's approprite column cell.
 * 
 * @param {String}  searchSheet:  the name of the sheet you want to search
 * @param {String}  sourceSheet:  the name of the sheet you want to search from 
 * @param {String}  colSearch:  the column you want to search with, for if colMatch column has values identical to the ones in the sourceSheet
 * @param {String}  colMatch:   the column you want the values from, if the values in the ColSearch column are identical to the source sheets' values
 * @param {String}  sourceSheet:  the name of the sheet you want to search from 
 */

function newVlookup(searchSheet, sourceSheet, colVal, colMatch, colSearch, colWrite){
  let gh = activeSpreadSheet.getSheetByName(sourceSheet).getRange(colVal+"1:"+colVal).getValues();
  let values = gh.filter(String);
  var i = 3;
  values.forEach(name=>{
    let hgp = querySearch(name, colMatch, colSearch, searchSheet);
    activeSpreadSheet.getSheetByName(sourceSheet).getRange(colWrite+i).setValue(hgp);
    i=i+1;
  });  
}

function gpo( queryString, x){
  
  return x.map(arg =>queryString+"'"+arg+"'");
}

function matcher(x, goo){
  for(let i = 0; i<goo.length;i++){
    if(goo[i] === x){
      return goo[i];
    }
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

function jod(sheetId, sheetName, d){
  return d.map(arg => 'https://docs.google.com/spreadsheets/d/'+sheetId+'/gviz/tq?'+
            'sheet='+sheetName+
            '&tqx=out:csv' +
            '&tq=' + encodeURIComponent(arg));
}


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

function isOdd(num) { return num & 1; };
//end add

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
