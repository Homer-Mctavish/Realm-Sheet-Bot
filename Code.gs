function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('dataform')
    .setTitle('Data Validation Tool');
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showSidebar(html);
}

function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .createMenu('Realm Custom Scripts')
    .addItem('Show Data Validation sidebar', 'showSidebar')
    .addToUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Prewire Order");
  var tt = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hardware Order");
  var uu = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Add Ons");
  ss.hideColumns(1, 3);
  ss,hideColumn("Q");

  tt.hideColumns(1, 3);
  tt.hideColumn("Q");

  uu.hideColumns(1, 3);
  uu.hideColumn("Q");
}

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

function getLetter(num){
  var letter = String.fromCharCode(num + 64);
  return letter;
}

function nextLetter(s){
    return s.replace(/([a-zA-Z])[^a-zA-Z]*$/, function(a){
        var c= a.charCodeAt(0);
        switch(c){
            case 90: return 'A';
            case 122: return 'a';
            default: return String.fromCharCode(++c);
        }
    });
}

function setReservedQuantity(){
  SpreadsheetApp.getActiveSpreadsheet().toast('Adding reserve quantity to TRXIO sheet...');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var trx = ss.getSheetByName("TRXIO");
  var quo = "SELECT R WHERE R LIKE '#%'";
  var namer = queryASpreadsheet(ss.getId(), 'TRXIO', quo);
  var ger = [];
  namer.forEach(name=>{
    var totalResserveQty = 0;
    var quantities = name.match(/\(.*?\)/g)
        quantities = quantities.map(function(match) { 
           quantitiy = match.slice(1, -1);
           totalResserveQty += Number(quantitiy);
      })
      ger.push(totalResserveQty);
  })
  var valus = trx.getRange("R2:R").getValues();
  var data = valus.find(/#/).map(x=>x+2);
  var i = 0;
  data.forEach(index=>{
    trx.getRange("S"+index).setValue(ger[i]);
    i=i+1;
  })
}

function formula(){
  SpreadsheetApp.getActiveSpreadsheet().toast('Adding total quantity value calculation to TRXIO sheet...');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trix = ss.getSheetByName("TRXIO");
  const g = getLastDataRow(trix)
  const a = getLastDataCol(trix)
  const f = getLetter(a);
  var rano = trix.getRange(f+'2').setFormula("=MINUS(O2,S2)")
  rano.copyTo(trix.getRange(f+"2:"+f+g))
}

function stockChecklist(checksheetname, ordersheetname, requestRange, stockRange){
  var joj = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(checksheetname);
  var noj = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ordersheetname);
  var egg = joj.getRange(requestRange).getValues();
  var neg = noj.getRange(stockRange).getValues();
  var gmp = egg.join("ღ").split("ღ").flat();
  var mpo = neg.join("ღ").split("ღ").flat();
  const g = [...new Set(gmp)];
  var k = [...new Set(mpo)];
  k = k.map((name)=>{
    name.replace
    return name.replaceAll("\"",'');
  });
  k = k.filter(function(o){
    if(o!=''){
      return o;
    }
  })
  var lindices = [];  
  const gerb = g.filter(function(thempo, index){
    if(k.includes(thempo)){
      lindices.push(index+2);
    }
  });
  return lindices;
}

function setRowColors(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var range = sheet.getDataRange();
  var lRow = sheet.getLastRow();
  var numRows = lRow - 1;
  var numCols = sheet.getLastColumn();
  var [rows1d, cols1d] = [numRows, numCols].map(function(num){ 
    return Array.apply([],new Array(num)); 
  })
  
  var colors2d = rows1d.map(function(row, i){
    var color = i%2 === 0 ? "#ffffff" : "#efefef";
    return cols1d.map(function(col){
        return color; 
    })
  })
  range.setHorizontalAlignment('left');
  range.setVerticalAlignment('top');
  sheet.getRange(1,1,1,numCols).setBackground('#000000').setFontColor('#FFFFFF').setFontSize('16');
  sheet.getRange(2, 1, numRows, numCols).setBackgrounds(colors2d).setFontSize('14');
  
  sheet.setColumnWidth(5,900);
  sheet.getRange("E:E").setWrap(true);
  sheet.autoResizeColumn(1);
  SpreadsheetApp.flush();
  }

function checkmate(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(), activeSheetName=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetName();
  items = queryASpreadsheet(ss.getId(), activeSheetName, 'SELECT D WHERE F = TRUE AND Q = FALSE'), gamer = items.map(function(item) {
  return item.toString();
  });
  if(gamer.length ===0){
    SpreadsheetApp.getUi().alert("Nothing new selected (or nothing at all). please Check off an item's respective 'Order Request' box and try again.");
  }else{
    SpreadsheetApp.getActiveSpreadsheet().toast('Checking for TRXIO ID and available quantity...');
    var counter = 0;
    var summa = 0;
    gamer.forEach(name=>{
      const qu = "SELECT T WHERE J MATCHES "+name;
      const tv = "SELECT J WHERE J MATCHES "+name;
      const b = queryASpreadsheet(ss.getId(), 'TRXIO', tv);
      let camero = queryASpreadsheet(ss.getId(), 'TRXIO', qu);
      let thestrings = camero.map(function(item) {
        return item.toString();
      });
      let arrOfNum = thestrings.map(str => {
        return Number(str);
      }).reduce((a, b) =>a+b, 0);
      counter +=b.length;
      summa +=arrOfNum;
    });
    if(counter===0 || summa===0){
        SpreadsheetApp.getUi().alert("No Reference IDs found in the TRXIO sheet, or items' available quantity is zero.");
        return;
    }else{
    SpreadsheetApp.getActiveSpreadsheet().toast('Creating new Stock Pull List...');
    const date = new Date();
    const newSheet = "Stock Pull List - "+ date.toLocaleDateString()+": "+date.getMilliseconds();
    const orderisGiven=ss.insertSheet(newSheet, ss.getSheets().length);
    orderisGiven.getRange("A1").setValue("Ref #");
    orderisGiven.getRange("B1").setValue("Item");
    orderisGiven.getRange("C1").setValue("Qty Requested");
    orderisGiven.getRange("D1").setValue("Qty on Hand");
    orderisGiven.getRange("E1").setValue("Location (Shows Reserved Locations too");
    orderisGiven.getRange("F1").setValue(activeSheetName);
    const orderSheetName = orderisGiven.getName();
    var i = orderisGiven.getLastRow()+1;
    SpreadsheetApp.getActiveSpreadsheet().toast('Getting information from TRXIO sheet...');
    gamer.forEach(name=>{
      const q = "SELECT E WHERE J MATCHES "+name;
      const qu = "SELECT T WHERE J MATCHES "+name;
      const quo = "SELECT J WHERE J MATCHES "+name;
      const quot = "SELECT C WHERE J MATCHES "+name;
      const quote = "SELECT E WHERE D MATCHES "+name;
      var samero = queryASpreadsheet(ss.getId(),activeSheetName, quote);
      var bamero = queryASpreadsheet(ss.getId(), 'TRXIO', quot);
        var jamero = queryASpreadsheet(ss.getId(), 'TRXIO', quo);
        const camero = queryASpreadsheet(ss.getId(), 'TRXIO', qu);
        const gamero = queryASpreadsheet(ss.getId(), 'TRXIO', q);
        jamero = jamero.map(value=>value.replaceAll('""', ''));
      var thestrings = camero.map(function(item) {
        return item.toString();
      });
      var cont = bamero.join(",")
      var arrOfNum = thestrings.map(str => {
        return Number(str);
      }).reduce((a, b) =>a+b, 0);
        if(arrOfNum>0){
          orderisGiven.getRange("A"+i).setValue(jamero[0].replace(/"/g, ""));
          orderisGiven.getRange("B"+i).setValue(gamero[0].replace(/"/g, ""));
          orderisGiven.getRange("C"+i).setValue(samero[0]);
          orderisGiven.getRange("D"+i).setValue(arrOfNum);
          orderisGiven.getRange("E"+i).setValue(cont.replace(/"/g, ""));
          i=i+1;
        }else{
          return;
        }
    });
  const checkOff = stockChecklist(activeSheetName, orderSheetName, "D2:D", "A2:A");
  checkOff.forEach(x=>{
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(activeSheetName).getRange("Q"+x).setValue(true)
  });
  setRowColors(newSheet);
  }
  SpreadsheetApp.getUi().alert("New Order List Created.");
  }
}


function joj(){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Internal for order");
  var range1 = ss.getRange("C2:C").getValues();
  // ss.getRange("H2").setValue(range1[1][0]);
  var range2 = ss.getRange("D2:D").getValues();
  var dataList = range2.join("ღ").split("ღ");
  var index = dataList.indexOf("");
  var indices = [];
  while (index !== -1) {
    indices.push(index);
    index = dataList.indexOf("", index + 1);
  }
  var rList = [];
  indices.forEach(idx=>{
    var g = 2;
    var i = 0;
    if(range1[idx][0] != range2[idx][0]){
      rList.push(range1[idx][i]);
      // ss.getRange("H"+g).setValue(range1[idx][i]);
      g += 1;
      i += 1;
    }
  });
  return rList;
}

function matcher(x, goo, ber){
  if(goo === x){
    return ber;
  }else{
    return x;
  }
}

function findAndReplace(rangeo, found, replaced){
  let gop = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(rangeo);
  let array = gop.getValues();
  let doubledArray = array.map(([a]) => [matcher(a, found, replaced)]);
  gop.setValues(doubledArray);
  SpreadsheetApp.getActiveSpreadsheet().toast('Items replaced');
}

function onEdit(event) {
  var ss = SpreadsheetApp.getActiveSheet();
  if (event.range.isChecked()){
    var stonk = nextLetter(event.range.getA1Notation()[0]);
    var ston = event.range.getA1Notation().replace(/\D/g,'');
    var stonko = nextLetter(stonk);
    ss.getRange(stonk+ston).setValue(new Date());
    ss.getRange(stonko+ston).setValue(Session.getEffectiveUser().getUsername());
  } 
  // else if(event.range.isChecked() == false) {
  //   var stonk = nextLetter(event.range.getA1Notation()[0]);
  //   var ston = event.range.getA1Notation().replace(/\D/g,'');
  //   var stonko = nextLetter(stonk);
  //   ss.getRange(stonk+ston).setValue("");
  //   ss.getRange(stonko+ston).setValue("");

  // }  
}

