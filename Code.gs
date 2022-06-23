function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('dataform')
    .setTitle('Data Validation');
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showSidebar(html);
}
//in on open set the value of the stock checklist checkmarks to false so as to ensure each session has stock refreshed.
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .createMenu('Realm Custom Scripts')
    .addItem('Show Sidebar', 'showSidebar')
    .addToUi();

    var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Prewire Order");
    var tt = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hardware Order");
    var uu = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Add Ons");

    ss.hideColumns(1);
    ss.hideColumns(2);
    ss.hideColumns(12);

    tt.hideColumns(1);
    tt.hideColumns(2);
    tt.hideColumns(12);
    
    uu.hideColumns(1);
    uu.hideColumns(2);
    uu.hideColumns(12);
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

function checkmate(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(), activeSheetName=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetName();
  items = queryASpreadsheet(ss.getId(), activeSheetName, 'SELECT D WHERE F = TRUE AND I = FALSE'), gamer = items.map(function(item) {
  return item.toString();
  });
  if(gamer.length ===0){
    SpreadsheetApp.getUi().alert("Nothing selected for new Order Sheet. please Check off an item's respective 'Order Request' box and try again.");
  }else{
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
        SpreadsheetApp.getUi().alert('No Reference IDs found in the TRXIO sheet, or items available quantity is zero.');
        return;
    }else{
    const templateSheet = ss.getSheetByName('Order List');
    const date = new Date();
    const newSheet = "Stock Pull List - "+ (date.getMonth() + 1) + "/" + date.getDate() +"/" + date.getFullYear();
    //const newSheet = "Order List Dated "+new Date();
    const orderisGiven=ss.insertSheet( newSheet, ss.getSheets().length, {template:templateSheet});
    ss.setActiveSheet(ss.getSheetByName(activeSheetName));
    const orderSheetName = orderisGiven.getName();
    var i = orderisGiven.getLastRow()+1;
    gamer.forEach(name=>{
      const q = "SELECT E WHERE J MATCHES "+name;
      const qu = "SELECT T WHERE J MATCHES "+name;
      const quo = "SELECT J WHERE J MATCHES "+name;
        var jamero = queryASpreadsheet(ss.getId(), 'TRXIO', quo);
        const camero = queryASpreadsheet(ss.getId(), 'TRXIO', qu);
        const gamero = queryASpreadsheet(ss.getId(), 'TRXIO', q);
        jamero = jamero.map(value=>value.replaceAll('""', ''));
      var thestrings = camero.map(function(item) {
        return item.toString();
      });
      var arrOfNum = thestrings.map(str => {
        return Number(str);
      }).reduce((a, b) =>a+b, 0);
        if(arrOfNum>0){
          orderisGiven.getRange("A"+i).setValue(jamero[0])
          orderisGiven.getRange("B"+i).setValue(arrOfNum)
          orderisGiven.getRange("C"+i).setValue(gamero[0])
          i=i+1 
        }else{
          return;
        }
    });
  const checkOff = stockChecklist(activeSheetName, orderSheetName, "D2:D", "A2:A");
  checkOff.forEach(x=>{
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(activeSheetName).getRange("I"+x).setValue(true)
  });
  }
  SpreadsheetApp.getUi().alert("New Order List Created. If an item you checked off wasn't added, it's Reference ID wasn't found on the TRXIO sheet");
  }
}

  // k = k.map((name)=>{
  //   name.replace
  //   return name.replaceAll("\"",'');
  // });
  // k = k.filter(function(o){
  //   if(o!=''){
  //     return o;
  //   }
  // })

function formula(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trix = ss.getSheetByName("Copy of TRXIO");
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

function testRange(){
  var ss = SpreadsheetApp.getActiveSpreadsheet(), activeSheetName=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetName();
  items = queryASpreadsheet(ss.getId(), activeSheetName, 'SELECT D WHERE F = TRUE AND I = FALSE'), gamer = items.map(function(item) {
  return item.toString();
  });
    var counter = 0;
    // gamer.forEach(name=>{
      const tv = "SELECT J WHERE J MATCHES "+"'Transition Networks SISPM1040-3'";
      const b = queryASpreadsheet(ss.getId(), 'TRXIO', tv);
      
    //   counter +=b.length;
    // });
    return b;
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

      function getLetter(num){
      var letter = String.fromCharCode(num + 64);
      return letter;
    }

function onEdit(event) {
  var ss = SpreadsheetApp.getActiveSheet();
  var me = Session.getActiveUser();
  if (event.range.isChecked()){
    var stonk = nextLetter(event.range.getA1Notation()[0]);
    var ston = event.range.getA1Notation().replace(/\D/g,'');
    var stonko = nextLetter(stonk);
    ss.getRange(stonk+ston).setValue(new Date());
    ss.getRange(stonko+ston).setValue(Session.getEffectiveUser().getUsername());
  }
   else if(event.range.isChecked() == false) {
    var stonk = nextLetter(event.range.getA1Notation()[0]);
    var ston = event.range.getA1Notation().replace(/\D/g,'');
    var stonko = nextLetter(stonk);
    ss.getRange(stonk+ston).setValue("");
    ss.getRange(stonko+ston).setValue("");

  } 


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
