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

//keep track of items with zero quantyity, if they are checked and passed, notify the user that certain items they attempted to add have not been added due to lack of available quantity
// function checkmate(){
//   var orderisGiven = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Order List'), activeSheetName=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetName(),
//   items = queryASpreadsheet(SpreadsheetApp.getActiveSpreadsheet().getId(), activeSheetName, 'SELECT D WHERE F = TRUE AND I = FALSE'), gamer = items.map(function(item) {
//   return item.toString();
//   });
//   if(gamer.length !=0){
//   var i = orderisGiven.getLastRow()+1;
//   gamer.forEach(name=>{
//     var q = "SELECT E WHERE J MATCHES "+name;
//     var qu = "SELECT T WHERE J MATCHES "+name;
//     var quo = "SELECT J WHERE J MATCHES "+name;
//     var gamero = queryASpreadsheet(SpreadsheetApp.getActiveSpreadsheet().getId(), 'TRXIO', q);
//     var camero = queryASpreadsheet(SpreadsheetApp.getActiveSpreadsheet().getId(), 'TRXIO', qu);
//     var jamero = queryASpreadsheet(SpreadsheetApp.getActiveSpreadsheet().getId(), 'TRXIO', quo);
//     if(Number(camero[0])>0){
//       orderisGiven.getRange("A"+i).setValue(jamero[0])
//       orderisGiven.getRange("B"+i).setValue(camero[0])
//       orderisGiven.getRange("C"+i).setValue(gamero[0])
//       i=i+1 
//     }else{
//       return;
//     }
//   });
//   const checkOff = stockChecklist(activeSheetName, "Order List", "D2:D", "A2:A");
//   checkOff.forEach(x=>{
//     SpreadsheetApp.getActiveSpreadsheet().getSheetByName(activeSheetName).getRange("I"+x).setValue(true)
//   })
//   }else{
//     return "idiot";
//   }
// }

function checkmate(){
  var activeSheetName=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetName(), ss=SpreadsheetApp.getActiveSpreadsheet(),
  items = queryASpreadsheet(ss.getId(), activeSheetName, 'SELECT D WHERE F = TRUE AND I = FALSE'), gamer = items.map(function(item) {
  return item.toString();
  });
  if(gamer.length !=0){
    var templateSheet = ss.getSheetByName('Order List'), orderisGiven=ss.insertSheet(1, {template:templateSheet}), orderSheetName = orderisGiven.getName();
    var i = orderisGiven.getLastRow()+1;
    gamer.forEach(name=>{
      var q = "SELECT E WHERE J MATCHES "+name;
      var qu = "SELECT T WHERE J MATCHES "+name;
      var quo = "SELECT J WHERE J MATCHES "+name;
      var gamero = queryASpreadsheet(ss.getId(), 'TRXIO', q);
      var camero = queryASpreadsheet(ss.getId(), 'TRXIO', qu);
      var jamero = queryASpreadsheet(ss.getId(), 'TRXIO', quo);
      if(Number(camero[0])>0){
        orderisGiven.getRange("A"+i).setValue(jamero[0])
        orderisGiven.getRange("B"+i).setValue(camero[0])
        orderisGiven.getRange("C"+i).setValue(gamero[0])
        i=i+1 
      }else{
        return;
      }
  });
  const checkOff = stockChecklist(activeSheetName, orderSheetName, "D2:D", "A2:A");
  checkOff.forEach(x=>{
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(activeSheetName).getRange("I"+x).setValue(true)
  })
  }else{
    //the items you are attempting to add to an order sheet have already been added previously.
    SpreadsheetApp.getUi().alert('Nothing added to new Order Sheet. either you have added these items to a prior sheet or the Trxio sheet is misnamed/not formatted correctly.');
  }
}

function obtainListofCheckedwithoutStock(ss){
  var id = '1-YBuCQ7bRuJbE3eiP8RmkfXYSaGZqQHhowRsBEIb-5o';
  var qu = 'SELECT D WHERE F = TRUE AND I = FALSE';
  var data = queryASpreadsheet(id, ss, qu);
  return data;
}

function itDoesIt(){
  var name = 'Middle Atlantic UFAF-1';
  var f = "'"+name+"'";
  var q = "SELECT E WHERE J MATCHES "+f;
  var qu = "SELECT O WHERE J MATCHES"+f;
  var quo = "SELECT J WHERE J MATCHES"+f;
  var gamer = queryASpreadsheet(SpreadsheetApp.getActiveSpreadsheet().getId(), 'Hardware order', 'SELECT D WHERE F = TRUE AND I = FALSE');
  if(gamer.length !=0){
  let orderl = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Order list');
  var gamero = queryASpreadsheet(SpreadsheetApp.getActiveSpreadsheet().getId(), 'TRXIO', q);
  var camero = queryASpreadsheet(SpreadsheetApp.getActiveSpreadsheet().getId(), 'TRXIO', qu);
  var jamero = queryASpreadsheet(SpreadsheetApp.getActiveSpreadsheet().getId(), 'TRXIO', quo);
  orderl.getRange("A2").setValue(jamero[0])
  orderl.getRange("B2").setValue(camero[0])
  orderl.getRange("C2").setValue(gamero[0])
  }
}

function uncheckit(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var goj = ss.getSheetByName("Hardware order").getRange("F2:F")
  var gem = ss.getSheetByName("Hardware order").getRange("D2:D")
  var geb = ss.getSheetByName("Order List").getRang("A2:A")

}

function wee(){
  var joj=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Order List");
  var items = queryASpreadsheet(SpreadsheetApp.getActiveSpreadsheet().getId(), 'Hardware order', 'SELECT D WHERE F = TRUE AND I = FALSE');
  var names = items.map(function(item) {
  return item.toString();
  });
  var i =joj.getLastRow()+1;
  names.forEach(name=>{
    var q = "SELECT E WHERE J MATCHES "+name;
    var qu = "SELECT T WHERE J MATCHES "+name;
    joj.getFilter()
    let gamero = queryASpreadsheet(SpreadsheetApp.getActiveSpreadsheet().getId(), 'TRXIO', q)
    let camero = queryASpreadsheet(SpreadsheetApp.getActiveSpreadsheet().getId(), 'TRXIO', qu)
    joj.getRange("E"+i).setValue(gamero[0]);
    joj.getRange("F"+i).setValue(camero[0]);
    joj.getRange("G"+i).setValue(name);
    i=i+1;
  })
  return names;
}

// function stockChecklist(sheetname, requestRange, stockRange){
//   joj = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname)
//   var arbys = joj.getRange(requestRange).getValues();
//   var bob = joj.getRange(stockRange).getValues();
//   var bar = arbys.join("ღ").split("ღ").flat();
//   var lop = bob.join("ღ").split("ღ").flat();  
//   var indices = [];
//   var bindices = [];
//   const hop = bar.filter(function(yourArray, index) {
//  if(yourArray === "true"){
//    indices.push(index+2)
//  }
// });
//   const nop = lop.filter(function(yourArray, index) {
//  if(yourArray === "false"){
//    bindices.push(index+2)
//  }
// });
//   const filteredArray = indices.filter(value =>bindices.includes(value));
//   return filteredArray;
// }

function stockChecklist(checksheetname, ordersheetname, requestRange, stockRange){
  var joj = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(checksheetname);
  var noj = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ordersheetname);
  var egg = joj.getRange(requestRange).getValues();
  var neg = noj.getRange(stockRange).getValues();
  var gmp = egg.join("ღ").split("ღ").flat();
  var mpo = neg.join("ღ").split("ღ").flat();
  const g = [...new Set(gmp)];
  const k = [...new Set(mpo)];
  var lindices = [];  
  const gerb = g.filter(function(thempo, index){
    if(k.includes(thempo)){
      lindices.push(index+1);
    }
  });
  return lindices;
}

  // const filteredArray = indices.filter(value =>bindices.includes(value));
  // filteredArray.forEach(x=>{
  //   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hardware order").getRange("I"+x).setValue(true)
  // })
function testRange(){
  var joj = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hardware order");
  var noj = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Order List");
  var egg = joj.getRange("D2:D").getValues();
  var neg = noj.getRange("A2:A").getValues();
  var gmp = egg.join("ღ").split("ღ").flat();
  var mpo = neg.join("ღ").split("ღ").flat();
  const g = [...new Set(gmp)];
  const k = [...new Set(mpo)];
  var lindices = [];  
  const gerb = g.filter(function(thempo, index){
    if(k.includes(thempo)){
      lindices.push(index);
    }
  });
  return lindices;
}

function setReservedQuantity(){
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


//list is the getvalues of the colum, item is whatever you're looking for and column is where your wanting to put it

function binarySearch(list, item,column) {
    var min = 0;
    var max = list.length - 1;
    var guess;
    var column = column || 0
    while (min <= max) {
        guess = Math.floor((min + max) / 2);

        if (list[guess][column] === item) {
            return guess;
        }
        else {
            if (list[guess][column] < item) {
                min = guess + 1;
            }
            else {
                max = guess - 1;
            }
        }
    }
    return -1;
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


// // SAM on Edit 
// function onEdit(event) {
//   var ss = SpreadsheetApp.getActiveSheet();
//   var me = Session.getActiveUser();
//   if (event.range.isChecked()){
//     var stonk = nextLetter(event.range.getA1Notation()[0]);
//     var ston = event.range.getA1Notation().replace(/\D/g,'');
//     var stonko = nextLetter(stonk);
//     ss.getRange(stonk+ston).setValue(new Date());
//     ss.getRange(stonko+ston).setValue(Session.getEffectiveUser().getUsername());
//   }
//   //  else if(event.range.isChecked() == false) {
//   //   var stonk = nextLetter(event.range.getA1Notation()[0]);
//   //   var ston = event.range.getA1Notation().replace(/\D/g,'');
//   //   var stonko = nextLetter(stonk);
//   //   ss.getRange(stonk+ston).setValue("");
//   //   ss.getRange(stonko+ston).setValue("");

//   // } 


// }

// function nextLetter(s){
//     return s.replace(/([a-zA-Z])[^a-zA-Z]*$/, function(a){
//         var c= a.charCodeAt(0);
//         switch(c){
//             case 90: return 'A';
//             case 122: return 'a';
//             default: return String.fromCharCode(++c);
//         }
//     });
// }
