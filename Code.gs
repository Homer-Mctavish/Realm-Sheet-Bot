function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('dataform')
    .setTitle('Data Validation');
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showSidebar(html);
}

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
//   // var ge = gg.map((c,i)=>c==true? `F${i+2}`:'').filter(c=>c!='');
// function checkcheckbox(){
//   var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(), gg;
//   gg=ss.getRange("F2:F").getValues().flat();
//   var ge = gg.map((c,i)=>{if(c===true){return i+2;}else{return '';}}).filter(c=>c!='');
//   if(ge.length!==0){
//     var newOne = SpreadsheetApp.getActiveSpreadsheet();
//     let datto = new Date();
//     let templateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Order List");
//     newOne.insertSheet('Order List of '+datto, 10, {template: templateSheet});
//     SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(newOne.getSheetByName("Hardware order"));
//     ge.forEach(number=>{
//       if(ss.getRange("L"+number).getValue()===""){
//         var trix = newOne.getSheetByName("TRXIO");
//         var gero =  searchTrxio(trix, "E2:E", ss.getRange("D"+number).getValue());
//         if(gero!==-1){
//           let refNo =gero+2;
//           let loc = trix.getRange("C"+refNo).getValue();
//           let qty = trix.getRange("O"+refNo).getValue();
//           let item = trix.getRange("E"+refNo).getValue();

//           newOne.getSheetByName('Order List of '+datto).getRange("A"+getLastDataRow(newOne)).setValue(loc);
//           newOne.getSheetByName('Order List of '+datto).getRange("B"+getLastDataRow(newOne)).setValue(qty);
//           newOne.getSheetByName('Order List of '+datto).getRange("C"+getLastDataRow(newOne)).setValue(item);
//           ss.getRange("L"+number).setValue(1);
//         }else{
//           return;
//         }
//       }else{
//         return;
//       }
//     });
//   }else{
//     return;
//   }
// }
///\(.*?\)/g

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
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(), trix = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TRXIO");
  var gamer = queryASpreadsheet(SpreadsheetApp.getActiveSpreadsheet().getId(), 'Hardware order', 'SELECT D WHERE F = TRUE AND I = FALSE');
  if(gamer.length !=0){
  gamer.forEach(name=>{
    var f = "'"+name+"'";
    var q = "SELECT E WHERE J MATCHES "+f;
    var qu = "SELECT O WHERE J MATCHES"+f;
    var quo = "SELECT J WHERE J MATCHES"+f;
    var gamero = queryASpreadsheet(SpreadsheetApp.getActiveSpreadsheet().getId(), 'TRXIO', q);
    var camero = queryASpreadsheet(SpreadsheetApp.getActiveSpreadsheet().getId(), 'TRXIO', qu);
    var jamero = queryASpreadsheet(SpreadsheetApp.getActiveSpreadsheet().getId(), 'TRXIO', quo);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Order list').getRange("A2").setValue(jamero[0])
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Order list').getRange("B2").setValue(camero[0])
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Order list').getRange("C2").setValue(gamero[0])
  })
  }else{
    return "idiot";
  }
}

function obtainListofCheckedwithoutStock(){
  var ss = 'Hardware order';
  var id = '1-YBuCQ7bRuJbE3eiP8RmkfXYSaGZqQHhowRsBEIb-5o';
  var qu = 'SELECT D WHERE F = TRUE AND I = FALSE';
  var data = queryASpreadsheet(id, ss, qu);
  return data;
}

function simplystrsing(){
  var ss = 'TRXIO';
  var id = '1-YBuCQ7bRuJbE3eiP8RmkfXYSaGZqQHhowRsBEIb-5o';
  var name = 'Middle Atlantic UFAF-1';
  var f = "'"+name+"'";
  var qu = "SELECT K WHERE J MATCHES "+f;
  var data = queryASpreadsheet(SpreadsheetApp.getActiveSpreadsheet().getId(), 'TRXIO',qu);
  //= queryASpreadsheet(id, ss, qu);
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
    trx.getRange("T"+index).setValue(ger[i]);
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
