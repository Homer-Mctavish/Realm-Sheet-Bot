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
  // var ge = gg.map((c,i)=>c==true? `F${i+2}`:'').filter(c=>c!='');
function checkcheckbox(){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(), gg;
  gg=ss.getRange("F2:F").getValues().flat();
  var ge = gg.map((c,i)=>{if(c===true){return i+2;}else{return '';}}).filter(c=>c!='');
  if(ge.length!==0){
    var newOne = SpreadsheetApp.getActiveSpreadsheet();
    let datto = new Date();
    let templateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Order List");
    newOne.insertSheet('Order List of '+datto, 10, {template: templateSheet});
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(newOne.getSheetByName("Hardware order"));
    ge.forEach(number=>{
      if(ss.getRange("L"+number).getValue()===""){
        var trix = newOne.getSheetByName("TRXIO");
        var gero =  searchTrxio(trix, "E2:E", ss.getRange("D"+number).getValue());
        if(gero!==-1){
          let refNo =gero+2;
          let loc = trix.getRange("C"+refNo).getValue();
          let qty = trix.getRange("O"+refNo).getValue();
          let item = trix.getRange("E"+refNo).getValue();

          newOne.getSheetByName('Order List of '+datto).getRange("A"+getLastDataRow(newOne)).setValue(loc);
          newOne.getSheetByName('Order List of '+datto).getRange("B"+getLastDataRow(newOne)).setValue(qty);
          newOne.getSheetByName('Order List of '+datto).getRange("C"+getLastDataRow(newOne)).setValue(item);
          ss.getRange("L"+number).setValue(1);
        }else{
          return;
        }
      }else{
        return;
      }
    });
  }else{
    return;
  }
}

function testoo(){
  // searchTrxio(trix, "E2:E",)
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(), gg;
  gg=ss.getRange("F2:F").getValues().flat();
  var ge = gg.map((c,i)=>{
    if(c===true){
      return i+2;
    }else{
      return '';
    }
  }).filter(c=>c!='');
  return ge;
}

//range = "D2:D"
//ss1 
function searchTrxio(ss, atrange, svalue){
  var range1 = ss.getRange(atrange).getValues();
  var dataList = range1.join("ღ").split("ღ");
  var index = dataList.indexOf(svalue);
  if(index !==-1){
    return index;
  }else{
    return -1;
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

// SAM on Edit 
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
  //  else if(event.range.isChecked() == false) {
  //   var stonk = nextLetter(event.range.getA1Notation()[0]);
  //   var ston = event.range.getA1Notation().replace(/\D/g,'');
  //   var stonko = nextLetter(stonk);
  //   ss.getRange(stonk+ston).setValue("");
  //   ss.getRange(stonko+ston).setValue("");

  // } 


}

function buton() {
  const  ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.sort(1);
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
