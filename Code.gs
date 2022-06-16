function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('dataform')
    .setTitle('Data Validation');
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showSidebar(html);
}

function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .createMenu('Realm Custom Scripts')
    .addItem('Show Estimator sidebar', 'showSidebar')
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

function checkcheckbox(){
  var ss = SpreadsheetApp.getActiveSheet(), gg;
  // if(range.getValues()[0].indexOf("Order Request")!=-1){
  //   var eF = getLetter((range.getValues()[0].indexOf("Order Request")+1));
    // gg=ss.getRange(eF+"1"+":"+eF);
  gg=ss.getRange("F2:F");
  var ge = gg.map((c,i)=>c===true? `F${i+1}`:'').filter(c=>c!='').join(', ');
  for(let i =0;i<ge.length;i++){
    var number = ge[i].replace(/\D/g,'');
    if(ss.getRange("L"+number).getValue()===""){
      //get trxio sheet, check if the thing is in it
      var trix = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TRXIO")
      if(searchTrxio(trix, "D2:D", ss.getRange("D"+number).getValue())!==-1){
        var newOne = SpreadsheetApp.getActiveSpreadsheet();
        let datto = new Date();
        let templateSheet = ss.getSheetByName("Order List");
        newOne.insertSheet('Order List of'+datto, 1, {template: templateSheet});
        let trix = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TRXIO")
        let refNo = (searchTrxio(trix, "D2:D", ss.getRange("D"+number).getValue()))+2;
        let loc = trix.getRange("C"+refNo).getValue();
        let qty = trix.getRange("E"+refNo).getValue();
        let item = trix.getRange("F"+refNo).getValue();

        let ware = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Order List")
        newOne.getRange("A"+getLastDataRow(ware)).setValue(loc);
        newOne.getRange("A"+getLastDataRow(ware)).setValue(qty);
        newOne.getRange("A"+getLastDataRow(ware)).setValue(item);
        ss.getRange("L"+number).setValue(1);
      }
    }
  }
}

function testoo(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var gg = ss.getRange("F2:F").getValues().flat();
  var ge = gg.map((c,i)=>c===true? `F${i+1}`:'').filter(c=>c!='');
  for(let i= 0; i<ge.length;i++){
    let z = 2;
    ss.getRange("M"+z).setValue(ge[i]);
    z+=1;
  }
}

//range = "D2:D"
//ss1 
function searchTrxio(ss, atrange, svalue){
  var range1 = ss.getRange(atrange).getValues();
  var dataList = range1.join("ღ").split("ღ");
  var index = dataList.indexOf(svalue);
  if(index !==-1){
    return index;
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

function joj(){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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

function onEdit(event) {
  var ss = SpreadsheetApp.getActiveSheet();
  var me = Session.getActiveUser();
  if (event.range.isChecked()){
    var stonk = nextLetter(event.range.getA1Notation()[0]);
    var ston = event.range.getA1Notation().replace(/\D/g,'');
    var stonko = nextLetter(stonk);
    ss.getRange(stonk+ston).setValue(new Date());
    ss.getRange(stonko+ston).setValue(Session.getEffectiveUser().getUsername());
    var p = ss.getRange(stonk+ston+":"+stonko+ston).protect();
    p.addEditor(me);
    p.removeEditors(p.getEditors());
    if(p.canDomainEdit()){
      p.setDomainEdit(false);
    }
  } 
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
