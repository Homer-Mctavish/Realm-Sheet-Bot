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
}

function joj(){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range1 = ss.getRange("C2:C").getValues();
  // ss.getRange("H2").setValue(range1[1][0]);
  var range2 = ss.getRange("D2:D").getValues();
  var dataList = range2.join("ღ").split("ღ");
  var index = dataList.indexOf("");
  var indices = [];
  while (index != -1) {
    indices.push(index);
    index = dataList.indexOf("", index + 1);
  }
  indices.forEach(idx=>{
    let g = 2;
    if(range1[idx][0] != range2[idx][0]){
      ss.getRange("H"+g).setValue(range1[idx][0]);
      g += 1;
    }
  });
}

function onEdit(event) {
  var ss = SpreadsheetApp.getActiveSheet();
  var me = Session.getActiveUser();
  var stonk = nextLetter(event.range.getA1Notation()[0]);
  var ston = event.range.getA1Notation().replace(/\D/g,'');
  var stonko = nextLetter(stonk);
  if (event.range.isChecked()){
    // ss.getRange(stonk+ston).setValue(new Date());
    // ss.getRange(stonko+ston).setValue(Session.getEffectiveUser().getUsername());
    var p = ss.getRange(stonk+ston+":"+stonko+ston).protect();
    p.addEditor(me);
    p.removeEditors(p.getEditors());
    if(p.canDomainEdit()){
      p.setDomainEdit(false);
    }
    ss.getRange(stonk+ston).setValue(p.getEditors()[0]);
  } 
  else if (!event.range.isChecked()&&p.getEditors){
    p.remove();
    ss.getRange(stonk+ston).setValue("");
    ss.getRange(stonko+ston).setValue("");
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