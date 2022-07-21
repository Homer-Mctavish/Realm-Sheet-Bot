//https://mashe.hawksey.info/2018/02/google-apps-script-patterns-writing-rows-of-data-to-google-sheets/


function onOpen() {
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
        .createMenu('Realm Custom Scripts')
        .addItem('Create Pull Schedule', 'runCreatePullSchedule')
        .addItem('Set Row Colors & Sort', 'setRowColors')
        .addItem('Speaker Verification','createSpeakerVerification')
        .addItem('Delete Rows','deleteAllRows')
        .addItem('Show Pull sidebar', 'showSidebar')
        .addToUi();
}

function showSidebar() {
  let html = HtmlService.createHtmlOutputFromFile('testSheet')
    .setTitle('Pull Schedule Automata');
    SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showSidebar(html);
}


function onEdit(event) {
     var sheetName = 'Speaker Verification',
         watchCol = [1], 
         stampCol = [9],
         userCol = [8],
         ind = watchCol.indexOf(event.range.columnStart);
     if (event.source.getActiveSheet()
         .getName() !== sheetName ||  ind == -1 || event.range.rowStart < 2) return;

    checkedCell = event.range;
    if (checkedCell.isChecked()) {
    event.source.getActiveSheet()
             .getRange(event.range.rowStart, stampCol[ind])
             .setValue(event.value ? new Date() : null);
    event.source.getActiveSheet()
             .getRange(event.range.rowStart, userCol[ind])
              .setValue(Session.getEffectiveUser().getUsername());

    } else if (!checkedCell.isChecked()){
    event.source.getActiveSheet()
             .getRange(event.range.rowStart, stampCol[ind])
             .setValue(null);
    event.source.getActiveSheet()
             .getRange(event.range.rowStart, userCol[ind])
              .setValue(null);
    }

 }

function createSpeakerVerification(){
req = "=query('Pull Schedule'!C9:G, \"select * where G = '16/4 SPEAKER WIRE'\")";
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Speaker Verification');

queryCell = sheet.getRange(2,3);
queryCell.setValue(req);

sheetData = sheet.getDataRange().getValues();

destination = sheet.getRange(1,1,sheetData.length,sheetData[0].length);
destination.setValues(sheetData);

}


function deleteAllRows(){

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pull Schedule");
  var rowCount = sheet.getMaxRows();
  Logger.log(rowCount);
  if(rowCount >9){
  sheet.deleteRows(9, rowCount-9);
  }
  sheet.getRange("A9:I9").clear();
}

const deepGet = (obj, keys) =>
  keys.reduce(
    (xs, x) => (xs && xs[x] !== null && xs[x] !== undefined ? xs[x] : null),
    obj
  );

function queryASpreadsheet2(sheetId, sheetName, queryString) {
 var url = 'https://docs.google.com/spreadsheets/d/'+sheetId+'/gviz/tq?'+
            'sheet='+sheetName+
            '&tqx=out:csv' +
            '&tq=' + encodeURIComponent(queryString);
  var params = {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  };
  let csvData   = UrlFetchApp.fetch(url, params);
  let dataTwoD  = Utilities.parseCsv(csvData);// array of the format [[a, b, c], [d, e, f]] where [a, b, c] is a row and b is a value
  return dataTwoD;
}

//sheetId, sheetName, queryString
function queryASpreadsheet(sheetId, sheetName, queryString) {
 let url = 'https://docs.google.com/spreadsheets/d/'+sheetId+'/gviz/tq?'+
            'sheet='+sheetName+
            '&tq=' + encodeURIComponent(queryString);
  let params = {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  };
  let ret  = UrlFetchApp.fetch(url, params).getContentText();
  let k = JSON.parse(ret.replace("/*O_o*/", "").replace("google.visualization.Query.setResponse(", "").slice(0, -2));
  let depp = deepGet(k, ['table','rows']);
  let arr = [];
  depp.forEach(column=>{
    arr.push(JSON.stringify(column['c'][0].v))
  });
  return arr;
}

function removeZeros(jobob){
  return jobob.replace(/^0+/, '');
}

function hasNumber(myString) {
  return /\d/.test(myString);
}

function vLooku(){
  let darto = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Import");
  let cellu = darto.getRange("F3");
  cellu.setFormula("=VLOOKUP(E3,'Room Names and numbers'!A1:B"+darto.getLastRow()+",2)");
  let Avals = darto.getRange("E1:E").getValues();
  let Alast = Avals.filter(String).length;
  let ranje = darto.getRange("F3:F"+Alast);
  cellu.copyTo(ranje, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
}

function conCato(){
    let oldSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Old extract");
    let newSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("New extract");
    let cell1 = oldSheet.getRange("E2");
    let cell2 = newSheet.getRange("E2");
    cell1.setFormula('=CONCATENATE(C2,".",D2)');
    cell2.setFormula('=CONCATENATE(C2,".",D2)');
    let range1 = oldSheet.getRange("E2:E"+oldSheet.getLastRow());
    let range2 = newSheet.getRange("E2:E"+newSheet.getLastRow());
    cell1.copyTo(range1, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
    cell2.copyTo(range2, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  }
function processXLSsheet(){
  let oldSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Old extract");
  let newSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("New extract");
  let newOne = newSheet.getRange("A2:E"+newSheet.getLastRow()).getValues();
  let oldOne = oldSheet.getRange("A2:E"+oldSheet.getLastRow()).getValues();
  let toStrike = [];
  //let newAdds = [];
  //finds from the old sheet missing things in the new sheet
      let stringOfNew = newOne.map(x => x.toString());
    oldOne.forEach(army =>{
      if(stringOfNew.indexOf(army.toString())===-1){
        toStrike.push(army);
      }
    });
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Old extract");
  var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Import");
  var sheetLastRow = sheet.getLastRow();
  var dataValues1 = sheet.getRange(2,3,sheetLastRow).getValues();
  var dataValues2 = sheet.getRange(2,4,sheetLastRow).getValues();
  var dataValues3 = sheet.getRange(2,2,sheetLastRow).getValues();
  var combined = [];
  var pullTypes = [];
  
  for(let i=0; i<sheetLastRow; i++){
   
    combined[i] = [dataValues1[i][0]+"."+ dataValues2[i][0]];
    pullTypes[i] = [dataValues3[i][0]];
    
  }
  var rowcount = combined.length;
  sheet2.getRange(3,1,rowcount).setValues(combined);
  sheet2.getRange(3,2,rowcount).setValues(pullTypes);
  let loength= toStrike.length+2;
  sheet2.getRange("N3:R"+loength).setValues(toStrike);
  vLooku();
}
  

function sortRows(){
  var sheet =  SpreadsheetApp.getActiveSheet();
  sheet.setFrozenRows(8);
  var sheetLastRow = sheet.getLastRow();
  var sortrange = sheet.getRange("A9:" + sheetLastRow);
  sortrange.sort([{column: 4, ascending: true}, {column: 6, ascending: true}])
}

function fokault(tho, fog){
  fog.forEach(jam =>{
    let pog = jam[1];
    let mog = jam[4];
    for (j of tho){
      let b = j
      let zog = j.indexOf(pog);
      let log = j.indexOf(Number(mog));
      if(zog===log+1){
        let bobo = tho.indexOf(j);
        tho.splice(bobo, 1);
      };
    };
  });
  return tho;
}


function runCreatePullSchedule() {
  // lets delete anything that was in the pull list first.
   
  deleteAllRows();

  let oldSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Old extract");
  let newSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("New extract");
  let newAdds = [];
  let toStrike = [];
  let newOne = newSheet.getRange("A2:E"+newSheet.getLastRow()).getValues();
  let oldOne = oldSheet.getRange("A2:E"+oldSheet.getLastRow()).getValues();
      if(newSheet){
        let stringOfNew = newOne.map(x => x.toString());
      oldOne.forEach(army =>{
        if(stringOfNew.indexOf(army.toString())===-1){
          toStrike.push(army);
        }
      });
      }
  
    //we need to loop through a sheet that has tag number and wire type. Then we will add to the current Pull Shedule sheet the wire number, Wire type and wire orgin/destination.
    //we will extract the room names wire labels and type from Data Import Sheet. Need to figure out easy way for data import sheet to populate names of rooms.
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    const pullScheduleSheet = ss.getSheetByName("Pull Schedule");
    const dataSetSheet = ss.getSheetByName("Data Set");
    const dataImportSheet = ss.getSheetByName("Data Import");

    //Lets get our data from Data import Sheet and Data Set sheet 
    let dataImportLastRow = dataImportSheet.getLastRow() + 1;


    let importToClean = dataImportSheet.getRange("A2:" + "C" + dataImportLastRow).getValues();
    let dataImportValues =fokault(importToClean, toStrike);

    
    const originName = dataImportSheet.getRange("g3").getValue();
    const originRoomNum = dataImportSheet.getRange("h3").getValue();
    let dataSetLastRow = dataSetSheet.getLastRow() + 1;
    let dataSetLastColumn = dataSetSheet.getLastColumn() + 1;
    let dataSetValues = dataSetSheet.getRange("A2:" + "Z" + dataSetLastRow).getValues();
    let insertValues = [];
    let striken = [];
    let addition = [];
    const reference = dataImportValues.flat().map(x =>x.toString());
    if(newSheet){
      //finds in the new sheet things missing from the old
      let stringOfOld = oldOne.map(x => x.toString());
      newOne.forEach(lbo =>{
        if(stringOfOld.indexOf(lbo.toString())===-1){
          newAdds.push(lbo);
        }
      });

      for(let h = 0; h<toStrike.length;h++){
            
            if(toStrike.length !== 0 && h<toStrike.length){
                var dataImportPullType1 = toStrike[h][1];
                var dataImportTagNumber1 = toStrike[h][4]; 
            }

          var dataImportTagNumber = dataImportValues[h+1][0];
          var dataImportPullType =  dataImportValues[h][1];
          var destinatainName =  dataImportValues[h+1][2]; 
          var dataImportTagNumberSplit =  dataImportTagNumber.toString();
    
          var destinationRoomNumber = dataImportTagNumberSplit.split(".")[0];
          ///Logger.log(destinationRoomNumber);
        // now lets loop through "Data Set" to match up column B in Data Import sheet (TV, SPK etc) with Column A in "Data Set" Sheet
        for (let hh = 0; hh < (dataSetLastRow - 1); hh++) {
            var dataSetPullType1 = dataSetValues[hh][0];
            //If we find a match we can move forward.
            if(dataImportPullType1 === dataSetPullType1){
              var alphaNes = '';
              for (let hhh = 2; hhh < (dataSetLastColumn - 1); hhh++){
                    let bobatea = dataSetValues[hh][hhh];
                    if (dataSetValues[hh][hhh]) {
                      //Wire #	Wire Type	Wire Origin	Wire Destination	Comments
                      //new order should be Origin, Origin Room #, Destination, Destination Room Number, Destination Description, Cable Number, Wire Type
                        alphaNes = nextString(alphaNes);
                        let destinationDesc = dataSetValues[hh][1];
                        let wireCategory = dataSetPullType1;
                        let wireNumber = dataImportTagNumber1 + alphaNes;
                        let wireType = dataSetValues[hh][hhh];
                        let wireComment =  dataSetValues[hh][13];
                        let bobi = new Date();
                        let vee = bobi.toDateString().replaceAll(" ", "/");
                        //Logger.log(wireCategory + "-" + wireNumber + " " + wireType+"strike");
                        striken.push([originName,originRoomNum,destinatainName, destinationRoomNumber, destinationDesc, wireCategory + "-" + wireNumber, wireType, wireComment, vee]);
                    } 
              }  
            }
        }

      }
  }
  // we are going to loop through the "Data Import" sheet  

  for (var i = 0; i < (dataImportValues.length - 1); i++) {
          var dataImportTagNumber = dataImportValues[i+1][0];
          var dataImportPullType =  dataImportValues[i][1];
          var destinatainName =  dataImportValues[i+1][2]; 
          var dataImportTagNumberSplit =  dataImportTagNumber.toString();
    
          var destinationRoomNumber = dataImportTagNumberSplit.split(".")[0];
          //Logger.log(destinationRoomNumber);
        // now lets loop through "Data Set" to match up column B in Data Import sheet (TV, SPK etc) with Column A in "Data Set" Sheet
        for (var ii = 0; ii < (dataSetLastRow - 1); ii++) {

            if (newAdds.length !== 0 && i<newAdds.length){
                var dataImportPullType2 = newAdds[i][1];
                var dataImportTagNumber2 = newAdds[i][4]; 
            }

            var dataSetPullType = dataSetValues[ii][0];
            var dataSetPullType2 = dataSetValues[ii][0];
            
            if(dataImportPullType2 === dataSetPullType2 && reference.indexOf(dataImportTagNumber2.toString())===-1){
                var alphaOes = '';
              for (var iii = 2; iii < (dataSetLastColumn - 1); iii++){
                    if (dataSetValues[ii][iii]) {
                        alphaOes = nextString(alphaOes);
                        let destinationDesc = dataSetValues[ii][1];
                        let wireCategory = dataSetPullType2;
                        let wireNumber = dataImportTagNumber2 + alphaOes;
                        if(newAdds[i+1]!=null){
                          dataImportTagNumber2 = newAdds[i+1][4];
                        }else{
                          dataImportTagNumber2 = dataImportTagNumber;
                        }

                        let wireType = dataSetValues[ii][iii];
                        let wireComment =  dataSetValues[ii][13];
                        let bobi = new Date();
                        let vee = bobi.toDateString().replaceAll(" ", "/");
                        //Logger.log(wireCategory + "-" + wireNumber + " " + wireType+"new");
                        addition.push([originName,originRoomNum,destinatainName, destinationRoomNumber, destinationDesc, wireCategory + "-" + wireNumber, wireType, wireComment, vee]);
                    } 
              }
            }else if (dataImportPullType === dataSetPullType) {
         
                //Now we will loop through the columns of "Data Set" We need to skip B because that has our wire category i.i Flat Panel, Wireless Access Point 
                var alphaDes = '';
                for (var iii = 2; iii < (dataSetLastColumn - 1); iii++) {
                    //make sure cell isn't empty before moving on
                    if (dataSetValues[ii][iii]) {
                        //Wire #	Wire Type	Wire Origin	Wire Destination	Comments
                      //new order should be Origin, Origin Room #, Destination, Destination Room Number, Destination Description, Cable Number, Wire Type
                        alphaDes = nextString(alphaDes);
                        let destinationDesc = dataSetValues[ii][1];
                        let wireCategory = dataSetPullType;
                        let wireNumber = dataImportTagNumber + alphaDes;
                        let wireType = dataSetValues[ii][iii];
                        let wireComment =  dataSetValues[ii][13]
                        Logger.log(wireCategory + "-" + wireNumber + " " + wireType);
                        insertValues.push([originName,originRoomNum,destinatainName, destinationRoomNumber, destinationDesc, wireCategory + "-" + wireNumber, wireType, wireComment, ""]);
                    }  
                }

            } 
            else {
                       
             // app.alert("Did Not Find").CLOSE;
            }             
        }
    }//here it is
    if(newSheet){
    addition = Array.from(new Set(addition.map(JSON.stringify)), JSON.parse);
    let mathemagical = pullScheduleSheet.getRange(pullScheduleSheet.getLastRow()+1, 1, addition.length, addition[0].length);
    bong= pullScheduleSheet.getLastRow()+1;
    mathemagical.setValues(addition);

            let strikeYerOuttaHere = pullScheduleSheet.getRange(pullScheduleSheet.getLastRow()+1, 1, striken.length, striken[0].length);

    // let bong = bobn(strikeYerOuttaHere);
    strikeYerOuttaHere.setValues(striken);

    let range = pullScheduleSheet.getRange(pullScheduleSheet.getLastRow()+1, 1, insertValues.length, insertValues[0].length);
    range.setValues(insertValues);

    const textStyle = SpreadsheetApp.newTextStyle().setStrikethrough(true).build();
    const richTextValues = strikeYerOuttaHere.getRichTextValues();
    for(let z= 0; z<striken.length; z++){
      for(let o=0; o<striken[0].length; o++){
        richTextValues[z][o]=richTextValues[z][o].copy().setTextStyle(textStyle).build();
      }
    }
    strikeYerOuttaHere.setRichTextValues(richTextValues);
        let changeRange = pullScheduleSheet.getRange("A9:J"+pullScheduleSheet.getLastRow());
    changeRange.setBackgroundRGB(255, 255, 255);
    changeRange.setFontSize(12);
    changeRange.setFontFamily("Share Tech Mono");
    }else{
    let range = pullScheduleSheet.getRange(pullScheduleSheet.getLastRow()+1, 1, insertValues.length, insertValues[0].length);
    range.setValues(insertValues);

    
    let changeRange = pullScheduleSheet.getRange("A9:J"+pullScheduleSheet.getLastRow());
    changeRange.setBackgroundRGB(255, 255, 255);
    changeRange.setFontSize(12);
    changeRange.setFontFamily("Share Tech Mono");
    }
}

function bobn(strik){
  return strik.getA1Notation();
}

//not a working function. doing this outside of scripting now. 
function addRoomNames(){
  
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var app = SpreadsheetApp.getUi();
    var dataImportSheet = ss.getSheetByName("Data Import");
    var dataImportRoomNames = dataImportSheet.getRange("G3:" + "H" + dataImportValues.length).getValues();
  
                //////THIS IS WHERE TROUBLE STARTS we should change 50 to actual number of rooms. Should loop through this first and update cells in dataimport then just take the column
              
              //now lets get the destination name

              for (var di = 0; di < 25; di++) { 
                var roomNumber = dataImportRoomNames[di][0].toString();
                Logger.log("*********");
              Logger.log(roomNumber);
              Logger.log(destinationRoomNumber);
              Logger.log("*********");
                
                if(destinationRoomNumber === roomNumber ){
                  
                  var destinatinName = dataImportRoomNames[di][1];
                  Logger.log("*********");
                  Logger.log("*********");
                  Logger.log("*********");
                  Logger.log(destinatinName);
                  Logger.log("*********");
                  Logger.log("*********");
                  Logger.log("*********");
                  
                  
                } 
              } 
              
              ////IT SHOULD END HERE. THE TROUBLE THAT IS.
  
}


//function that returns the next string in lexicographic order: 'A' -> 'B' -> ... 'Z' -> 'AA' -> 'AB' -> 'AC' -> ... 'AZ' -> 'BA' -> 'BB' -> ... 'ZZ' -> 'AAA' etc.
//https://stackoverflow.com/questions/32157500/increment-alphabet-characters-to-next-character-using-javascript
function nextString(str) {
    if (!str)
        return 'A'; // return 'A' if str is empty or null

    var tail = '';
    var i = str.length - 1;
    var char = str[i];
    // find the index of the first character from the right that is not a 'Z'
    while (char === 'Z' && i > 0) {
        i--;
        char = str[i];
        tail = 'A' + tail; // tail contains a string of 'A'
    }
    if (char === 'Z') // the string was made only of 'Z'
        return 'AA' + tail;
    // increment the character that was not a 'Z'
    return str.slice(0, i) + String.fromCharCode(char.charCodeAt(0) + 1) + tail;
}


function formatText() {
    var range1 = pullScheduleSheet.getRange("C5:E5");
    range1.mergeAcross();
    range1.setHorizontalAlignment("center");
    range1.setVerticalAlignment("middle");
    range1.setBackgroundRGB(169, 169, 169);
    range1.setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    range1.setFontWeight("bold");
    var fontSizes = [
        [44, 46, 48]
    ];

    range1.setFontSizes(fontSizes);
 

}


function setRowColors() {
  sortRows();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var range = sheet.getDataRange();
  
  var lRow = sheet.getLastRow();
  var headerRows = 8;
  var numRows = lRow - headerRows;
  var numCols = sheet.getLastColumn();
  var [rows1d, cols1d] = [numRows, numCols].map(function(num){ 
    return Array.apply([],new Array(num)); //or just `getBackgrounds()` to get a 2d array 
  })
  
  var colors2d = rows1d.map(function(row, i){
    var color = i%2 === 0 ? "#ffffff" : "#efefef";
    return cols1d.map(function(col){
        return color;
    })
  })

  sheet.getRange(headerRows + 1, 1, numRows, numCols).setBackgrounds(colors2d);
  
  setCellColors();
  }


function setCellColors() {  
  var range = SpreadsheetApp.getActiveSheet().getDataRange();
  
  //lets find Lutron and Power in Column G and set background color to yellow and red

    var gi = 0;
  
  // we set every other row white or grey
  for (var i = range.getRow()+7; i < range.getLastRow(); i++) {
    var rowRow = i +1;
    var pullScheduleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pull Schedule");
  //  var pullScheduleSheet = ss.getSheetName("Pull Schedule");
    var pullScheduleLastRow = pullScheduleSheet.getLastRow() + 1;
    var pullScheduleValues = pullScheduleSheet.getRange("G9:" + "G" + pullScheduleLastRow).getValues();

  if (pullScheduleValues[gi][0] == "Lutron QSC" || pullScheduleValues[gi][0] == "Lutron Yellow" || pullScheduleValues[gi][0] == "LUTRON QSC" || pullScheduleValues[gi][0] == "LUTRON YELLOW"){
    pullScheduleSheet.getRange("G"+rowRow).setBackgroundColor('#fff187');
    pullScheduleSheet.getRange("A"+rowRow).setBackgroundColor('#fff187');
    pullScheduleSheet.getRange("B"+rowRow).setBackgroundColor('#fff187');
    }
    
  if (pullScheduleValues[gi][0] == "120V"){
    pullScheduleSheet.getRange("G"+rowRow).setBackgroundColor('#ef3737');
    }
    
   gi = gi + 1;
  Logger.log(gi);
  }
  }

