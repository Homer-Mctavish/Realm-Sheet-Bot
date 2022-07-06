//https://mashe.hawksey.info/2018/02/google-apps-script-patterns-writing-rows-of-data-to-google-sheets/


function onOpen() {
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
        .createMenu('Realm Custom Scripts')
        .addItem('Create Pull Schedule', 'runCreatePullSchedule')
        .addItem('Set Row Colors & Sort', 'setRowColors')
        .addItem('Speaker Verification','createSpeakerVerification')
        .addItem('Delete Rows','deleteAllRows')
        .addToUi();
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

function processXLSsheet(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary");
  var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Import");
  var sheetLastRow = sheet.getLastRow();
  var dataValues1 = sheet.getRange(2,3,sheetLastRow).getValues();
  var dataValues2 = sheet.getRange(2,4,sheetLastRow).getValues();
  var dataValues3 = sheet.getRange(2,2,sheetLastRow).getValues();
  var combined = [];
  var pullTypes = [];
  
  for(var i=0; i<sheetLastRow; i++){
   
    combined[i] = [dataValues1[i][0]+"."+ dataValues2[i][0]];
    pullTypes[i] = [dataValues3[i][0]];
    
  }
  var rowcount = combined.length;
  sheet2.getRange(3,1,rowcount).setValues(combined);
  sheet2.getRange(3,2,rowcount).setValues(pullTypes);
  
}
  

function sortRows(){
  var sheet =  SpreadsheetApp.getActiveSheet();
  sheet.setFrozenRows(8);
  var sheetLastRow = sheet.getLastRow();
  var sortrange = sheet.getRange("A9:" + sheetLastRow);
  sortrange.sort([{column: 4, ascending: true}, {column: 6, ascending: true}])
}

function runCreatePullSchedule() {
  // lets delete anything that was in the pull list first.
   
  deleteAllRows();
  
    //we need to loop through a sheet that has tag number and wire type. Then we will add to the current Pull Shedule sheet the wire number, Wire type and wire orgin/destination.
    //we will extract the room names wire labels and type from Data Import Sheet. Need to figure out easy way for data import sheet to populate names of rooms.
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var app = SpreadsheetApp.getUi();
    var pullScheduleSheet = ss.getSheetByName("Pull Schedule");
    var dataSetSheet = ss.getSheetByName("Data Set");
    var dataImportSheet = ss.getSheetByName("Data Import");

    //Lets get our data from Data import Sheet and Data Set sheet 
    var dataImportLastRow = dataImportSheet.getLastRow() + 1;
    var dataImportValues = dataImportSheet.getRange("A2:" + "C" + dataImportLastRow).getValues();
    var dataImportRoomNames = dataImportSheet.getRange("G3:" + "H" + dataImportLastRow).getValues();
    var originName = dataImportSheet.getRange("g3").getValue();
    var originRoomNum = dataImportSheet.getRange("h3").getValue();
    var dataSetLastRow = dataSetSheet.getLastRow() + 1;
    var dataSetLastColumn = dataSetSheet.getLastColumn() + 1;
    var dataSetValues = dataSetSheet.getRange("A2:" + "Z" + dataSetLastRow).getValues();
    var insertValues = [];
  
//DOCUMENT THIS BETTER THIS NOT WORKING RIGHT. TV appears out of nowhere for some reason???
  // we are going to loop through the "Data Import" sheet  
  for (var i = 0; i < (dataImportLastRow - 1); i++) {

          var dataImportTagNumber = dataImportValues[i][0];
          var dataImportPullType =  dataImportValues[i][1]; 
          var destinatainName =  dataImportValues[i][2]; 
          var dataImportTagNumberSplit =  dataImportTagNumber.toString();
    
          var destinationRoomNumber = dataImportTagNumberSplit.split(".")[0];
          Logger.log(destinationRoomNumber);
        // now lets loop through "Data Set" to match up column B in Data Import sheet (TV, SPK etc) with Column A in "Data Set" Sheet
        for (var ii = 0; ii < (dataSetLastRow - 1); ii++) {
              var dataSetPullType = dataSetValues[ii][0];
             //If we find a match we can move forward. 
            if (dataImportPullType === dataSetPullType) {
         
                //Now we will loop through the columns of "Data Set" We need to skip B because that has our wire category i.i Flat Panel, Wireless Access Point 
                var alphaDes = '';
                for (var iii = 2; iii < (dataSetLastColumn - 1); iii++) {
                    //make sure cell isn't empty before moving on
                    if (dataSetValues[ii][iii]) {
                        //Wire #	Wire Type	Wire Origin	Wire Destination	Comments
                      //new order should be Origin, Origin Room #, Destination, Destination Room Number, Destination Description, Cable Number, Wire Type
                        alphaDes = nextString(alphaDes);
                        var destinationDesc = dataSetValues[ii][1];
                        var wireCategory = dataSetPullType;
                        var wireNumber = dataImportTagNumber + alphaDes;
                        var wireType = dataSetValues[ii][iii];
                        var wireComment =  dataSetValues[ii][13]
                        Logger.log(wireCategory + "-" + wireNumber + " " + wireType);
                        insertValues.push([originName,originRoomNum,destinatainName, destinationRoomNumber, destinationDesc, wireCategory + "-" + wireNumber, wireType, wireComment]);
                    }  
                }

            } else {
                       
             // app.alert("Did Not Find").CLOSE;
}
                     
        }

        
       
    }
    var range = pullScheduleSheet.getRange(pullScheduleSheet.getLastRow()+1, 1, insertValues.length, insertValues[0].length);
    var changeRange = pullScheduleSheet.getRange(pullScheduleSheet.getLastRow()+1,1,insertValues.length,pullScheduleSheet.getLastColumn());
    range.setValues(insertValues);
    changeRange.setBackgroundRGB(255, 255, 255);
    changeRange.setFontSize(12);
    changeRange.setFontFamily("Share Tech Mono");
  

}

//not a working function. doing this outside of scripting now. 
function addRoomNames(){
  
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var app = SpreadsheetApp.getUi();
    var dataImportSheet = ss.getSheetByName("Data Import");
    var dataImportRoomNames = dataImportSheet.getRange("G3:" + "H" + dataImportLastRow).getValues();
  
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

// const deepGet = (obj, keys) =>
//   keys.reduce(
//     (xs, x) => (xs && xs[x] !== null && xs[x] !== undefined ? xs[x] : null),
//     obj
//   );

// //sheetId, sheetName, queryString
// function queryASpreadsheet(sheetId, sheetName, queryString) {
//  var url = 'https://docs.google.com/spreadsheets/d/'+sheetId+'/gviz/tq?'+
//             'sheet='+sheetName+
//             '&tq=' + encodeURIComponent(queryString);
//   var params = {
//     headers: {
//       'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
//     },
//     muteHttpExceptions: true
//   };
//   var ret  = UrlFetchApp.fetch(url, params).getContentText();
//   var k = JSON.parse(ret.replace("/*O_o*/", "").replace("google.visualization.Query.setResponse(", "").slice(0, -2));
//   var depp = deepGet(k, ['table','rows']);
//   var arr = [];
//   depp.forEach(column=>{
//     arr.push(JSON.stringify(column['c'][0].v))
//   });
//   return arr;
// }

//   function queryImport(){
//     const items = queryASpreadsheet("1iNOyqZuLorKOO3qOctOD6QfqJYeuvuXK9I_AkO4hh2o", "Data Import", 'SELECT E WHERE E IS NOT NULL'), gamer = items.map(function(item) {
//     return item.toString();
//     });
//     let final = [];
//     gamer.forEach(item =>{
//       let gobi =queryASpreadsheet("1iNOyqZuLorKOO3qOctOD6QfqJYeuvuXK9I_AkO4hh2o", "Rooms and Numbers", "SELECT B WHERE A MATCHES "+item);
//       final.push(gobi);
//     });
//     return final;
//   };   

// function grabvals(){
//   //make it detect the name of the first sheet (so that "Summary" is replaced with whatever) and the second sheet ("Summary" but it has (number) where number should ideally be 1 because you delete the other sheet after the comparisons are obtained)
//   let old = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary");
//   let ew = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary(1)");
//   old.getDataRange().getValues().flat();
//   ew.getDataRange().getValues().flat();
//   let oldSet = new
// }