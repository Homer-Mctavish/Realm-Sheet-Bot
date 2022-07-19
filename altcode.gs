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
  
  for(var i=0; i<sheetLastRow; i++){
   
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

function runCreatePullSchedule() {
  // lets delete anything that was in the pull list first.
   
  deleteAllRows();

  let oldSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Old extract");
  let newSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("New extract");
  let newAdds = [];
  let toStrike = [];
  if(newSheet){
    let newOne = newSheet.getRange("A2:E"+newSheet.getLastRow()).getValues();
    let oldOne = oldSheet.getRange("A2:E"+oldSheet.getLastRow()).getValues();
        let stringOfNew = newOne.map(x => x.toString());
      oldOne.forEach(army =>{
        if(stringOfNew.indexOf(army.toString())===-1){
          toStrike.push(army);
        }
      });
      //finds in the new sheet things missing from the old
      let stringOfOld = oldOne.map(x => x.toString());
      newOne.forEach(lbo =>{
        if(stringOfOld.indexOf(lbo.toString())===-1){
          newAdds.push(lbo);
        }
      });
  }
  
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
    //var dataImportRoomNames = dataImportSheet.getRange("G3:" + "H" + dataImportLastRow).getValues();
    var originName = dataImportSheet.getRange("g3").getValue();
    var originRoomNum = dataImportSheet.getRange("h3").getValue();
    var dataSetLastRow = dataSetSheet.getLastRow() + 1;
    var dataSetLastColumn = dataSetSheet.getLastColumn() + 1;
    var dataSetValues = dataSetSheet.getRange("A2:" + "Z" + dataSetLastRow).getValues();
    var insertValues = [];
    var striken = [];
    var addition = [];
    const reference = dataImportValues.flat().map(x =>x.toString());
    //return reference
  
//DOCUMENT THIS BETTER THIS NOT WORKING RIGHT. TV appears out of nowhere for some reason???
  // we are going to loop through the "Data Import" sheet  


  for (var i = 0; i < (dataImportLastRow - 1); i++) {


          if (newAdds.length !== 0 && i<newAdds.length){
            var dataImportPullType2 = newAdds[i][1];
            var dataImportTagNumber2 = newAdds[i][4];
          }

          if(toStrike.length !== 0 && i<toStrike.length){
            var dataImportPullType1 = toStrike[i][1];
            var dataImportTagNumber1 = toStrike[i][4];
          }
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
            if(dataImportPullType1 === dataSetPullType && reference.indexOf(dataImportPullType1)===(reference.indexOf(dataImportTagNumber1)+1)){
              var alphaNes = '';
              for (var iii = 2; iii < (dataSetLastColumn - 1); iii++){
                    if (dataSetValues[ii][iii]) {
                        //Wire #	Wire Type	Wire Origin	Wire Destination	Comments
                      //new order should be Origin, Origin Room #, Destination, Destination Room Number, Destination Description, Cable Number, Wire Type
                        alphaNes = nextString(alphaNes);
                        let destinationDesc = dataSetValues[ii][1];
                        let wireCategory = dataSetPullType;
                        let wireNumber = dataImportTagNumber1 + alphaNes;
                        let wireType = dataSetValues[ii][iii];
                        let wireComment =  dataSetValues[ii][13]
                        Logger.log(wireCategory + "-" + wireNumber + " " + wireType);
                        striken.push([originName,originRoomNum,destinatainName, destinationRoomNumber, destinationDesc, wireCategory + "-" + wireNumber, wireType, wireComment]);
                    } 
              }  
            }
            
            else if(dataImportPullType2 === dataSetPullType && reference.indexOf(dataImportPullType2)===(reference.indexOf(dataImportTagNumber2)+1)){
                var alphaOes = '';
              for (var iii = 2; iii < (dataSetLastColumn - 1); iii++){
                    if (dataSetValues[ii][iii]) {
                        //Wire #	Wire Type	Wire Origin	Wire Destination	Comments
                      //new order should be Origin, Origin Room #, Destination, Destination Room Number, Destination Description, Cable Number, Wire Type
                        alphaOes = nextString(alphaOes);
                        let destinationDesc = dataSetValues[ii][1];
                        let wireCategory = dataSetPullType;
                        let wireNumber = dataImportTagNumber2 + alphaOes;
                        let wireType = dataSetValues[ii][iii];
                        let wireComment =  dataSetValues[ii][13]
                        Logger.log(wireCategory + "-" + wireNumber + " " + wireType);
                        addition.push([originName,originRoomNum,destinatainName, destinationRoomNumber, destinationDesc, wireCategory + "-" + wireNumber, wireType, wireComment]);
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
                        insertValues.push([originName,originRoomNum,destinatainName, destinationRoomNumber, destinationDesc, wireCategory + "-" + wireNumber, wireType, wireComment]);
                    }  
                }

            } 
            else {
                       
             // app.alert("Did Not Find").CLOSE;
            }             
        }
    }
        var strikeYerOuttaHere = pullScheduleSheet.getRange(pullScheduleSheet.getLastRow()+1, 1, striken.length, insertValues[0].length);
    strikeYerOuttaHere.setValues(striken);
    let totalLength = insertValues.length+striken.length+addition.length;
    var range = pullScheduleSheet.getRange(pullScheduleSheet.getLastRow()+1, 1, insertValues.length, insertValues[0].length);
    var changeRange = pullScheduleSheet.getRange(pullScheduleSheet.getLastRow()+1,1,totalLength,pullScheduleSheet.getLastColumn());
    range.setValues(insertValues);
    try{
    const textStyle = SpreadsheetApp.newTextStyle().setStrikethrough(true).build();
    const richTextValues = strikeYerOuttaHere.getRichTextValues();
    strikeYerOuttaHere.setRichTextValues([[textStyle]]);
    var mathemagical = pullScheduleSheet.getRange("A9:G"+addition.length);
    mathemagical.setValues(addition);
    }catch(err){
      return "Oh shit,"+err.message;
    }
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

  function removeSpaces(k){
  return k!== '';
}

  function strikeIn(textsForStrikethrough, sheetName) {
  // const textsForStrikethrough = ["TBD"];  
  // const sheetName = "Pull Schedule";  
  const date = new Date();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const range = sheet.getDataRange();
  const dataRange = sheet.getRange("A9:I"+sheet.getLastRow());
  const modify = dataRange.getValues().filter(x => removeSpaces(x)).reduce((ar, e, r) => {
        if(Number(textsForStrikethrough[2])===Number(e[3]) && textsForStrikethrough[1].toString()===e[5].split("-")[0] && textsForStrikethrough[3].toString()===e[5].split(".")[1].replace(/\D/g,'')){
          ar.push({col:0, row: r})
          ar.push({col:1, row: r})
          ar.push({col:2, row: r})
          ar.push({col:3, row: r})
          ar.push({col:4, row: r})
          ar.push({col:5, row: r})
          ar.push({col:6, row: r})}
        else if(textsForStrikethrough[2].toString()===e[3] && textsForStrikethrough[1].toString()===e[5].split("-")[0] && textsForStrikethrough[3].toString()===e[5].split(".")[1].replace(/\D/g,'')){
          ar.push({col:0, row: r})
          ar.push({col:1, row: r})
          ar.push({col:2, row: r})
          ar.push({col:3, row: r})
          ar.push({col:4, row: r})
          ar.push({col:5, row: r})
          ar.push({col:6, row: r})}
        // else{
        //   SpreadsheetApp.getActiveSpreadsheet().toast(textsForStrikethrough[2]+' is not here.');
        // }
    return ar;
  }, []);
  const textStyle = SpreadsheetApp.newTextStyle().setStrikethrough(true).build();
  const richTextValues = range.getRichTextValues();
  for(i in modify){
    let row = modify[i].row;
    let col = modify[i].col;
    richTextValues[row][col]=richTextValues[row][col].copy().setTextStyle(textStyle).build();
  }
  range.setRichTextValues(richTextValues);
  for (i in modify){
    let row = modify[i].row+1;
    sheet.getRange("I"+row).setValue(date.toDateString());
  }
}

function reducer(rope, neck){
  return rope!== neck;
}

function theGreatFilter(rod, ring){
  rod = rod.map(x =>x.toString());
  ring = ring.map(x => x.toString());
  rod.filter(x => reducer(x, ))
}

function getIntersection(setA, setB) {
  const intersection = [
    [...setA].filter(element => setB.has(element))
  ];

  return intersection;
}

function reValue(twodimensionalArray, oneDimensionalValues){

}

function compareContrast(newOne, oldOne){
  const date = new Date();
  const puller = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pull Schedule");
  const dater = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Set");
  const balues= dater.getRange("A2:A").getValues().flat();
  let toStrike = [];
  let newAdds = [];
  //finds from the old sheet missing things in the new sheet
      let stringOfNew = newOne.map(x => x.toString());
    oldOne.forEach(army =>{
      if(stringOfNew.indexOf(army.toString())===-1){
        toStrike.push(army);
      }
    });
    //finds in the new sheet things missing from the old
    let stringOfOld = oldOne.map(x => x.toString());
    newOne.forEach(lbo =>{
      if(stringOfOld.indexOf(lbo.toString())===-1){
        newAdds.push(lbo);
      }
    });
    //[...new Set(listName)]
    //make a set of both, find the two sets intersection and transform the strings back into arrays using JSON.parse(services)

    const dataRange = puller.getRange("A9:I"+puller.getLastRow()).getValues();
    let identifiers = [];
    for(i of dataRange){
      let bob = i[5].toString();
      let jane = bob.split('-')[0];
      let larry = bob.split('-')[1];
      let nog = larry.split(/(?=[A-Z])/)[0];
      identifiers.push([jane, nog])
    }
    let mydentifiers = [];
    for(i of toStrike){
      mydentifiers.push([i[1].toString(), i[4].toString()]);
    }
    const bagration = identifiers.map(x =>x.toString());
    const damnation = mydentifiers.map(x =>x.toString());
    const dataSet = new Set(bagration);
    const toSet = new Set(damnation);
    const jg = getIntersection(toSet, dataSet);
    const seto =JSON.parse(jg);
    return seto;
    return getIntersection(dataSet, toSet);
    SpreadsheetApp.getActiveSpreadsheet().toast('Striking out deleted items...');
    let bobert = theGreatFilter(toStrike, )
    for(datablock in toStrike){
      strikeIn(toStrike[datablock], "Pull Schedule");
    }


    SpreadsheetApp.getActiveSpreadsheet().toast('Appending added items to Pull Schedule...');
    var lastOne =  puller.getLastRow()+1;
    var alphaMess = '';
    for(counter in newAdds){
      let index = balues.indexOf(newAdds[counter][1])+2;
      let pulltype = dater.getRange("B"+index).getValue();
      let brumpo = "'"+newAdds[counter][2]+"'"; 
      let zumpo = "'"+newAdds[counter][1]+"'";
      let napb = queryASpreadsheet2("1RFZ3lJyqch9wf2pEMIGagxVOp8AvInuoPtVtnppUjW0", "Room Names and numbers", "SELECT B WHERE A MATCHES "+brumpo);
      let kapb = queryASpreadsheet2("1RFZ3lJyqch9wf2pEMIGagxVOp8AvInuoPtVtnppUjW0", "Data Set", "SELECT C WHERE A MATCHES "+zumpo);
      alphaMess = nextString(alphaMess);
      puller.getRange("A"+lastOne).setValue(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Import").getRange("G3").getValue());
      puller.getRange("B"+lastOne).setValue(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Import").getRange("H3").getValue());
      puller.getRange("E"+lastOne).setValue(pulltype);
      puller.getRange("D"+lastOne).setValue(newAdds[counter][2]);
      puller.getRange("C"+lastOne).setValue(napb[0][0]);
      puller.getRange("G"+lastOne).setValue(kapb[0][0]);
      puller.getRange("F"+lastOne).setValue(newAdds[counter][1]+'-'+newAdds[counter][4]+alphaMess);
      puller.getRange("I"+lastOne).setValue(date.toDateString());
      lastOne = lastOne+1;
    }

  }

function grabvals(){
  //make it detect the name of the first sheet (so that "Summary" is replaced with whatever) and the second sheet ("Summary" but it has (number) where number should ideally be 1 because you delete the other sheet after the comparisons are obtained)
  //remember to add the edge case handlers for when one sheet is longer or shorter than the other
  let bob = conCato();
  let oldSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Old extract");
  let newSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("New extract");
  let newOne = newSheet.getRange("A2:E"+newSheet.getLastRow()).getValues();
  let oldOne = oldSheet.getRange("A2:E"+oldSheet.getLastRow()).getValues();
  return compareContrast(newOne, oldOne);
  SpreadsheetApp.getActiveSpreadsheet().toast('Finished.');
}
