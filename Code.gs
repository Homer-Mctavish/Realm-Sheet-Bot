
// https://mashe.hawksey.info/2018/02/google-apps-script-patterns-writing-rows-of-data-to-google-sheets/

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
     let sheetName = 'Speaker Verification',
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
let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Speaker Verification');

queryCell = sheet.getRange(2,3);
queryCell.setValue(req);

sheetData = sheet.getDataRange().getValues();

destination = sheet.getRange(1,1,sheetData.length,sheetData[0].length);
destination.setValues(sheetData);

}


function deleteAllRows(){

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pull Schedule");
  let rowCount = sheet.getMaxRows();
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

  function superV(){//use the findrow function from item import sheet to find row to insert the string of mog[o]
    let vapb = queryASpreadsheet("1iNOyqZuLorKOO3qOctOD6QfqJYeuvuXK9I_AkO4hh2o", "Rooms and Numbers", 'SELECT A WHERE B IS NOT NULL');
    let napb = queryASpreadsheet("1iNOyqZuLorKOO3qOctOD6QfqJYeuvuXK9I_AkO4hh2o", "Rooms and Numbers", 'SELECT B WHERE B IS NOT NULL');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Import");
    const shrt = sheet.getRange("E3:E").getValues().filter(String);
    let mog = [];    
    shrt.forEach(argk =>{
      mog.push(argk.map(x => removeZeros(x.toString()).toString()));
    });
    var o = 0;
    let flatten = mog.flat();
    vapb.forEach(noomber =>{
      if(flatten.indexOf(noomber)!==-1){
        let placement = flatten.indexOf(noomber)+3;
        sheet.getRange("F"+placement).setValue(napb[o].replaceAll("\"",''));
        o = o+1;
      };
    });
  };

function processXLSsheet(){
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary");
  let sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Import");
  let sheetLastRow = sheet.getLastRow();
  let dataValues1 = sheet.getRange(2,3,sheetLastRow).getValues();
  let dataValues2 = sheet.getRange(2,4,sheetLastRow).getValues();
  let dataValues3 = sheet.getRange(2,2,sheetLastRow).getValues();
  let combined = [];
  let pullTypes = [];
  
  for(let i=0; i<sheetLastRow; i++){
   
    combined[i] = [dataValues1[i][0]+"."+ dataValues2[i][0]];
    pullTypes[i] = [dataValues3[i][0]];
    
  }
  let rowcount = combined.length;
  sheet2.getRange(3,1,rowcount).setValues(combined);
  sheet2.getRange(3,2,rowcount).setValues(pullTypes);
  superV()
};


function sortRows(){
  let sheet =  SpreadsheetApp.getActiveSheet();
  sheet.setFrozenRows(8);
  let sheetLastRow = sheet.getLastRow();
  let sortrange = sheet.getRange("A9:" + sheetLastRow);
  sortrange.sort([{column: 4, ascending: true}, {column: 6, ascending: true}])
}

function runCreatePullSchedule() {
  // lets delete anything that was in the pull list first.
   
  deleteAllRows();
  
    //we need to loop through a sheet that has tag number and wire type. Then we will add to the current Pull Shedule sheet the wire number, Wire type and wire orgin/destination.
    //we will extract the room names wire labels and type from Data Import Sheet. Need to figure out easy way for data import sheet to populate names of rooms.
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let app = SpreadsheetApp.getUi();
    let pullScheduleSheet = ss.getSheetByName("Pull Schedule");
    let dataSetSheet = ss.getSheetByName("Data Set");
    let dataImportSheet = ss.getSheetByName("Data Import");

    //Lets get our data from Data import Sheet and Data Set sheet 
    let dataImportLastRow = dataImportSheet.getLastRow() + 1;
    let dataImportValues = dataImportSheet.getRange("A2:" + "C" + dataImportLastRow).getValues();
    let dataImportRoomNames = dataImportSheet.getRange("G3:" + "H" + dataImportLastRow).getValues();
    let originName = dataImportSheet.getRange("g3").getValue();
    let originRoomNum = dataImportSheet.getRange("h3").getValue();
    let dataSetLastRow = dataSetSheet.getLastRow() + 1;
    let dataSetLastColumn = dataSetSheet.getLastColumn() + 1;
    let dataSetValues = dataSetSheet.getRange("A2:" + "Z" + dataSetLastRow).getValues();
    let insertValues = [];
  
//DOCUMENT THIS BETTER THIS NOT WORKING RIGHT. TV appears out of nowhere for some reason???
  // we are going to loop through the "Data Import" sheet  
  for (let i = 0; i < (dataImportLastRow - 1); i++) {

          let dataImportTagNumber = dataImportValues[i][0];
          let dataImportPullType =  dataImportValues[i][1]; 
          let destinatainName =  dataImportValues[i][2]; 
          let dataImportTagNumberSplit =  dataImportTagNumber.toString();
    
          let destinationRoomNumber = dataImportTagNumberSplit.split(".")[0];
          Logger.log(destinationRoomNumber);
        // now lets loop through "Data Set" to match up column B in Data Import sheet (TV, SPK etc) with Column A in "Data Set" Sheet
        for (let ii = 0; ii < (dataSetLastRow - 1); ii++) {
              let dataSetPullType = dataSetValues[ii][0];
             //If we find a match we can move forward. 
            if (dataImportPullType === dataSetPullType) {
         
                //Now we will loop through the columns of "Data Set" We need to skip B because that has our wire category i.i Flat Panel, Wireless Access Point 
                let alphaDes = '';
                for (let iii = 2; iii < (dataSetLastColumn - 1); iii++) {
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

            } else {
                       
             // app.alert("Did Not Find").CLOSE;
              }            
        }
    }
    let range = pullScheduleSheet.getRange(pullScheduleSheet.getLastRow()+1, 1, insertValues.length, insertValues[0].length);
    let changeRange = pullScheduleSheet.getRange(pullScheduleSheet.getLastRow()+1,1,insertValues.length,pullScheduleSheet.getLastColumn());
    range.setValues(insertValues);
    changeRange.setBackgroundRGB(255, 255, 255);
    changeRange.setFontSize(12);
    changeRange.setFontFamily("Share Tech Mono");
}


//function that returns the next string in lexicographic order: 'A' -> 'B' -> ... 'Z' -> 'AA' -> 'AB' -> 'AC' -> ... 'AZ' -> 'BA' -> 'BB' -> ... 'ZZ' -> 'AAA' etc.
//https://stackoverflow.com/questions/32157500/increment-alphabet-characters-to-next-character-using-javascript
function nextString(str) {
    if (!str)
        return 'A'; // return 'A' if str is empty or null

    let tail = '';
    let i = str.length - 1;
    let char = str[i];
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
    let range1 = pullScheduleSheet.getRange("C5:E5");
    range1.mergeAcross();
    range1.setHorizontalAlignment("center");
    range1.setVerticalAlignment("middle");
    range1.setBackgroundRGB(169, 169, 169);
    range1.setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    range1.setFontWeight("bold");
    let fontSizes = [
        [44, 46, 48]
    ];
    range1.setFontSizes(fontSizes);
}


function setRowColors() {
  sortRows();
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheets()[0];
  let range = sheet.getDataRange();
  
  let lRow = sheet.getLastRow();
  let headerRows = 8;
  let numRows = lRow - headerRows;
  let numCols = sheet.getLastColumn();
  let [rows1d, cols1d] = [numRows, numCols].map(function(num){ 
    return Array.apply([],new Array(num)); //or just `getBackgrounds()` to get a 2d array 
  })
  
  let colors2d = rows1d.map(function(row, i){
    let color = i%2 === 0 ? "#ffffff" : "#efefef";
    return cols1d.map(function(col){
        return color;
    })
  })

  sheet.getRange(headerRows + 1, 1, numRows, numCols).setBackgrounds(colors2d);
  
  setCellColors();
  }


function setCellColors() {  
  let range = SpreadsheetApp.getActiveSheet().getDataRange();
  
  //lets find Lutron and Power in Column G and set background color to yellow and red

    let gi = 0;
  
  // we set every other row white or grey
  for (let i = range.getRow()+7; i < range.getLastRow(); i++) {
    let rowRow = i +1;
    let pullScheduleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pull Schedule");
  //  let pullScheduleSheet = ss.getSheetName("Pull Schedule");
    let pullScheduleLastRow = pullScheduleSheet.getLastRow() + 1;
    let pullScheduleValues = pullScheduleSheet.getRange("G9:" + "G" + pullScheduleLastRow).getValues();

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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const range = sheet.getDataRange();
  const modify = range.getValues().reduce((ar, e, r) => {
      textsForStrikethrough.forEach(g => {
        const idx = e.indexOf(g);
        const nidx = e.indexOf(textsForStrikethrough[1]);
        if(idx===nidx){
          let meato = []
          e.filter(function(yourArray, index) {
          if(yourArray === g){
            meato.push(index)
          }
          meato.forEach(meat =>{
            ar.push({col:meat, row:r})
          })
          });
        }else if (idx > -1 && nidx > -1){ ar.push({col: idx, row: r})};
    });
    return ar;
  }, []);
  const textStyle = SpreadsheetApp.newTextStyle().setStrikethrough(true).build();
  const richTextValues = range.getRichTextValues();
  modify.forEach(e =>{
      richTextValues[e.row][e.col]=richTextValues[e.row][e.col].copy().setTextStyle(textStyle).build();
  }); 
  range.setRichTextValues(richTextValues);
}

function compareContrast(newOne, oldOne){
  let puller = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pull Schedule");
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
    let dumb = [];
    toStrike.forEach(argk =>{
      dumb.push(argk.map(x => removeZeros(x.toString()).toString()));
    }); 
    let value =puller.getDataRange().getValues();
    let baloo = [];
    value.forEach(bobert =>{
      baloo.push(bobert.filter(x => removeSpaces(x.toString())));
    }); 
    let formofdata = [];
    baloo.forEach(itoo => {
      if(dumb.indexOf(itoo.toString())!==-1){
        formofdata.push(itoo);
      }
    });


    let newerAdds = [];
    newAdds.forEach(item =>{
      newerAdds.push(item.map(x => removeZeros(x.toString()).toString()));
    });
    dumb.forEach(datablock =>{
      strikeIn(datablock, "Pull Schedule");
    });
    // the part where we add the new stuff from Summary (1) to Pull Schedule
    var lastOne =  puller.getLastRow()+1;
    var counter = 0;
    newerAdds.forEach(item =>{
      let lvo = newerAdds[counter][2].toString();
      puller.getRange("F"+lastOne).setValue(newerAdds[counter][3].toString());
      puller.getRange("G"+lastOne).setValue(lvo);
      puller.getRange("E"+lastOne).setValue(newerAdds[counter][1].toString());
      counter = counter+1;
      lastOne = lastOne+1;
    });
}

    // let dumb = [];
    // toStrike.forEach(argk =>{
    //   dumb.push(argk.map(x => removeZeros(x.toString()).toString()));
      
    // })
    
    // value.forEach(bobert =>{
    //   let jojh = bobert.filter(x => removeSpaces(x.toString()));
    //   jojh.forEach(itoo =>{
    //     if(dumb.indexOf(itoo.toString))
    //   })
    // }); 

function grabvals(){
  //make it detect the name of the first sheet (so that "Summary" is replaced with whatever) and the second sheet ("Summary" but it has (number) where number should ideally be 1 because you delete the other sheet after the comparisons are obtained)
  //remember to add the edge case handlers for when one sheet is longer or shorter than the other
  let oldSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary");
  let newSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary (1)");
  let newOne = newSheet.getRange("A2:D"+newSheet.getLastRow()).getValues();
  let oldOne = oldSheet.getRange("A2:D"+oldSheet.getLastRow()).getValues();
  // let idiot = compareContrast(newOne, oldOne);
  // return idiot;
  let booobs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pull Schedule").getRange("F995").getValue();
  return typeof(booobs);
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

