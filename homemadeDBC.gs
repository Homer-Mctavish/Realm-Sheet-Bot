function Query(type,db, name,tab,col,row,query){
  this.type = type;
  this.db = db;
  this.name = name;
  this.tab = tab; 
  this.col = col;
  this.row = row;
  this.query = query;
}

function createMysqlPrompt(queryTitle, queryTab,queryTab, queryColumn, queryRow, queryText){
  var queryType = 'mysql';
  var queryDb = '%YOUR DATABASE%';
  queryTitle = "Query - " + queryTitle;
  var queryObj = new Query(queryType,queryDb,queryTitle,queryTab, queryColumn,queryRow,queryText);
  PropertiesService.getScriptProperties().setProperty(queryTitle, JSON.stringify(queryObj));
  refreshPrompt();
}

function createMssqlPrompt(queryTitle, queryDb, queryTab, queryColumn, queryRow, queryText){
  var queryType = 'sqlserver';
  var QueryPrompt = ui.prompt('Name It','Title your Query' , ui.ButtonSet.OK_CANCEL);
  var queryTitle = QueryPrompt.getResponseText();
  queryTitle = "Query - " + queryTitle;
  var queryObj = new Query(queryType,queryDb,queryTitle,queryTab, queryColumn,queryRow,queryText);
  PropertiesService.getScriptProperties().setProperty(queryTitle, JSON.stringify(queryObj));
  refreshPrompt();
}

function refreshPrompt(){
  var ui = SpreadsheetApp.getUi();
  var refreshPrompt = ui.alert('Refresh Connection', 'This might overwrite current information, continue?' , ui.ButtonSet.YES_NO);
  if(refreshPrompt == ui.Button.YES) {
    ui.alert("Attempting to connect to database");
    var queries = PropertiesService.getScriptProperties().getProperties();
    for(var query in queries){
      var str = queries[query];
      if(str.indexOf("Query - ") > -1){
        var queryObj = JSON.parse(queries[query]);
        var startCell = queryObj.col + queryObj.row;
      
        readFromTable(queryObj.type,queryObj.db,queryObj.query,queryObj.tab,startCell);  
        Logger.log('Query: %s, Proof: %s', queryObj.name, queryObj.query);
      }
      Logger.log('Query: %s, Info: %s', query, queries[query]);
    }     
  }
}

function readFromTable(queryType, queryDb, query, tab, startCell) {
  // Replace the variables in this block with real values.
  var address;
  var user;
  var userPwd ;
  var dbUrl;

  switch(queryType) {
    case 'sqlserver':
      address = '%YOUR SQL HOSTNAME%';
      user = '%YOUR USE%';
      userPwd = '%YOUR PW%';
      dbUrl = 'jdbc:sqlserver://' + address + ':1433;databaseName=' + queryDb;
      break;
    case 'mysql':  
      address = '%YOUR MYSQL HOSTNAME%';
      user = '%YOUR USER';
      userPwd = '%YOUR PW%';
      dbUrl = 'jdbc:mysql://'+address + '/' + queryDb;
    break;
  }

  var conn = Jdbc.getConnection(dbUrl, user, userPwd);
  var start = new Date();
  var stmt = conn.createStatement();
  var results = stmt.executeQuery(query);
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetTab = sheet.getSheetByName(tab);
  var cell = sheetTab.getRange(startCell);
  var numCols = results.getMetaData().getColumnCount();
  var numRows = sheetTab.getLastRow();
  var headers ;
  var row =0;
  
  clearRange(tab,startCell,numRows, numCols);
  
  
  for(var i = 1; i <= numCols; i++){
    headers = results.getMetaData().getColumnName(i);
      cell.offset(row, i-1).setValue(headers);
      }
  
  while (results.next()) {
    var rowString = '';
    for (var col = 0; col < numCols; col++) {
      rowString += results.getString(col + 1) + '\t';
      cell.offset(row +1, col).setValue(results.getString(col +1 ));
    }
    row++
    Logger.log(rowString)
  }

  results.close();
  stmt.close();

  var end = new Date();
  Logger.log('Time elapsed: %sms', end - start);
}

function clearRange(tab, startCell,numRows, numCols){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tab);
  var startIndex = cellA1ToIndex(startCell);
  
  var doomedRange = sheet.getRange(startIndex.row+1, startIndex.col+1, numRows, numCols);
  
  doomedRange.clearContent();
}

/**
 * Convert a cell reference from A1Notation to 0-based indices (for arrays)
 * or 1-based indices (for Spreadsheet Service methods).
 *
 * @param {String}    cellA1   Cell reference to be converted.
 * @param {Number}    index    (optional, default 0) Indicate 0 or 1 indexing
 *
 * @return {object}            {row,col}, both 0-based array indices.
 *
 * @throws                     Error if invalid parameter
 */
function cellA1ToIndex( cellA1, index ) {
  // Ensure index is (default) 0 or 1, no other values accepted.
  index = index || 0;
  index = (index == 0) ? 0 : 1;

  // Use regex match to find column & row references.
  // Must start with letters, end with numbers.
  // This regex still allows induhviduals to provide illegal strings like "AB.#%123"
  var match = cellA1.match(/(^[A-Z]+)|([0-9]+$)/gm);

  if (match.length != 2) throw new Error( "Invalid cell reference" );

  var colA1 = match[0];
  var rowA1 = match[1];

  return { row: rowA1ToIndex( rowA1, index ),
           col: colA1ToIndex( colA1, index ) };
}

/**
 * Return a 0-based array index corresponding to a spreadsheet column
 * label, as in A1 notation.
 *
 * @param {String}    colA1    Column label to be converted.
 *
 * @return {Number}            0-based array index.
 * @param {Number}    index    (optional, default 0) Indicate 0 or 1 indexing
 *
 * @throws                     Error if invalid parameter
 */
function colA1ToIndex( colA1, index ) {
  if (typeof colA1 !== 'string' || colA1.length > 2) 
    throw new Error( "Expected column label." );

  // Ensure index is (default) 0 or 1, no other values accepted.
  index = index || 0;
  index = (index == 0) ? 0 : 1;

  var A = "A".charCodeAt(0);

  var number = colA1.charCodeAt(colA1.length-1) - A;
  if (colA1.length == 2) {
    number += 26 * (colA1.charCodeAt(0) - A + 1);
  }
  return number + index;
}


/**
 * Return a 0-based array index corresponding to a spreadsheet row
 * number, as in A1 notation. Almost pointless, really, but maintains
 * symmetry with colA1ToIndex().
 *
 * @param {Number}    rowA1    Row number to be converted.
 * @param {Number}    index    (optional, default 0) Indicate 0 or 1 indexing
 *
 * @return {Number}            0-based array index.
 */
function rowA1ToIndex( rowA1, index ) {
  // Ensure index is (default) 0 or 1, no other values accepted.
  index = index || 0;
  index = (index == 0) ? 0 : 1;

  return rowA1 - 1 + index;
}