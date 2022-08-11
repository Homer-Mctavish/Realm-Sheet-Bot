function showSidebar() {
  let html = HtmlService.createHtmlOutputFromFile('CALENDAR')
    .setTitle('SQL Menu');
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showSidebar(html);
}

function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .createMenu('Realm Custom Scripts')
    .addItem('Show SQL Menu', 'showSidebar')
    .addToUi();
}
//preface for entire 'readFromTable' function: make the SQL database public, needs to be accessed by google itself, cannot use local connections. spooky!
/**
 *  for connection to microsoft sql server:
 *  address = 'ip address';
    user = 'database username';
    userPwd = 'database password';
    dbUrl = 'jdbc:sqlserver://' + address + ':1433;databaseName=' + 'name';
    const conn = Jdbc.getConnection(dbUrl, user, userPwd);
 */

function readFromTable() {
  try {
    let resulto = []
    //the first string is the connection string, with the ip, same port for any mySQL database, and the /being the table accessed,
    //'google' being the username and 'quote-item-price' being the password. hope nothing important's on there!
    const conn = Jdbc.getConnection('jdbc:mysql://ip:3306/sqltablename', 'usr', 'passwd');
    const start = new Date();
    const stmt = conn.createStatement();
    stmt.setMaxRows(1000);
    const results = stmt.executeQuery('SELECT 3 FROM entries');
    const numCols = results.getMetaData().getColumnCount();
    while (results.next()) {
      let rowString = '';
      for (let col = 0; col < numCols; col++) {
        rowString += results.getString(col + 1) + '\t';
        resulto.push(rowString);
      }
    }

    results.close();
    stmt.close();
    const end = new Date();
    Logger.log('Time elapsed: %sms', end - start);
    return resulto;
  } catch(err) {
    Logger.log('Failed with an error %s', err.message);
  }
}

function calendar(){
  let balendar = CalendarApp.getCalendarById("your.email@realmcontrol.com");
  let sqlQuery = readFromTable();
  balendar.createEvent(sqlQuery[1],new Date(),new Date())
}
