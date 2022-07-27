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


function readFromTable() {
  try {
    let resulto = []
    const conn = Jdbc.getConnection('jdbc:mysql://167.206.59.124:3306/google', 'google', 'quote-item-price');
    const start = new Date();
    const stmt = conn.createStatement();
    stmt.setMaxRows(1000);
    const results = stmt.executeQuery('SELECT 3 FROM entries');
    const numCols = results.getMetaData().getColumnCount();
    //to get whole database results.next()
    // let boli = JSON.parse(results);
    // Logger.log(boli);
    // return boli;
    // let i =0;
    while (results.next()) {
      let rowString = '';
      for (let col = 0; col < numCols; col++) {
        rowString += results.getString(col + 1) + '\t';
        resulto.push(rowString);
      }
      //Logger.log(rowString);
    }

    results.close();
    stmt.close();
    const end = new Date();
    Logger.log('Time elapsed: %sms', end - start);
    return resulto;
  } catch(err) {
    // TODO(developer) - Handle exception from the API
    Logger.log('Failed with an error %s', err.message);
  }
}

function calendar(){
  let balendar = CalendarApp.getCalendarById("samuel.spearing@realmcontrol.com");
  let sqlQuery = readFromTable();
  balendar.createEvent(sqlQuery[1],new Date(),new Date())
}

