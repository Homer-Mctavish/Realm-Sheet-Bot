

// function getConnection(){
// try{
//   var address = '167.206.59.124:3306';
//   var user = 'google'; 
//   var userPwd = 'quote-item-price'; 
//   var db = 'google';
//   var dbUrl = 'jdbc:mysql://' + address + '/' + db;
//   var conn = Jdbc.getConnection(dbUrl, user, userPwd);
//   return conn;
// }catch(err){
//   Logger.log('Failed with error %s', err)
// }
//   }


// function createDatabase() {
//   try {
//     jcdb.conn.stmt.execute('CREATE DATABASE ' + db);
//     jcdb.conn.close();
//   } catch (err) {
//     // TODO(developer) - Handle exception from the API
//     Logger.log('Failed with an error %s', err.message);
//   }
// }

// /**
//  * Create a new user for your database with full privileges.
//  */
// function createAdminUser() {
//   try {
//     const conn = jcdb.conn;
//     const stmt = conn.prepareStatement('CREATE USER ? IDENTIFIED BY ?');
//     stmt.setString(1, user);
//     stmt.setString(2, userPwd);
//     stmt.execute();
//     conn.createStatement().execute('GRANT ALL ON `%`.* TO ' + user);
//     conn.close();
//   } catch (err) {
//     // TODO(developer) - Handle exception from the API
//     Logger.log('Failed with an error %s', err.message);
//   }
// }

// /**
//  * Create a new table in the database.
//  */
// function createTable() {
//   try {
//     let stmt = jcdb.conn.stmt.execute('CREATE TABLE entries ' +
//       '(guestName VARCHAR(255), content VARCHAR(255), ' +
//       'entryID INT NOT NULL AUTO_INCREMENT, PRIMARY KEY(entryID));');
//     stmt.close();
//   } catch (err) {
//     // TODO(developer) - Handle exception from the API
//     Logger.log('Failed with an error %s', err.message);
//   }
// }

// function writeOneRecord() {
//   try {
//     const stmt = jcdb.conn.prepareStatement('INSERT INTO entries ' +
//       '(guestName, content) values (?, ?)');
//     stmt.setString(1, 'First Guest');
//     stmt.setString(2, 'Hello, world');
//     stmt.execute();
//     stmt.close();
//   } catch (err) {
//     // TODO(developer) - Handle exception from the API
//     Logger.log('Failed with an error %s', err.message);
//   }
// }

// /**
//  * Write recordNum rows of data to a table in a single batch.
//  */
// function writeManyRecords(recordNum) {
//   try {
//     jcdb.conn.setAutoCommit(false);
//     const start = new Date();
//     const stmt = jcdb.conn.prepareStatement('INSERT INTO entries ' +
//       '(guestName, content) values (?, ?)');
//     for (let i = 0; i < recordNum; i++) {
//       stmt.setString(1, 'Name ' + i);
//       stmt.setString(2, 'Hello, world ' + i);
//       stmt.addBatch();
//     }
//     stmt.close();
//     const batch = stmt.executeBatch();
//     jcdb.conn.commit();
//     jcdb.conn.close();

//     const end = new Date();
//     Logger.log('Time elapsed: %sms for %s rows.', end - start, batch.length);
//   } catch {
//     // TODO(developer) - Handle exception from the API
//     Logger.log('Failed with an error %s', err.message);
//   }
// }

function readFromTable() {
  try {
    resulto = []
    const conn = Jdbc.getConnection('jdbc:mysql://167.206.59.124:3306/google', 'google', 'quote-item-price');
    const start = new Date();
    const stmt = conn.createStatement();
    stmt.setMaxRows(1000);
    const results = stmt.executeQuery('SELECT * FROM entries');
    const numCols = results.getMetaData().getColumnCount();

    while (results.next()) {
      let rowString = '';
      for (let col = 0; col < numCols; col++) {
        rowString += results.getString(col + 1) + '\t';
        resulto.push(results);
      }
      // Logger.log(rowString);
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



// function sendRes(){
//   var t = readFromTable();
//   t.forEach(val=>{
//     Logger.log(val);
//   })
// }

// function importtosheets(name){
//   var imptar = Spreadsheet.getSheetByName(name);
//   var values = readFromTable()
//   var rowstocopy = getLastDataRow(s);
//   var colstocopy = getLastDataCol(s);
//   imptar.getRange(1, 1, rowstocopy, colstocopy).setValues(values);
// }
