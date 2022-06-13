function getConnection(){
try{
  var address = 'db4free.net:3306';
  var user = 'ferrisbu'; 
  var userPwd = 'holycrumpet'; 
  var db = 'testthisstupid';
  var dbUrl = 'jdbc:mysql://' + address + '/' + db;
  var conn = Jdbc.getConnection(dbUrl, user, userPwd);
  return conn;
}catch(err){
  Logger.log('Failed with error %s', err)
}
}

function createDatabase() {
  try {
    const conn = Jdbc.getConnection('jdbc:mysql://yoursqlserver.example.com:3306/database_name',{user: 'username', password: 'password'});
    conn.createStatement().execute('CREATE DATABASE ' + db);
    conn.close();
  } catch (err) {
    // TODO(developer) - Handle exception from the API
    Logger.log('Failed with an error %s', err.message);
  }
}

/**
 * Create a new user for your database with full privileges.
 */
function createAdminUser() {
  try {
    const conn = Jdbc.getConnection('jdbc:mysql://yoursqlserver.example.com:3306/database_name',{user: 'username', password: 'password'});
    const stmt = conn.prepareStatement('CREATE USER ? IDENTIFIED BY ?');
    stmt.setString(1, user);
    stmt.setString(2, userPwd);
    stmt.execute();
    conn.createStatement().execute('GRANT ALL ON `%`.* TO ' + user);
    conn.close();
  } catch (err) {
    // TODO(developer) - Handle exception from the API
    Logger.log('Failed with an error %s', err.message);
  }
}

/**
 * Create a new table in the database.
 */
function createTable() {
  try {
    const conn = Jdbc.getConnection('jdbc:mysql://yoursqlserver.example.com:3306/database_name',{user: 'username', password: 'password'});
    let stmt = conn.createStatement().execute('CREATE TABLE entries ' +
      '(guestName VARCHAR(255), content VARCHAR(255), ' +
      'entryID INT NOT NULL AUTO_INCREMENT, PRIMARY KEY(entryID));');
    stmt.close();
  } catch (err) {
    // TODO(developer) - Handle exception from the API
    Logger.log('Failed with an error %s', err.message);
  }
}

function writeOneRecord() {
  try {
    const conn = Jdbc.getConnection('jdbc:mysql://yoursqlserver.example.com:3306/database_name',{user: 'username', password: 'password'});
    const stmt = conn.prepareStatement('INSERT INTO entries ' +
      '(guestName, content) values (?, ?)');
    stmt.setString(1, 'First Guest');
    stmt.setString(2, 'Hello, world');
    stmt.execute();
    stmt.close();
  } catch (err) {
    // TODO(developer) - Handle exception from the API
    Logger.log('Failed with an error %s', err.message);
  }
}

/**
 * Write 500 rows of data to a table in a single batch.
 */
function writeManyRecords(recordNum) {
  try {
    const conn = Jdbc.getConnection('jdbc:mysql://yoursqlserver.example.com:3306/database_name',{user: 'username', password: 'password'});
    conn.setAutoCommit(false);
    const start = new Date();
    const stmt = conn.prepareStatement('INSERT INTO entries ' +
      '(guestName, content) values (?, ?)');
    for (let i = 0; i < recordNum; i++) {
      stmt.setString(1, 'Name ' + i);
      stmt.setString(2, 'Hello, world ' + i);
      stmt.addBatch();
    }
    stmt.close();
    const batch = stmt.executeBatch();
    conn.commit();
    conn.close();

    const end = new Date();
    Logger.log('Time elapsed: %sms for %s rows.', end - start, batch.length);
  } catch {
    // TODO(developer) - Handle exception from the API
    Logger.log('Failed with an error %s', err.message);
  }
}

function readFromTable() {
  try {
    const conn = Jdbc.getConnection('jdbc:mysql://yoursqlserver.example.com:3306/database_name',{user: 'username', password: 'password'});
    const start = new Date();
    const stmt = conn.createStatement();
    stmt.setMaxRows(1000);
    const results = stmt.executeQuery('SELECT * FROM entries');
    const numCols = results.getMetaData().getColumnCount();

    while (results.next()) {
      let rowString = '';
      for (let col = 0; col < numCols; col++) {
        rowString += results.getString(col + 1) + '\t';
      }
      Logger.log(rowString);
    }

    results.close();
    stmt.close();

    const end = new Date();
    Logger.log('Time elapsed: %sms', end - start);
  } catch {
    // TODO(developer) - Handle exception from the API
    Logger.log('Failed with an error %s', err.message);
  }
}