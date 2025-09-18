const mysql = require('mysql2');


const pool = mysql.createPool({
  host: process.env.DB_HOST || 'localhost',
  port: process.env.DB_PORT || 3306,
  user: process.env.DB_USER || 'root',
  password: process.env.DB_PASSWORD || 'Admin@123',
  database: process.env.DB_NAME || 'amc',
  waitForConnections: true,
  connectionLimit: 10,
  queueLimit: 0,

});

// Test connection
pool.getConnection((err, connection) => {
  if (err) {
    console.error('❌ Database connection failed:', err.message);
    process.exit(1);
  } else {
    console.log('✅ Database connected successfully');
    connection.release();
  }
});

module.exports = pool.promise();