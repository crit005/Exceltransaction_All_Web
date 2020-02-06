const mysql = require('mysql');
const BaseDatabase = require("mysql-async-wrapper").default;

const pool = mysql.createPool({
    connectionLimit: 100, //important
    host: "server",
    user: "konthea",
    password: "konthea123",
    port: 3305,
    database: "mainkartuqq"
    // debug: false
});
const maxRetryCount = 3; // Number of Times To Retry
const retryErrorCodes = ["ER_LOCK_DEADLOCK", "ERR_LOCK_WAIT_TIMEOUT"] // Retry On which Error Codes 

const db = new BaseDatabase(pool, {
    maxRetryCount,
    retryErrorCodes
});
module.exports = db;