const yargs = require('yargs');
const xlsx = require('xlsx')
const mysql = require('mysql');
const colors = require('colors');
const moment = require('moment');
const cTable = require('console.table');
const Table = require('cli-table');
const _progress = require('cli-progress');
const io = require('console-read-write');
const db = require("./database.js");
var cowsay = require("cowsay");
const fs = require('fs');
const mysqldump = require('mysqldump');
const serverAdr = {
    host: "server",
    user: "konthea",
    password: "konthea123",
    port: 3305,
    database: "bookiepalacedb"
}

var con = mysql.createConnection(serverAdr);


const fromDate = '2009-01-01';
let comBank = [];

let fileName;
let sheetName;
let transactionDate;
let excelComBank = [];
let excelDatas = new Object;
let excelUsers = [];
let excelCss = [];
let excelAccs = [];
let serverUserDatas = [];
let serverStaffs = [];
var arrTransaction = [];
let connectToDB;
let importBy;
let transactionCode;

async function main() {
    if (yargs.argv._[0] && yargs.argv._[0] == 'rollback') {
        //rollback --trcod 234234234-324
        initRollback();
    } else {
        console.log(cowsay.say({
            text: `Welcom To BOOKIEPALACE Excel Transaction`,
            e: "oO",
            T: "U "
        }).brightCyan);
        try {
            connectToDB = await ConnectToDB();
        } catch (e) {
            console.log("========Error=======".red);
            console.log(e.message.red);
            return;
        }

        await catchParamInut();
        Start();
    }

}

async function catchParamInut() {

    fileName = yargs.argv._[0] ? yargs.argv._[0] : null;
    CatchFileName: while (true) {
        if (!fileName) {
            fileName = await InputPrompt({ label: "Enter Your File: ".brightCyan, required: true });
        }
        if (!isFileExis(fileName)) {
            console.log("File Not Found!".red);
            fileName = null;
            continue CatchFileName;
        }
        break;
    }
    sheetName = yargs.argv.sheet ? yargs.argv.sheet : null;
    CatchSheetName: while (true) {
        if (!sheetName) {
            sheetName = await InputPrompt({ label: "Enter Sheet Name: ".brightCyan, required: true });
        }
        if (sheetName == "Banks") {
            console.log("Sheet Banks It Contains Of Company Bank Please Try Another Sheet!".red);
            sheetName = null;
            continue CatchSheetName;
        }
        break;
    }

    transactionDate = yargs.argv.date ? yargs.argv.date : null;
    CatchDateTransaction: while (true) {
        if (!transactionDate) {
            transactionDate = await InputPrompt({ label: "Enter Date Transaction: \nYYYY-MM-DD".brightCyan, required: true });
        }
        if (!CheckDate(transactionDate)) {
            console.log("Invalid Date".red);
            transactionDate = null;
            continue CatchDateTransaction;
        }
        break;
    }

    importBy = yargs.argv.importby ? yargs.argv.importby : null;
    CatchImportBy: while (true) {
        if (!importBy) {
            importBy = await InputPrompt({ label: "Your Name: ".brightCyan, required: true });
        }
        if (importBy == "") {
            console.log("Your Name Cannot Empty".red);
            importBy = null;
            continue CatchImportBy;
        }
        break;
    }

}

main();
async function Start() {
    consoleTitle("Check Status");
    try {
        arrTransaction = [];
        if (connectToDB.status) {
            console.log(connectToDB.message.green);
            //** Check youser input
            if (!await CheckUserInput()) {
                throw { message: "Error Data Input!" };
            };
            if (excelComBank.status) {
                await InitCombank(excelComBank.banks);
            } else throw excelComBank;
            serverUserDatas = await getCustomerDAta();
            CheckSelectUser(serverUserDatas, excelUsers);
            await InitStaff(excelAccs, excelCss);
            console.log("Staff data checked".green);
            ValidateExcelData(excelDatas);
            console.log("Data Validation Checked".green);
            await GenerateTransactionData(excelDatas);
            //**=========En check user input==============


            //* Confirm for commit data
            //* ============================
            console.log(`\n\nYour transaction code is `.brightGreen + transactionCode.brightCyan + `, please note this code befor Submit to server`.brightGreen);
            console.log(`\nIn case you got Unhandle error, send this transaction code to IT to restor your data.`.brightCyan);

            await CommitToserver(arrTransaction);

            consoleTitle("::All The Progress Are Done::");
            var restart = await ConfirmPrompt("\nDo You Want To Start With Other File?".brightCyan);
            if (restart) {
                process.stdout.write('\033c') // Good then console.clear()
                await initInput();
                return;
            } else {
                console.log(cowsay.say({
                    text: `Good By`,
                    e: "oO",
                    T: "U "
                }).brightRed);
            }
            con.end();
            process.exit(1);
        } else { throw connectToDB; }
    } catch (e) {
        console.log("========Error=======".red);
        console.log(e.message.red);
        console.log(e);
        var reTry = await InputPrompt({ label: "Type ".brightCyan + "(R)".brightGreen + " For Retry Or Type ".brightCyan + "(RE)".brightGreen + " For ReEnter ".brightCyan + "Or Other for Stop!".brightRed });
        reTry = reTry.toString().toUpperCase();
        if (reTry == "R") {
            process.stdout.write('\033c') // Good then console.clear()
            Start();
            return;
        } else if (reTry == "RE") {
            process.stdout.write('\033c') // Good then console.clear()
            await initInput();
            return;
        } else {
            console.log(cowsay.say({
                text: `Good By`,
                e: "oO",
                T: "U "
            }).brightRed);
        }
        con.end();
        process.exit(1);
    }
}


function ConnectToDB() {
    return new Promise((resolve, reject) => {
        con.connect(function (err) {
            //if (err) throw err;
            if (err) reject({ "status": false, "message": err.message });
            resolve({ "status": true, "message": "Connect to Database Successfully" });
        });
    });
}

function CheckUserInput() {
    return new Promise((resolve, reject) => {
        let err = false;
        try {
            var wb = xlsx.readFile(fileName, { cellDates: true });
            var wsb = wb.Sheets["Banks"];
            var ws = wb.Sheets[sheetName];
            console.log(("File name " + fileName + " checked").green);

            //* Vlidate Date input
            if (CheckDate(transactionDate)) {
                console.log(("Date Transaction " + transactionDate + " checked").green);
            }
            else {
                err = true;
                console.log("Invalid Input date ".red + transactionDate.red);
            }
            //* Check Sheet Bank
            excelComBank = ExcelComBank(wsb);
            if (excelComBank.status) {
                console.log(excelComBank.message.green);
            } else {
                err = true;
                console.log(excelComBank.message.red);
            }

            //* ValidateDate Sheet input
            if (CheckSheet(ws)) {
                console.log(("Sheet name " + sheetName + " checked").green)
            }
            else {
                err = true;
                console.log(("Sheet name " + sheetName + " not fount").red);
            }
            resolve(!err);

        } catch (e) {
            reject(e);
        }
    });
}

function CheckDate(trDate) {
    let arrDate = trDate.split("-");
    if (arrDate.length == 3 && arrDate[1].length == 2 && arrDate[2].length == 2) {
        if ((!isNaN(arrDate[0]) && arrDate[0] > 2000) && (!isNaN(arrDate[1]) && arrDate[1] <= 12) && (!isNaN(arrDate[2]) && arrDate[2] <= 31)) {
            let ymd = new Date(trDate);
            if ("Invalid Date" != ymd) return true;
            else return false
        }
    }
    return false;
}

function CheckSheet(cws) {
    if (!cws) {
        return false;
    } else {
        excelDatas = ExcelData(cws);
        return true;
    }
}

function ExcelComBank(cwsb) {
    if (!cwsb) {
        return { message: "Cannot fined sheet name (Bank) in your file!", status: false };
    } else {
        var bankDatas = xlsx.utils.sheet_to_json(cwsb);
        const count = Object.keys(bankDatas).length;
        if (count != 1) {
            return { message: "Invalid Format of Sheet Bank! Check and Follow Format File in Folder SampleData", status: false }
        }
        return { banks: bankDatas[0], status: true, message: "Bank Sheet checked" };
    }
}

/** Fielter data from excelt add to varable excelDatas */
function ExcelData(ws) {
    var datas = xlsx.utils.sheet_to_json(ws);
    try {
        datas = datas.filter((item) => { return (item["Remark"].toString().toUpperCase().trim() == "DP" || item["Remark"].toString().toUpperCase().trim() == "WD"); });
    } catch (e) { 
        throw { message: "Invalid Value in Remark" };
    }

    var dataTitle = Object.keys(datas[0]);
    if (!dataTitle.includes("No") || !dataTitle.includes("CS") || !dataTitle.includes("Acc") || !dataTitle.includes("Nama") || !dataTitle.includes("Username") || !dataTitle.includes("Time") || !dataTitle.includes("Remark") || dataTitle.includes("__EMPTY_1") || !matchBankTitle(dataTitle)) {
        throw { message: "Invalid Data Format in Excell File \nCheck the follow file in folder sample" };
    }
    datas.map((data) => {
        if (data.CS) excelCss.push(data["CS"].trim());
        if (data.Acc) excelAccs.push(data["Acc"].trim());
        if (data.Username) excelUsers.push(data["Username"].trim());
    });

    excelCss = UniqueArray(excelCss);
    excelAccs = UniqueArray(excelAccs);
    excelUsers = UniqueArray(excelUsers);

    return datas;
}

function matchBankTitle(dataTitle) {
    var bankNames = Object.keys(excelComBank.banks);
    var match = false;
    bankNames.some((bankName) => {
        if (dataTitle.includes(bankName)) {
            match = true;
            return true;// brack loop
        }
    });
    if (!match) throw { message: "Bank In Sheet Containt Transactions are not match with Bank in Sheet Banks" }
    return match;
}

async function InitCombank(bankDatas) {
    await Promise.all(
        Object.entries(bankDatas).map((bank) => {
            var sql = "SELECT Id, accountName, accountNumber, bankName, '" + bank[0] + "' as exBank FROM bank where id = " + bank[1];
            return getServerBank(sql, bank[0], bank[1]);
        })
    );
}

function getServerBank(sql, bankName, id) {
    return new Promise((resolve, reject) => {
        try {
            con.query(sql, function (err, result, fields) {
                if (err) reject(err);
                if (result && result.length > 0) {
                    comBank[result[0].exBank] = { Id: result[0].Id, accountName: result[0].accountName, accountNumber: result[0].accountNumber, bankName: result[0].bankName };
                    resolve();
                } else {
                    reject({ message: "Bank " + bankName + " with ID " + id + " not match in SYSTEM" });
                }
            });
        } catch (e) {
            reject(e.message);
        }

    });
}

function getCustomerDAta() {
    return new Promise((resolve, reject) => {
        try {
            var strCondiction = ArrayToSqlString(excelUsers);
            var sql = "SELECT webaccount.customerId AS customerid, customer.mobile AS customerMobile, webaccount.Id AS webAccountId, webaccount.loginid AS webAccountLoginId, webaccount.clubId AS clubId, bank.bankName, bank.accountName, bank.accountNumber, 0 as oldBalance FROM customer INNER JOIN webaccount ON customer.Id = webaccount.customerId INNER JOIN bank ON webaccount.customerId = bank.customerId WHERE webaccount.loginId in (" + strCondiction + ")";
            con.query(sql, function (err, result) {
                if (err) reject(err);
                if (result && result.length > 0) {
                    const customers = [];
                    result.map((customer) => {
                        customers[customer.webAccountLoginId] = customer;
                    }
                    );
                    resolve(customers);
                } else {
                    reject({ message: "There are not User found" });
                }
            });
        } catch (e) {
            reject(e);
        }
    });
}



function UniqueArray(arr) {
    // var arr = ["Mike","Matt","Nancy","Adam","Jenny","Nancy","Carl"];
    uniqueArray = arr.filter(function (item, pos) {
        return arr.indexOf(item) == pos;
    })
    return uniqueArray;
}

function ArrayToSqlString(arr) {
    var str = "'" + arr.toString().replace(/,/gi, "','") + "'";
    return str;
}

function CheckSelectUser(serverCustomers, excelUsers) {
    msg = "";
    err = false;
    if (Object.keys(serverCustomers).length == excelUsers.length) {
        excelUsers.map((user) => {
            if (!serverCustomers[user]) {
                err = true;
                msg += (user + " Problem with lowercase or uppercase").red + "\n";
            }

        })
    } else {
        excelUsers.map((user) => {
            if (!serverCustomers[user]) {
                err = true;
                msg += (user + " is not match with system or problem with Charater").red + "\n";
            }
        })
    }
    if (err) {
        throw { message: msg };
    }
}
async function InitStaff(arrAcc, arrCs) {
    var arrStaffs = arrAcc.concat(arrCs);
    await Promise.all(

        arrStaffs.map((staff) => {
            var sql = "SELECT loginid, '" + staff + "' as exStaff FROM staff WHERE loginid = '" + staff + "'";
            return getServerStaff(sql, staff);
        })
    );
}

function getServerStaff(sql, staff) {
    return new Promise((resolve, reject) => {
        //console.log(sql);
        try {
            con.query(sql, function (err, result) {
                if (err) reject(err);
                if (result && result.length > 0) {
                    serverStaffs[result[0].exStaff] = result[0].loginid;
                    resolve();
                } else {
                    reject({ message: staff + " Not Match In Server" });
                }
            });
        } catch (e) {
            reject(e);
        }

    });
}

function getCustomerBalance(webid, fromDate, toDate) {
    return new Promise((resolve, reject) => {
        try {
            var arrCondition = [webid, webid, webid, webid, webid, fromDate, toDate, webid, fromDate, toDate, webid, fromDate, toDate];
            var sql = "select (ifnull(totalR,0) + ifnull(totalB,0) + ifnull(totalWl,0) - ifnull(totalW,0) + ifnull(totalJ,0)) as balance from "
                + "(select sum(refill)as totalR,sum(bonus)as totalB, sum(withdrawal) as totalW, sum(wlamt)as totalWl,sum(jamount)as totalJ from "
                + "(select groupfor, rollover,totalTurnover, bonusdetail, id as transId ,webaccountid,type ,fromBankname, "
                + "if(type in (2,7),if(type=2,inputamount,if(fromBankname=?,inputamount,null)) ,null) as withdrawal, "
                + "if(type in (1,3,4,5,6,7),if(type=7,if(webaccountid=?,inputamount,null),inputamount) ,null) as refill, "
                + "if(type in (1,3,4,5,6,7),if(type=7,if(webaccountid=?,bonus,null),bonus), null) as bonus, "
                + "0 as trunover, 0 as wlamt, 0 as jamount from `transaction` "
                + "where (webaccountid=? or fromBankname=?) and status=3 and groupfor between ? and ? "
                + "union all select groupfor, null as rollover,null as totalTurnover, null as bonusdetail,100000000 as transId, accountid as webaccountid,null as type,null as fromBankname,null as withdrawal,null as refill,null bonus,trunover,wlamt,null as jamount "
                + "from winlose where accountid=?  and groupfor between ? and ? "
                + "union all select groupfor, null as rollover,null as totalTurnover, null as bonusdetail,id as transId, "
                + "if(status=2,fromId,toId) as webaccountid, null as type,null as fromBankname,0 as withdrawal,0 as refill,0 as bonus,null as trunover, null as wlamt,if(status=1,jounal.credit,jounal.debit) as jamount "
                + "from jounal where type=2 and if(status=2,fromId,toId)=?  and groupfor between ? and ? )s left join webaccount on webaccount.id=s.webaccountid) as t";
            con.query(sql, arrCondition, function (err, result) {
                if (err) reject(err);
                if (result && result.length > 0) {
                    // console.log(result[0].balance);                   
                    resolve(result[0].balance);
                    // resolve(true);
                } else {
                    reject({ message: webid + " Not found" });
                }
            });
        } catch (e) {
            reject(e);
        }
    });
}

function ValidateExcelData(excelData) {
    var err = false;
    var errData = [];
    excelData.map((data) => {
        if (!requiredField(data)) {
            err = true;
            errData.push(data);
        } else if (!validateValue(data)) {
            err = true;
            errData.push(data);
        }

    });

    if (err) throw { message: "There are somthing wrong with these records:\n" + cTable.getTable(errData) };

    function requiredField(data) {
        if (!data.CS || !data.Acc || !data.Nama || !data.Username || !data.Time || !data.Remark || !data[BankOfRecord(data)]) {
            return false;
        }
        return true;
    }

    function validateValue(dataRecord) {
        var data = dataRecord;
        var validateStringType = (typeof (data.CS) === "string" && typeof (data.Acc) === "string" && typeof (data.Username) === "string" && typeof (data.Time) === "string" && typeof (data.Remark) === "string");
        var validateBonus = false;

        if (data.Bonus) {
            if (!isNaN(data.Bonus) && typeof (data.Bonus) == "number") {
                validateBonus = true;
            } else validateBonus = false;
        } else validateBonus = true;

        var validateBankAmount = false;
        var bankAmount = data[BankOfRecord(data)];
        if (!isNaN(bankAmount) && typeof (bankAmount) == "number") {
            validateBankAmount = true;
        } else { validateBankAmount = false; }

        var validateTime = false;
        var arrTime = data.Time.split(":");
        if (arrTime.length == 2) {
            validateTime = (0 <= arrTime[0] && arrTime[0] < 24) && (0 <= arrTime[1] && arrTime[1] < 60);
        }

        return (validateStringType && validateBonus && validateBankAmount && validateTime);

    }
}

function BankOfRecord(dataRecord) {
    var bankNames = Object.keys(excelComBank.banks);
    var bankName = "";
    bankNames.some(name => {
        if (dataRecord[name] && dataRecord[name] != 0) {
            bankName = name;
            return true;
        }
    });
    if (bankName == "") { return false } else { return bankName; }
}
function consoleTitle(title) {
    if (title.length > 0) {
        strUnderline = "";
        for (var i = 0; i < title.length; i++) {
            strUnderline += "=";
        }
        console.log(("\n" + title + "\n" + strUnderline).bold.green);
    }
}

async function GenerateTransactionData(excelDatas) {
    consoleTitle("Transaction Date");
    console.log(transactionDate);
    consoleTitle("Company Bank");
    console.table(Object.values(comBank));
    consoleTitle("Generate Transaction Data for Deposit and Withdraw");
    console.log("Affected Record: " + excelDatas.length);
    console.log("Data initiating please wait...");
    transactionCode = getTransactionCode();
    const b1 = new _progress.SingleBar({
        format: 'Initiate Probress: ' + colors.brightCyan('{bar}') + '| {percentage}% || {value}/{total} || Remain Time: {eta_formatted} || Duration: {duration_formatted}',
        barCompleteChar: '\u2588',
        barIncompleteChar: '\u2591',
        hideCursor: true,
    });
    b1.start(excelDatas.length, 0);
    for (var j in excelDatas) {
        var i = parseInt(j) + 1;
        data = excelDatas[j];
        await new Promise(async (resolve, reject) => {
            let username = data.Username.trim();
            let oldBalance = 0;

            // console.log(serverUserDatas[username].oldBalance);
            if (serverUserDatas[username].oldBalance == 0) {
                oldBalance = await getCustomerBalance(serverUserDatas[username].customerid, fromDate, transactionDate);
                //console.log(oldBalance);
                serverUserDatas[username].oldBalance = oldBalance + data[BankOfRecord(data)];
            } else {
                oldBalance = serverUserDatas[username].oldBalance;
                serverUserDatas[username].oldBalance = oldBalance + data[BankOfRecord(data)];
            }

            TransactionData(data, oldBalance);
            b1.update(i);
            resolve();
        });

    }
    b1.stop();
    ShowDataBeforCommit();

}



function TransactionData(data, oldBalance) {
    const type = data["Remark"].toString().trim().toUpperCase();
    const cs = serverStaffs[data["CS"].trim()];
    const acc = serverStaffs[data["Acc"].trim()];
    const inputAmount = data[BankOfRecord(data)];
    const customer = serverUserDatas[data.Username.trim()];
    const dataBank = comBank[BankOfRecord(data)];
    var row = [];
    row["No"] = data.No;
    row["customerid"] = customer.customerid;
    row["customerMobile"] = customer.customerMobile;
    row["webAccountId"] = customer.webAccountId;
    row["webAccountLoginId"] = customer.webAccountLoginId;
    row["clubId"] = customer.clubId;
    row["oldbalance"] = oldBalance;
    row["currentBlance"] = oldBalance + inputAmount + (data.Bonus ? data.Bonus : 0);
    row["inputAmount"] = Math.abs(inputAmount);
    row["bonus"] = data.Bonus ? data.Bonus : 0;
    row["dealAmount"] = Math.abs(inputAmount);
    row["bonusDetail"] = ".|PT:0.0|MB:0.0|MD:0.0|WD:NO|RO:0.0|OT:";
    row["inputBy"] = cs;
    row["confirmBy"] = cs;
    row["dealBy"] = acc;
    row["type"] = (type == "DP") ? 1 : 2;
    row["fromBankAccountNO"] = (type == "DP") ? "" : customer.accountNumber;
    row["fromBankAccountName"] = (type == "DP") ? "" : customer.accountName;
    row["fromBankName"] = (type == "DP") ? "" : customer.bankName;
    row["toBankAccountNo"] = dataBank.accountNumber;
    row["toBankName"] = dataBank.bankName;
    row["remark"] = `Auto Input From Excel_${importBy}_` + transactionCode;
    row["status"] = "3";
    row["rollover"] = 0;
    row["totalTurnover"] = 0;
    row["groupfor"] = transactionDate.trim();
    row["createtime"] = moment().format('YYYY-MM-DD h:mm:ss');
    row["transTime"] = transactionDate + " " + data["Time"].trim() + ":00";
    // row["source"] = "";
    // row["fromSource"] = "";
    row["commission"] = 0;
    row["bankInTime"] = transactionDate + " " + data["Time"].trim() + ":00";
    row["charger"] = 0;
    arrTransaction.push(row);

}

function ShowDataBeforCommit() {
    var format = {
        'top': '═', 'top-mid': '╤', 'top-left': '╔', 'top-right': '╗'
        , 'bottom': '═', 'bottom-mid': '╧', 'bottom-left': '╚', 'bottom-right': '╝'
        , 'left': '║', 'left-mid': '╟', 'mid': '─', 'mid-mid': '┼'
        , 'right': '║', 'right-mid': '╢', 'middle': '│'
    };
    // setColorFormat(format);
    var table = new Table({
        chars: format,
        head: ["Index".brightCyan, "No".brightCyan, "Username".brightCyan, "InputBy".brightCyan, "Amount".brightCyan, "Bonus".brightCyan, "Old Balance".brightCyan, "Current Balance".brightCyan, "Deal By".brightCyan, "Date Time".brightCyan, "Type".brightCyan, "Company Bank".brightCyan, "Account No".brightCyan]
    });

    function setColorFormat(format) {
        const keys = Object.keys(format);
        keys.map(key => {
            format[key] = format[key].cyan;
        });
    }
    var i = 1;
    arrTransaction.map((row) => {
        // console.log(row.No);
        var colectRowData = [
            i.toString().gray,
            row["No"].toString().cyan,
            row["webAccountLoginId"].bold.green,
            row["inputBy"].yellow,
            row.inputAmount > 0 ? row["inputAmount"].toString().green.bold : row["inputAmount"].toString().red.bold,
            row.bonus ? row.bonus : 0,
            row.oldbalance > 0 ? row["oldbalance"].toString().green.bold : row["oldbalance"].toString().red.bold,
            row.currentBlance > 0 ? row["currentBlance"].toString().green.bold : row["currentBlance"].toString().red.bold,
            row["dealBy"].magenta,
            row.transTime,
            row.type == 1 ? "DP".green.bold : "WD".red.bold,
            row.toBankName,
            row.toBankAccountNo
        ];
        table.push(
            colectRowData
        );
        i++;
    })
    console.log(table.toString());
}

async function InputPrompt(options = { label: "Enter: ", required: "false", type: "string" }) {
    const label = options.label ? options.label : "Enter: ";
    const required = options.required ? options.required : false;
    const type = options.type ? (["number", "string", "date", "time"].includes(options.type) ? options.type : "string") : "string";
    var val = "";
    GetInput: while (true) {
        val = await io.ask(label);
        if (required && val == "") {
            console.log("Value Cannot Empty!".red.bold);
            continue GetInput;
        }
        break;
    }
    return val;
}

async function ConfirmPrompt(message) {
    var answer = false;
    AskForCommit: while (true) {
        const arrAnswers = ["y", "yes", "n", "no"];
        var input = await io.ask(message + " (" + "Yes".green + "/" + "No".red + ")");
        input = input.toString().toLowerCase();
        if (!arrAnswers.includes(input)) {
            continue AskForCommit;
        } else if (input == "yes" || input == "y") {
            answer = true;
        }
        break;
    }
    return answer;
}

async function CommitToserver(arrTransaction) {
    var commit = false;
    AskForCommit: while (true) {
        const arrAnswers = ["y", "yes", "n", "no"];
        var answer = await io.ask(("\nThere are " + arrTransaction.length + " transactions ready for submit @" + transactionDate.brightCyan + "\n\nDo you want to submit to server now? (" + "Yes".green + "/" + "No".red + ")").brightYellow);
        answer = answer.toString().toLowerCase();
        if (!arrAnswers.includes(answer)) {
            continue AskForCommit;
        } else if (answer == "yes" || answer == "y") {
            commit = true;
        }
        if (!commit) {
            if (!await ConfirmPrompt("\nDo you want to cancel this progress?".brightYellow)) {
                continue AskForCommit;
            }
        }

        break;
    }
    if (!commit) {
        // console.log("Transaction is canceled.");
        console.log(cowsay.say({
            text: `Transaction is canceled.`,
            e: "oO",
            T: "U "
        }).brightCyan);
    } else {
        var confirBackup = await ConfirmPrompt(`\nBackup Before Progress?`.brightYellow);
        if (confirBackup) {
            consoleTitle("Backup Before Progress");
            await backup();
        }
        await WriteLog(importBy, transactionCode);
        consoleTitle("Transaction Submit Progress");
        console.log("\nDon't Close This Apllication!".brightRed);
        await saveDataToserver(arrTransaction);
    }
}

async function saveDataToserver(arrTransaction) {
    try {
        const connection = await db.getConnection({ transaction: true }); // Will Begin Transaction

        let i = 0;
        const b1 = new _progress.SingleBar({
            format: 'Submit Progress:' + colors.brightYellow('{bar}') + '| {percentage}% || {value}/{total} || Remain Time: {eta_formatted} || Duration: {duration_formatted}',
            barCompleteChar: '\u2588',
            barIncompleteChar: '\u2591',
            hideCursor: true,
            // clearOnComplete: true
        });
        b1.start(arrTransaction.length, 0);
        await Promise.all(
            arrTransaction.map(async (row) => {
                const insertTransaction = getSqlInsertTransaction(row);
                const result = await connection.executeQuery(insertTransaction["columns"], insertTransaction["values"]);
                const sqlWebaccount = `UPDATE webaccount SET blance =? WHERE id = ?`;
                const updateResult = await connection.executeQuery(sqlWebaccount, [row.currentBlance, row.webAccountId]);
                if (result && updateResult) {
                    i++;
                    b1.update(i);
                    if (i >= b1.getTotal()) {
                        b1.stop();
                    }
                }
            })
        );

        await db.commit();
        console.log(cowsay.say({
            text: `All data are compleated on ${i} records\nThank Your :)`,
            e: "oO",
            T: "U "
        }).brightCyan);
        // or cowsay.think()
    } catch (err) {
        db.rollback(); // to rollback in case of errors other than query error
        b1.stop();
        throw err;
    } finally {
        db.close();
    }
}

function getSqlInsertTransaction(row) {
    var arrColumn = [];
    var arrValue = [];
    var rowAfter = [];
    var i = 0;
    for (var key in row) {
        // console.log(key);
        if (key == "No" || key == "oldbalance") continue;
        arrColumn[i] = key;
        arrValue[i] = row[key];
        i++
    }

    var strColumn = "INSERT INTO transaction (";
    for (var i = 0; i < arrColumn.length; i++) {
        strColumn += (i < arrColumn.length - 1) ? arrColumn[i] + ", " : arrColumn[i] + ") VALUES (";
    }
    for (var i = 0; i < arrColumn.length; i++) {
        strColumn += (i < arrColumn.length - 1) ? "?, " : "?)";
    }
    return { columns: strColumn, values: arrValue };
}


async function initInput() {
    console.log(cowsay.say({
        text: `Welcom To Excel Transaction`,
        e: "oO",
        T: "U "
    }).brightCyan);
    console.log("\n");

    CatchFileName: while (true) {
        fileName = await InputPrompt({ label: "Enter Your File: ".brightCyan, required: true });
        if (!isFileExis(fileName)) {
            console.log("File Not Found!".red);
            continue CatchFileName;
        }
        break;
    }

    CatchSheetName: while (true) {
        sheetName = await InputPrompt({ label: "Enter Sheet Name: ".brightCyan, required: true });
        if (sheetName == "Banks") {
            console.log("Sheet Banks It Contains Of Company Bank Please Try Another Sheet!");
            continue CatchSheetName;
        }
        break;
    }

    CatchDateTransaction: while (true) {
        transactionDate = await InputPrompt({ label: "Enter Date Transaction: \nYYYY-MM-DD".brightCyan, required: true });
        if (!CheckDate(transactionDate)) {
            console.log("Invalid Date".red);
            continue CatchDateTransaction;
        }
        break;
    }
    comBank = [];
    excelComBank = [];
    excelDatas = new Object;
    excelUsers = [];
    excelCss = [];
    excelAccs = [];
    serverUserDatas = [];
    serverStaffs = [];
    arrTransaction = [];

    Start();
}

function isFileExis(f) {
    const path = f;
    try {
        if (fs.existsSync(path)) {
            return true
        }
        return false
    } catch (err) {
        return false;
    }
}

function getTransactionCode() {
    var code = "";
    code += new Date().valueOf() + "-" + new Date().getUTCMilliseconds();
    return code;
}

function rollbackTrcode(trCode) {
    return new Promise(async (resolve, reject) => {
        try {
            var sql = `Select webAccountId, inputAmount, bonus, type From transaction Where remark like('%${trCode}')`;
            con.query(sql, async function (err, result) {
                if (err) reject(err);
                if (result && result.length > 0) {
                    console.log(`This application will delete ${result.length} records`.brightCyan);
                    var confirDelete = await ConfirmPrompt(`Do you realy want to Rollback them all form code ${trCode}?`.brightYellow);
                    if (confirDelete) {
                        await WriteLogRollback(trCode);
                        let i = 0;
                        const b1 = new _progress.SingleBar({
                            format: 'Rollback Progress:' + colors.brightYellow('{bar}') + '| {percentage}% || {value}/{total} || Remain Time: {eta_formatted} || Duration: {duration_formatted}',
                            barCompleteChar: '\u2588',
                            barIncompleteChar: '\u2591',
                            hideCursor: true,
                            // clearOnComplete: true
                        });
                        b1.start(result.length, 0);
                        await Promise.all(
                            result.map((tr) => {
                                return new Promise((resolve, reject) => {
                                    var inputAmount = (tr.type == 1) ? tr.inputAmount + tr.bonus : -(tr.inputAmount + tr.bonus);
                                    const sqlWebaccount = `UPDATE webaccount SET blance = blance-${inputAmount} WHERE id = ${tr.webAccountId}`;
                                    con.query(sqlWebaccount, function (err2, result2) {
                                        if (err2) { reject(err2) };
                                        i++;
                                        b1.update(i);
                                        if (i >= b1.getTotal()) {
                                            b1.stop();
                                        }
                                        resolve();
                                    });
                                });

                            })
                        );
                        await new Promise((resolve, reject) => {
                            const sqltr = `Delete From transaction Where remark like('%${trCode}')`;
                            con.query(sqltr, function (err3, result2) {
                                if (err3) { reject(err3) };
                                resolve();
                            });
                        });

                    } else {
                        console.log(`Rollback is canceled!`.brightCyan);
                    }
                    resolve(true);

                } else {
                    reject({ message: "No Transaction Code found" });
                }
            });
        } catch (e) {
            reject(e);
        }
    });
}

async function initRollback() {
    console.log(cowsay.say({
        text: `Application Rollback`,
        e: "RB",
        T: "U "
    }).brightCyan);
    try {
        connectToDB = await ConnectToDB();
    } catch (e) {
        console.log("========Error=======".red);
        console.log(e.message.red);
        return;
    }

    try {
        var keyword = yargs.argv._[0] ? yargs.argv._[0] : null;
        if (keyword == 'rollback') {
            var trCode = yargs.argv.trcode ? yargs.argv.trcode : null;
            if (trCode) {
                var status = await rollbackTrcode(trCode).catch((e) => { throw e; });
                if (status) {
                    console.log("Rollback Completed".green);

                } else {
                    throw { message: 'Error during Rollback!' };
                }
            } else { console.log(`Transaction Code is required (--trcode code)`.red); };
        }
        else console.log(`Invalid Keyword`.red);
        con.end();
    } catch (e) {
        console.log("\n======Error======".red);
        console.log(e.message.red);
        con.end();
        //return true;
    }
}

async function backup() {
    try {
        console.log(`\nBackup Inprogress Please Wait...`.yellow);
        await mysqldump({
            connection: serverAdr,
            dumpToFile: `./bk/${moment().format('YYYY-MM-DD_h_mm_ss')}_${transactionCode}.bk`,
        });
        console.log(`\nBackup Completed`.brightGreen);
    } catch (e) {
        throw e;
    }
}

async function WriteLog(importBy, trcode) {
    fs.appendFile('./log/log.txt', `${moment().format("YYYY-MM-DD : h:mm:ss")} Import By ${importBy} TransactionCode: ${trcode}` + "\n", function (err) {
        if (err) throw err;
        return true;
        // console.log('Saved!');
    });
}

async function WriteLogRollback(trcode) {
    fs.appendFile('./log/log.txt', `${moment().format("YYYY-MM-DD : h:mm:ss")} Transaction Rollback On TransactionCode: ${trcode}` + "\n", function (err) {
        if (err) throw err;
        return true;
        // console.log('Saved!');
    });
}