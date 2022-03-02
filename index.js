#!/usr/bin/env node

const Excel = require('exceljs');
const program = require('commander');
const fs = require('fs');
const moment = require('moment');
const axios = require('axios');
const crypto = require('crypto');

program
    .version('1.0.0')
    .description('YNAB tools to import bank statements');

program
    .command('post <filename>')
    .alias('p')
    .description('import Post statements')
    .action(filename => {
        readPostCsvFile(filename)
    });

program
    .command('wise <startDate> <endDate>')
    .alias('t')
    .description('import Wise statements (date format: YYYY-MM-DD)')
    .action((startDate, endDate) => {
        generateWiseCsv(startDate, endDate)
    })

program
    .command('wise-balance')
    .alias('b')
    .description('Output wise account balance')
    .action(() => {
        wiseBalance()
    })


program.parse(process.argv);


function generateWiseCsv(startDate, endDate) {
    const wiseConfigs = JSON.parse(readWiseConfig());

    for (const config of wiseConfigs) {
        getTransactions(config, startDate, endDate).then((response) => {
            if (response) {
                let {
                    transactions,
                    accountHolder: {
                        firstName,
                        lastName
                    }
                } = response.data;

                const filename = ('wise-' + firstName + '-' + lastName + '.csv').toLowerCase();
                createWiseCsv(transactions, filename)
            } else {
                console.error("No data available for profile [" + config.profileId + "] for date range " + startDate + " to " + endDate);
            }
        });
    }

}

function createWiseCsv(transactions, outputFilename) {
    let fileContent = 'Memo,Inflow,Outflow,Date\n';
    for (const transaction of transactions) {
        let {
            amount: {
                value: amount
            },
            date,
            details: {
                description
            },
            totalFees: {value: fees},
            type
        } = transaction;

        description = removeComma(description) + ' (fees: ' + fees + ')';

        date = moment(date).format('YYYY-MM-DD')
        let inflow = '', outflow = '';
        if (type === 'DEBIT') {
            outflow = Math.abs(amount).toString();
        } else {
            inflow = amount.toString();
        }
        fileContent += description + ',' + inflow + ',' + outflow + ',' + date + '\n';
    }
    fs.writeFile(outputFilename, fileContent, function (err) {
        if (err) {
            throw err
        } else {
            console.log('New file created : ' + outputFilename);
        }
    });
}

function getTransactions(config, startDate, endDate, x2faApproval = null, signedX2fa = null) {
    let url = 'https://api.transferwise.com/v3/profiles/' + config.profileId + '/borderless-accounts/' + config.accountId + '/statement.json?currency=' + config.currency + '&intervalStart=' + startDate + 'T00:00:00.000Z&intervalEnd=' + endDate + 'T23:59:59.999Z';

    const requestHeaders = {}
    requestHeaders.headers = {
        Authorization: 'Bearer ' + config.token
    }

    if (signedX2fa) {
        requestHeaders.headers['X-signature'] = signedX2fa
    }

    if (x2faApproval) {
        requestHeaders.headers['x-2fa-approval'] = x2faApproval
    }

    return axios.get(url, requestHeaders
    ).catch(error => {
        let {
            "x-2fa-approval-result": x2faApprovalResult,
            "x-2fa-approval": x2faApproval,
        } = error.response.headers;
        if (x2faApprovalResult === "REJECTED") {
            const privateKey = fs.readFileSync(config.privateKeyPath, "utf8")
            const sign = crypto.createSign('SHA256');
            sign.write(x2faApproval);
            sign.end();
            const signedX2fa = sign.sign(privateKey, 'base64');

            return getTransactions(config, startDate, endDate, x2faApproval, signedX2fa);
        }
    })
}

function getAccountBalance(config) {
    let url = 'https://api.transferwise.com/v3/profiles/' + config.profileId + '/balances?types=STANDARD';

    const requestHeaders = {}
    requestHeaders.headers = {
        Authorization: 'Bearer ' + config.token
    }

    return axios.get(url, requestHeaders)
}

function wiseBalance() {
    const wiseConfigs = JSON.parse(readWiseConfig());

    for (const [i, config] of wiseConfigs.entries()) {
        getAccountBalance(config).then((response) => {
            if (response) {
                console.log(`Balance for ${config.name}:`)
                console.log("------------------")
                for (const [i, balance] of response.data.entries()) {
                    console.log(` ${balance.currency} ${balance.amount.value} `)
                    if (i === response.data.length - 1) {
                        console.log();
                    }
                }

            } else {
                console.error(`No balance available for profile [${config.profileId}]`);
            }
        });
    }
}

function readWiseConfig() {
    return fs.readFileSync('wise.config.json', 'utf8');
}

function readPostCsvFile(filename) {
// Remove carriage return from initial csv file
    fs.readFile(filename, 'utf8', function (err, data) {
        if (err) {
            return console.log(err);
        }
        //Remove carriage returns
        let result = data.replace(/(.*)\r/gm, '$1');
        //Remove comma
        result = removeComma(result);
        fs.writeFile(filename, result, 'utf8', function (err) {
            if (err) return console.log(err);
            createResultFile(filename);
        });
    });
}

function createResultFile(filename) {
    let workbook = new Excel.Workbook();

    let CSVoptions = {
        // Used to return all value as string
        // without this option Date might not be correctly parsed depending on the System locale
        map(value, index) {
            switch (index) {
                case 2 : // Credit column
                    if (value) {
                        return parseFloat(value)
                    }
                case 3 : // Debit column
                    if (value) {
                        return parseFloat(value)
                    }
                default: // Everything else is parsed as String (date column included)
                    return value;
            }
        },
        delimiter: ';'
    }

    let data = fs.readFileSync(filename, {encoding: 'utf-8'}).toString();
    let filenameLatin = filename.replace('.csv', 'Latin1.csv');
    fs.writeFileSync(filenameLatin, data);

    workbook.csv.readFile(filenameLatin, CSVoptions)
        .then(function (worksheet) {
            // Make outflow value positive (YNAB doesn't accept negative value)
            worksheet.getColumn(4).eachCell(function (cell) {
                if (Number.isFinite(cell.value)) {
                    cell.value = Math.abs(cell.value);
                }
            });


            // Empty last "Solde" column (not used)
            worksheet.getColumn(6).eachCell(function (cell) {
                cell.value = null;
            });

            // Remove first date column (not used)
            worksheet.spliceColumns(1, 1);

            const nbRows = worksheet.lastRow.number;

            // Retrieve header row position
            let headerRowPosition = 0;
            for (var i = 0; i < nbRows; i++) {
                var firstCell = worksheet.getRow(i).getCell('A').value;
                if (firstCell === 'Texte de notification' || firstCell === 'Notification text') {
                    headerRowPosition = i;
                }
            }

            // Remove headers rows, just keep the last header row
            worksheet.spliceRows(0, headerRowPosition - 1);

            // Replace first header row with the following YNAB compatible header name
            const firstRow = worksheet.getRow(1);
            firstRow.getCell('A').value = 'Memo';
            firstRow.getCell('B').value = 'Inflow';
            firstRow.getCell('C').value = 'Outflow';
            firstRow.getCell('D').value = 'Date';

            worksheet.eachRow(function (row, rowNumber) {
                if (rowNumber > 1) {
                    let dateCell = row.getCell('D').value;
                    row.getCell('D').value = moment(dateCell, 'DD.MM.YYYY').format('YYYY-MM-DD');
                }
            });


            const resultFile = 'post-ynab-importable.csv';
            workbook.csv.writeFile(resultFile)
                .then(function () {
                    console.log('Converted file : ' + resultFile);
                    //Delete temporary file
                    fs.unlinkSync(filenameLatin);

                    //Remove superfluous double quote
                    fs.readFile(resultFile, 'utf8', function (err, data) {
                        if (err) {
                            return console.log(err);
                        }
                        let result = data.replace(/^"(.*)"[\s]*,/gm, '$1,');
                        fs.writeFile(resultFile, result, 'utf8', function (err) {
                            if (err) return console.log(err);
                        });
                    });
                });
        });

}

function removeComma(str) {
    return str.replace(/,\s|\s,|,/gm, ' ');
}
