#!/usr/bin/env node

const Excel = require('exceljs');
const program = require('commander');
const fs = require('fs');
const moment = require('moment');

let filenameValue;

program
    .version('1.0.0')
    .description('YNAB tools to import bank statements');

program
    .command('post <filename>')
    .alias('p')
    .description('import Post statement')
    .action(filename => {
        filenameValue = filename;
    });


program.parse(process.argv);

// Remove carriage return from initial csv file
fs.readFile(filenameValue, 'latin1', function (err, data) {
    if (err) {
        return console.log(err);
    }
    let result = data.replace(/(.*)\r/gm, '$1');
    fs.writeFile(filenameValue, result, 'latin1', function (err) {
        if (err) return console.log(err);
        createResultFile();
    });
});

function createResultFile() {
    let workbook = new Excel.Workbook();
    let CSVoptions = {
        delimiter: ';'
    }

    let data = fs.readFileSync(filenameValue, {encoding: 'latin1'}).toString();
    let filenameLatin = filenameValue.replace('.csv', 'Latin1.csv');
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
                if (firstCell === 'Texte de notification') {
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


            const resultFile = 'result.csv';
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
                        let result = data.replace(/^"(.*)",/gm, '$1');
                        fs.writeFile(resultFile, result, 'utf8', function (err) {
                            if (err) return console.log(err);
                        });
                    });
                });
        });

}