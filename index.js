#!/usr/bin/env node

const Excel = require('exceljs');
const program = require('commander');
const fs = require('fs');

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

console.log("filename:" + filenameValue);

let workbook = new Excel.Workbook();
let CSVoptions = {
    delimiter: ';',
    dateFormats: ['YYYY-MM-DD']
}

let data = fs.readFileSync(filenameValue, {encoding: 'latin1'}).toString();
let filenameLatin = filenameValue.replace('.csv', 'Latin1.csv');
fs.writeFileSync(filenameLatin, data);

workbook.csv.readFile(filenameLatin, CSVoptions)
    .then(function (worksheet) {
        worksheet.getColumn(4).eachCell(function (cell) {
            if (Number.isFinite(cell.value)) {
                cell.value = Math.abs(cell.value);
            }
        });

        worksheet.getColumn(6).eachCell(function (cell) {
            cell.value = null;
        });

        worksheet.spliceColumns(1, 1);


        let nbRows = worksheet.lastRow.number;

        worksheet.spliceRows(0, 6);

        worksheet.getRow(nbRows - 6).destroy();
        worksheet.getRow(nbRows - 7).destroy();

        let firstRow = worksheet.getRow(1);
        firstRow.getCell('A').value = 'Memo';
        firstRow.getCell('B').value = 'Inflow';
        firstRow.getCell('C').value = 'Outflow';
        firstRow.getCell('D').value = 'Date';

        let options = {
            dateFormat: 'YYYY-MM-DD'
        };
        workbook.csv.writeFile('test2.csv', options)
            .then(function () {
                //Delete temporary file
                fs.unlinkSync(filenameLatin);
            });
    });

