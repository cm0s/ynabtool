#!/usr/bin/env node

var Excel = require('exceljs');

const program = require('commander');
var filenameValue;

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

var workbook = new Excel.Workbook();
var CSVoptions = {
    delimiter: ';',
    dateFormats: ['YYYY-MM-DD']
}
workbook.csv.readFile(filenameValue, CSVoptions)
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


        var nbRows = worksheet.lastRow.number;

        worksheet.spliceRows(0, 4);

        worksheet.getRow(nbRows - 4).destroy();
        worksheet.getRow(nbRows - 5).destroy();

        var firstRow = worksheet.getRow(1);
        firstRow.getCell('A').value = 'Memo';
        firstRow.getCell('B').value = 'Inflow';
        firstRow.getCell('C').value = 'Outflow';
        firstRow.getCell('D').value = 'Date';

        var options = {
            dateFormat: 'YYYY-MM-DD'
        };
        workbook.csv.writeFile('test2.csv', options)
            .then(function () {
                // done
            });
    });

