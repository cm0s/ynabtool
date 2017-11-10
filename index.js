#!/usr/bin/env node

var XLSX = require('xlsx');

//var workbook = XLSX.readFile('test.xlsx');

process.argv.forEach(function (val, index, array) {
    console.log(index + ': ' + val);
});

const program = require('commander');

program
    .version('1.0.0')
    .description('YNAB tools to import bank statements');

program
    .command('post <filename>')
    .alias('p')
    .description('import Post statement')
    .action(filename => {
    console.log("filename:"+filename);});



program.parse(process.argv);

var workbook = XLSX.readFile(filename);