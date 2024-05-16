'use strict';

const XLSX = require('xlsx');
const extend = require('node.extend');

const excelToJson = (function() {

    let _config = {};

    const getCellRow = cell => Number(cell.replace(/[A-z]/gi, ''));
    const getCellColumn = cell => cell.replace(/[0-9]/g, '').toUpperCase();
    const getRangeBegin = cell => cell.match(/^[^:]*/)[0];
    const getRangeEnd = cell => cell.match(/[^:]*$/)[0];
    function getSheetCellValue(sheetCell) {
        if (!sheetCell) {
            return undefined;
        }
        if (sheetCell.t === 'z' && _config.sheetStubs) {
            return null;
        }
        return (sheetCell.t === 'n' || sheetCell.t === 'd') ? sheetCell.v : (sheetCell.w && sheetCell.w.trim && sheetCell.w.trim()) || sheetCell.w;
    };

    const parseSheet = (sheetData, workbook) => {
        const sheetName = (sheetData.constructor == String) ? sheetData : sheetData.name;
        const sheet = workbook.Sheets[sheetName];
        const range = sheet['!ref'];

        let maxCol = 0;
        let maxRow = 0;
        let rows = [];

        // Détecter le nombre maximal de colonnes
        for (let cell in sheet) {
            if (cell[0] === '!') continue;
            let col = cell.match(/[A-Z]+/)[0];
            let row = parseInt(cell.match(/\d+/)[0], 10);

            let colNumber = col.charCodeAt(0) - 'A'.charCodeAt(0) + 1;  // Convertir la colonne en nombre (A=1, B=2, ...)
            maxCol = Math.max(maxCol, colNumber);
            maxRow = Math.max(maxRow, row);
        }

        // Initialiser les lignes avec un nombre fixe de colonnes
        for (let i = 1; i <= maxRow; i++) {
            rows[i] = Array(maxCol).fill(null);  // Remplir chaque ligne avec 'null' pour toutes les colonnes
        }

        // Remplir les valeurs des cellules
        for (let cell in sheet) {
            if (cell[0] === '!') continue;
            let col = cell.match(/[A-Z]+/)[0];
            let row = parseInt(cell.match(/\d+/)[0], 10);
            let colIndex = col.charCodeAt(0) - 'A'.charCodeAt(0);

            rows[row][colIndex] = sheet[cell].v;
        }

        // Supprimer la première ligne vide due à l'indexation commençant à 1
        rows.shift();

        return rows;
    };


    const convertExcelToJson = function(config = {}, sourceFile) {
        const workbook = XLSX.readFile(sourceFile, {
            sheetStubs: true,
            cellDates: false
        });

        let sheetsData = {};
        workbook.SheetNames.forEach(sheetName => {
            sheetsData[sheetName] = parseSheet({ name: sheetName }, workbook);
        });

        return sheetsData;
    };


    return convertExcelToJson;
}());

module.exports = excelToJson;
