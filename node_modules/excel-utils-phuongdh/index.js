const ExcelJS = require('exceljs');
const XLSX = require('xlsx');

// Function to read Excel file
async function readExcel(filePath) {
    try {
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0]; // Assuming only one sheet
        const worksheet = workbook.Sheets[sheetName];
        return XLSX.utils.sheet_to_json(worksheet);
    } catch (error) {
        throw error;
    }
}

// Function to write data to Excel file
async function writeExcel(filePath, data) {
    try {
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Sheet 1');
        sheet.addRows(data);
        await workbook.xlsx.writeFile(filePath);
        console.log('Excel file created successfully.');
    } catch (error) {
        throw error;
    }
}

module.exports = { readExcel, writeExcel };
