const { readExcel, writeExcel } = require('./excel-utils-phuongdh');

// Reading Excel file
readExcel('input.xlsx')
    .then(data => {
        console.log('Data read from Excel:', data);
    })
    .catch(error => {
        console.error('Error reading Excel file:', error);
    });
