// ## Note: ---------------------------------------------------------------------------------
// Suggesting Use of `Excel Viewer` Vs Code extension
// Excel Viewer by GrapeCity is a popular choice that provides basic viewing and navigation 
// for Excel files (.xlsx and .xlsm) within VS Code. It allows you to see the contents of 
// the file in a FlexSheet control, navigate between multiple sheets if your spreadsheet has 
// them, and do some basic data selection and previewing
// ## #######################################################################################
 

const PORT = 2610;
const express = require('express');
const bodyParser = require('body-parser');
const app = express();
const ExcelJS = require('exceljs');
const PATH_TO_OUT = 'out/excel/'
const fs = require('fs')

init()

// ## Middleware Section ------------------------------
app.use(bodyParser.json());
// ## Middleware Section ==============================


// ## Endpoints  --------------------------------------
app.all('/convert-to-excel', (req, res) => {
    // console.log('Received data:', req.body);
    convertToExcel()
    res.send('Converting to Excel');
});
// ## Endpoints  ======================================


app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`)
});


function convertToExcel() {

    // Create a new workbook
    const workbook = new ExcelJS.Workbook();

    // Add a worksheet
    const worksheet = workbook.addWorksheet('Sheet 1');

    // Add some data to the worksheet
    worksheet.columns = [
        { header: 'Name', key: 'name', width: 20 },
        { header: 'Age', key: 'age', width: 10 },
        { header: 'Country', key: 'country', width: 20 }
    ];

    // Example data
    const data = [
        { name: 'John DOE', age: 30, country: 'USA' },
        { name: 'Jane Smith', age: 25, country: 'Canada' },
        { name: 'Bob Johnson', age: 40, country: 'UK' }
    ];

    // Add the data to the worksheet
    data.forEach((row) => {
        worksheet.addRow(row);
    });

    // Save the workbook
    workbook.xlsx.writeFile(`${PATH_TO_OUT}/test.xlsx`)
        .then(() => {
            console.log('Excel file generated successfully.');
        })
        .catch((error) => {
            console.error('Error generating Excel file:', error);
        });
}

function init() {
    let pth = `${__dirname}\\${PATH_TO_OUT}`
    if (fs.existsSync(pth) === false) {
        fs.mkdirSync(pth, { recursive: true })
    }
}