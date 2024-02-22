// file reading
//await workbook.xlsx.readFile("C:\Users\Lenovo\Documents\GIGGAT GARDENS\Water_Bills_2023_GIGGAT.xlsx");

const excel = require('exceljs');
const path = require('path');

const wb = new excel.Workbook();
const fileName = "C:/Users/Lenovo/Documents/GIGGAT_GARDENS/Water_Bills_2023_GIGGAT.xlsx";

wb.xlsx.readFile(fileName)
    .then(() => {
        const ws = wb.getWorksheet("December");

         for (let rowIndex = 3; rowIndex <= 19; rowIndex++) { // Column index starts from 1, H corresponds to 8th column
            const cell = ws.getCell(`B${rowIndex}`);//column B corresponding to house number
            const cell2 = ws.getCell(`H${rowIndex}`);//column H corresponding to water bill value
            // Iterate over each cell in the column
                console.log(cell.value);
                console.log(cell2.value.result);//value.result gives the exact value rather than value which contains alot of necessary data about the cell including its formula
        
        
            }
    })
    .catch(err => {
        console.error("Error:", err.message);
        //console.error("Full path:", path.resolve(fileName)); // Print the resolved path for debugging
    });

