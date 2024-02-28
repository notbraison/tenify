const excel = require('exceljs');
const fs = require('fs');

// Declare tb as a global Map variable
let tb = new Map();

let fileName; // Filename of the file that will be created

async function houseAndBill(filePath, columns, rows, worksheet) {
    const wb = new excel.Workbook();

    columns = processColumn(columns);
    const { start, end } = processRows(rows);

    try {
        await wb.xlsx.readFile(filePath);
        const ws = wb.getWorksheet(worksheet);
        fileName = ws.getCell('A1').value; // Assuming the filename is in cell A1

        iterateColumns(ws, start, end, columns);
        console.log("Data extracted and stored in 'tb'"); // Indicate that data extraction is complete
    } catch (err) {
        console.error("Error:", err.message);
    }
}

function iterateColumns(ws, start, end, lettersArray) {
    for (let rowIndex = start; rowIndex <= end; rowIndex++) {
        let a, b;
        for (let letter of lettersArray) {
            const cell = ws.getCell(`${letter}${rowIndex}`);
            let cellValue;

            if (typeof cell.value === 'string') {
                cellValue = cell.value;
            } else if (typeof cell.value === 'object' && 'result' in cell.value) {
                cellValue = cell.value.result;
            } else {
                cellValue = cell.value;
            }

            if (letter === 'B') {
                cellValue = String(cellValue);
                a = 'House : ' + cellValue;
            }
            if (letter === 'H') {
                b = cellValue;
            }

            tb.set(a, b);
        }
    }
}

function processRows(input) {
    let start, end;

    if (typeof input === 'string') {
        const rangeMatch = input.match(/^(\d+)-(\d+)$/);
        if (rangeMatch) {
            start = parseInt(rangeMatch[1]);
            end = parseInt(rangeMatch[2]);
        }
    }

    return { start, end };
}

function processColumn(input) {
    if (!Array.isArray(input)) {
        return [];
    }

    const result = [];

    for (let item of input) {
        if (typeof item === 'string') {
            const rangeMatch = item.match(/^([A-Za-z]+)-([A-Za-z]+)$/);
            if (rangeMatch) {
                const start = rangeMatch[1];
                const end = rangeMatch[2];
                result.push(...processRange(start, end));
            } else {
                result.push(...item.split(',').map(val => val.trim()));
            }
        } else {
            result.push(item);
        }
    }

    return result;
}

function processRange(start, end) {
    const result = [];
    const startCharCode = start.charCodeAt(0);
    const endCharCode = end.charCodeAt(0);

    for (let charCode = startCharCode; charCode <= endCharCode; charCode++) {
        result.push(String.fromCharCode(charCode));
    }
    return result;
}

function saveTextToFile(dataMap, fileName, directoryPath = './') {
    const filePath = `${directoryPath}/${fileName}.txt`;

    // Convert the Map to a string representation because there is an error if the data object is a map
    const dataString = Array.from(dataMap).map(([key, value]) => `${key}: ${value}`).join('\n');

    fs.writeFile(filePath, dataString, (err) => {
        if (err) {
            console.error('Error writing to file:', err);
        } else {
            console.log(`Data saved to file '${fileName}.txt' successfully.`);
        }
    });
}


// Example usage:

let directoryPath = './messageFiles'; // Specify the directory path where you want to save the file

//saveTextToFile(tb, fileName, directoryPath);

// Call houseAndBill with your parameters
const filePath = "C:/Users/Lenovo/Documents/GIGGAT_GARDENS/Water_Bills_2023_GIGGAT.xlsx";
const columns = ['B', 'H']; // Example array of letters representing columns
const rows = '3-19'; // Example row range
const worksheet = "December";

// Use async/await to wait for houseAndBill to finish execution
(async () => {
    await houseAndBill(filePath, columns, rows, worksheet);
    // Now `tb` is accessible globally
    //data will be tb
    saveTextToFile(tb,fileName, directoryPath)
    console.log(tb);

})();
