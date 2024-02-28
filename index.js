const excel = require('exceljs');
let tb = new Map();//test will delete after


async function houseAndBill(filePath, columns, rows, worksheet) {
    const wb = new excel.Workbook();

    columns = processColumn(columns);
    const { start, end } = processRows(rows);

    try {
        await wb.xlsx.readFile(filePath);
        const ws = wb.getWorksheet(worksheet);

        iterateColumns(ws, start, end, columns);
        console.log(tb); // Log the map containing all cell values
        return tb;
    } catch (err) {
        console.error("Error:", err.message);
    }
}

function iterateColumns(ws, start, end, lettersArray) {
    for (let rowIndex = start; rowIndex <= end; rowIndex++) {
        let a,b;
        for (let letter of lettersArray) {
            const cell = ws.getCell(`${letter}${rowIndex}`);
            let cellValue;
            
            if (typeof cell.value === 'string') {
                // Handle string values
                cellValue = cell.value;
            } else if (typeof cell.value === 'object' && 'result' in cell.value) {
                // Handle numeric values with a result property
                cellValue = cell.value.result;
            } else {
                // Handle other types of values (e.g., undefined)
                cellValue = cell.value;
            }

            if(letter  ==='B'){
                cellValue=String(cellValue)
                 a = 'House : '+ cellValue
            }
            if(letter  ==='H'){
                b =  cellValue
            }

            tb.set(a,b)     

        }
    }
}


function processRows(input) {
    let start, end;

    // Check if input is a string
    if (typeof input === 'string') {
        // Check if input is a range (e.g., "3-19")
        const rangeMatch = input.match(/^(\d+)-(\d+)$/);
        if (rangeMatch) {
            start = parseInt(rangeMatch[1]);
            end = parseInt(rangeMatch[2]);
        }
    }

    // Return start and end as separate variables
    return { start, end };
}

function processColumn(input) {
    if (!Array.isArray(input)) {
        return []; // Invalid input, return an empty array
    }

    const result = [];

    // Iterate over each item in the input array
    for (let item of input) {
        // Check if the item is a string
        if (typeof item === 'string') {
            // Check if the string is a range (e.g., "A-C")
            const rangeMatch = item.match(/^([A-Za-z]+)-([A-Za-z]+)$/);
            if (rangeMatch) {
                const start = rangeMatch[1];
                const end = rangeMatch[2];
                result.push(...processRange(start, end));
            } else {
                // Otherwise, treat the string as specific values (e.g., "A,B,C")
                result.push(...item.split(',').map(val => val.trim()));
            }
        } else {
            // If the item is not a string, assume it's a single value
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
        // Do something with each letter in the range
    }
    return result;
}

const filePath = "C:/Users/Lenovo/Documents/GIGGAT_GARDENS/Water_Bills_2023_GIGGAT.xlsx";
const columns = ['B', 'H']; // Example array of letters representing columns
const rows = '3-19'; // Example row range
const worksheet = "December";

houseAndBill(filePath, columns, rows, worksheet);

