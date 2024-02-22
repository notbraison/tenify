// file reading
//await workbook.xlsx.readFile("C:\Users\Lenovo\Documents\GIGGAT GARDENS\Water_Bills_2023_GIGGAT.xlsx");

const excel = require('exceljs');
const path = require('path');



//retrieve housenumber and waterBill 
function houseAndBill(filePath,columns,rows,worksheet){
 

    const wb = new excel.Workbook();
    

    rows = processInput(rows)
    columns=processColumn(columns)
    
    wb.xlsx.readFile(filePath)
        .then(() => {
            const ws = wb.getWorksheet(worksheet);
    
             for (let rowIndex = start; rowIndex <= end; rowIndex++) { // Column index starts from 1, H corresponds to 8th column
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

}

//handle input that could be either a range, a single value, or specific values, 
//you can create a function that accepts a flexible input format and processes it accordingly
//for both string and numerical input it checks the type of input and processes it accordingly
function processInput(input) {
    // Check if input is a string
    if (typeof input === 'string') {
        // Check if input is a range (e.g., "3-19")
        const rangeMatch = input.match(/^(\d+)-(\d+)$/);
        if (rangeMatch) {
            const start = parseInt(rangeMatch[1]);
            const end = parseInt(rangeMatch[2]);
            return start, end;
        }
                
        // Otherwise, treat input as specific values (e.g., "3,5,7")
        return input.split(',').map(val => parseInt(val));
    }
    
    // Check if input is a single number
    if (!isNaN(input)) {
        return [parseInt(input)];
    }

    // Otherwise, invalid input
    return [];
}



// Example usage
/* console.log(processInput('3-19')); // Range: [3, 4, ..., 19]
console.log(processInput('5')); // Single value: [5]
console.log(processInput('3,5,7')); // Specific values: [3, 5, 7]
console.log(processInput(10)); // Single number: [10]
 */

function processColumn(input) {
    // Check if input is a string
    if (typeof input === 'string') {
        // Check if input is a range (e.g., "A-C")
        const rangeMatch = input.match(/^([A-Za-z]+)-([A-Za-z]+)$/);
        if (rangeMatch) {
            const start = rangeMatch[1];
            const end = rangeMatch[2];
            return processRange(start, end);
        }
        
        // Otherwise, treat input as specific values (e.g., "A,B,C")
        return input.split(',').map(val => val.trim());
    }
    // Check if input is not a string but an array or a single value
    else if (!Array.isArray(input)) {
        return [input];
    }
    
    // Otherwise, invalid input
    return [];
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

// Example usage

/* console.log(processColumn('A-C')); // Range: ['A', 'B', 'C']
console.log(processColumn('A,B,C')); // Specific values: ['A', 'B', 'C']
console.log(processColumn(['A'])); // Single value in an array: ['A']
console.log(processColumn('X')); // Single value: ['X']
console.log(processColumn('')); // Invalid input: [] */


const filePath = "C:/Users/Lenovo/Documents/GIGGAT_GARDENS/Water_Bills_2023_GIGGAT.xlsx";
    //columns=B,H
    //rows('3-19')
    const worksheet = "December"
houseAndBill(filePath,[B,H],[3,19],worksheet);