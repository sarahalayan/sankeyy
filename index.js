const xlsx = require('xlsx');
const express = require('express');
const bodyParser = require('body-parser');


/*const app = express();
app.use(bodyParser.json());
app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', 'file:///C:/Users/USER/Desktop/sankey+node/index.html'); // Replace with your allowed origin
    res.header('Access-Control-Allow-Headers', 'Content-Type');
    next();
});
 
app.post('/filter', (req, res) => {
    const { startDate, endDate } = req.body;
    console.log(startDate);
    const startDateObject = new Date(startDate);
    const endDateObject = new Date(endDate);

  
    // Read Excel data and calculate cumulative sums
    const workbook = xlsx.readFile("C:\\Users\\USER\\Downloads\\2023_data.xlsx");
    const sheet = workbook.Sheets['Sheet1']; // Replace with your sheet name

    let startAddress, endAddress; // Initialize variables to store matching addresses
    console.log(sheet[`A${35015}`].v);
    /*for (let row = 35015; row >= 35000; row--) { // Loop from last row to first
        const cellValue = sheet[`A${row}`].v;
        console.log(cellValue);

    // Compare dates (adjust comparison logic as needed)
    if (cellValue >= startDateObject && cellValue <= endDateObject) {
        if (!startAddress) {
        startAddress = row; // First matching address
        } else {
        endAddress = row; // Last matching address
        break; // Stop iterating once both addresses are found
        }
    }
    }

    console.log(startAddress);
    console.log(endAddress);




































    //console.log(sheet);
    // ... (logic to calculate cumulative sums for each column using SUMIFS)
    let columnSums = {}; // Initialize object to store cumulative sums
  
    /*for (const col in sheet) {
      if (col[0] === 'A') continue; // Skip header
      console.log(col);
  
      // Excel formula (adjust column range as needed)
      const cumulativeSumFormula = `=SUMIFS(${col}:${col},A:A,">="&DATE(${startDateObject.getFullYear()},${startDateObject.getMonth() + 1},${startDateObject.getDate()}),A:A,"<="&DATE(${endDateObject.getFullYear()},${endDateObject.getMonth() + 1},${endDateObject.getDate()}))`;
  
      // Retrieve cumulative sum and store in object
      columnSums[col] = sheet[`${col}${sheet['!ref'].split(':')[1]}`].v; // Assuming cumulative sum formula is in the last row
    }
  
    // Generate Sankey diagram data
    const sankeyData = [];
    for (const col in columnSums) {
      sankeyData.push({
        name: col, // Column name
        value: columnSums[col] // Cumulative sum for the column
      });
    }
  
    res.json(sankeyData);
  });
  
   
app.listen(8080, () => {
    console.log('Server listening on port 3000');
});*/

const ExcelJS = require('exceljs');
const datetime = require('luxon');

// Define start date with time (replace with your desired date)
const startDateTime = datetime.DateTime.fromISO('2023-07-08T20:45:00.000Z');

(async () => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile("C:\\Users\\USER\\Downloads\\2023_data.xlsx");
  const worksheet = workbook.worksheets[1];

  let startRowNumber = null;

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 0) {
      // Skip the first row (index 0)
      return;
    }

    const dateString = row.values[0]; // Assuming the first element is the date/time string

    try {
      // Convert dateString to string
      const dateStringAsText = dateString.toString();

      // Split the string until "T"
      const datePartString = dateStringAsText.split("T")[0];

      // Parse the date part
      const datePart = datetime.DateTime.fromISO(datePartString);

      // Extract the date part only
      const adjustedDatePart = datePart.startOf('day');

      if (adjustedDatePart.toString() !== startDateTime.toString()) {
        console.log(`Mismatch found in row: ${row}`);
      }
    } catch (error) {
      console.log(`Invalid date format in row ${rowNumber + 1}: ${dateString}`);
    }
  });

  if (startRowNumber !== null) {
    console.log(`Start row number: ${startRowNumber}`);
  } else {
    console.log("No matching row found for start date.");
  }
})();
