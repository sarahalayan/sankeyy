
const express = require("express");
const bodyParser = require("body-parser");
const xlsx = require("xlsx");
const ExcelJS = require("exceljs");
const moment = require("moment");
const cors = require('cors')
const app = express();
app.use(bodyParser.json());
app.use(cors())
 
app.post('/filter', (req, res) => {
    const { startDate, endDate } = req.body;
   
    const startDateTime = moment(startDate);
    const endDateTime = moment(endDate);
    console.log({ startDateTime, endDateTime });
    
    (async () => {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(
        'C:\\Users\\USER\\Desktop\\sankey+node\\2023_data.xlsx'
      );
      const worksheet = workbook.worksheets[0];
    
      let startRowNumber = null;
    
      let filteredDates = [];
      let columns = worksheet.getRow(1).values;
      columns = columns.slice(1, columns.length);
      let start = moment();
    
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 0) {
          // Skip the first row (index 0)
          return;
        }
    
        try {
          let date = moment(row.getCell(1).value?.toISOString());
          if (
            date.isBetween(moment(startDateTime), moment(endDateTime), "()", "[]")
          ) {
            let obj = {};
    
            for (let idx = 0; idx < columns.length; idx++) {
              const element = columns[idx];
              obj[element] = row.getCell(idx + 1).value;
            }
            filteredDates.push(obj);
          }
        } catch (error) {
          //   console.log(`Invalid date format in row ${rowNumber + 1}`);
        }
      });
      let end = moment();
    console.log(filteredDates);
      //const result = sumObjects(filteredDates);
    const result=sumConsecutiveDifferences(filteredDates);
      console.log({ "process take ": end.milliseconds() - start.milliseconds() });
      console.log({ result });
    
      res.json(result);
    })();
    
  
  });
  
   
app.listen(5500, () => {
    console.log('Server listening on port 5500');
});


function sumObjects(arr) {
   const sum = {};
   arr.forEach((obj) => {
   Object.keys(obj).forEach((key) => {
   if (typeof obj[key] === "number") {
  if (!sum.hasOwnProperty(key)) {
   sum[key] = 0;
  }
  
 sum[key] += obj[key];
   }
   });
   });
  
   return sum;
  }
  function sumConsecutiveDifferences(arr) {
    const sumDifferences = {};
    arr.forEach((obj, index) => {
        Object.keys(obj).forEach((key) => {
            if (typeof obj[key] === "number") {
                if (!sumDifferences.hasOwnProperty(key)) {
                    sumDifferences[key] = 0;
                }
                if (index > 0) {
                    const diff = obj[key] - arr[index - 1][key];
                    sumDifferences[key] += diff;
                }
            }
        });
    });
    
    return sumDifferences;
}

 