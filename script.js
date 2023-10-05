const XLSX = require('xlsx');
const fs = require('fs');

// Load the Excel file
const workbook = XLSX.readFile('Assignment_Timecard.xlsx');

// Assuming you want to analyze the first sheet
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

// Define variables to store results
const consecutiveDays = {};
const shortBreaks = {};
const longShifts = {};

// Helper function to check if two dates have the same date part
function isSameDate(date1, date2) {
  // Check if both dates are valid
  if (!(date1 instanceof Date) || isNaN(date1) || !(date2 instanceof Date) || isNaN(date2)) {
    return false;
  }
  return date1.toISOString().substring(0, 10) === date2.toISOString().substring(0, 10);
}

// Iterate through the rows of the sheet
const rows = XLSX.utils.sheet_to_json(sheet);
for (let i = 1; i < rows.length; i++) {
  const row = rows[i];
  const name = row['Employee Name'];
  const position = row['Position ID'];
  const time = new Date(row['Time']);
  const timeOut = new Date(row['Time Out']);
  const hoursWorked = (timeOut - time) / (1000 * 60 * 60); // Calculate hours worked

  // Check for consecutive days
  const prevRow = rows[i - 1];
  const prevTime = new Date(prevRow['Time']);
  if (isSameDate(time, prevTime)) {
    consecutiveDays[name] = (consecutiveDays[name] || 0) + 1;
  }

  // Check for short breaks
  const breakHours = (time - new Date(prevRow['Time Out'])) / (1000 * 60 * 60);
  if (breakHours > 1 && breakHours < 10) {
    shortBreaks[name] = (shortBreaks[name] || 0) + 1;
  }

  // Check for long shifts
  if (hoursWorked > 14) {
    longShifts[name] = (longShifts[name] || 0) + 1;
  }
}

// Print results
console.log('Employees with 7 consecutive days:');
console.log(consecutiveDays);

console.log('Employees with less than 10 hours between shifts:');
console.log(shortBreaks);

console.log('Employees with shifts longer than 14 hours:');
console.log(longShifts);

// You can write the results to a file if needed
fs.writeFileSync('output.json', JSON.stringify({ consecutiveDays, shortBreaks, longShifts }, null, 2));
