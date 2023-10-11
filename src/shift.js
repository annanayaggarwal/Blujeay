const ExcelJS = require('exceljs');
const fs = require('fs');

async function analyzeShifts(inputFile) {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(inputFile);

    // Assuming that the data is in the first sheet of the workbook
    const worksheet = workbook.getWorksheet(1);

    // Define sets to store employees meeting the criteria
    const consecutiveDays = new Set();
    const lessThan10Hours = new Set();
    const moreThan14Hours = new Set();

    let currentEmployee = '';
    let currentShifts = [];
    let consecutiveDaysCount = 0;

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) {
        // Skip the header row
        return;
      }

      const [_3,positionID, positionStatus, timeIn, timeOut, timecardHours, _, _1, employeeName, _2] = row.values;

      // console.log('Processing employee:', employeeName);


      if (employeeName !== currentEmployee) {
        // Analyze shifts for the current employee
        analyzeEmployeeShifts(currentEmployee, currentShifts, consecutiveDays, lessThan10Hours, moreThan14Hours);

        // Reset variables for the new employee
        currentEmployee = employeeName;
        currentShifts = [];
        consecutiveDaysCount = 0;
      }

      // Extract timecard hours as a number
      const timecardHoursNumber = parseFloat(timecardHours);

      // Check for consecutive days
      if (currentShifts.length > 1) {
        const currentDate = new Date(timeIn);
        const previousDate = new Date(currentShifts[currentShifts.length - 2].timeOut);
        const oneDay = 24 * 60 * 60 * 1000; // One day in milliseconds
        if ((currentDate - previousDate) === oneDay) {
          consecutiveDaysCount++;
          if (consecutiveDaysCount === 6) {
            consecutiveDays.add(currentEmployee);
          }
        } else {
          consecutiveDaysCount = 0;
        }
      }

      currentShifts.push({ timeIn, timeOut, timecardHoursNumber });
      // console.log('Shifts:', currentShifts);
      
    });

    // Analyze shifts for the last employee
    analyzeEmployeeShifts(currentEmployee, currentShifts, consecutiveDays, lessThan10Hours, moreThan14Hours);

    // Output the results
    console.log('Employees with more than 14 hours in a single shift:', Array.from(moreThan14Hours).join(', '));
    console.log('Employees with 7 consecutive days of work:', Array.from(consecutiveDays).join(', '));
    console.log('Employees with less than 10 hours between shifts:', Array.from(lessThan10Hours).join(', '));

    // Write the console output to output.txt
    const outputText = `
Employees with more than 14 hours in a single shift: ${Array.from(moreThan14Hours).join(', ')}
Employees with 7 consecutive days of work: ${Array.from(consecutiveDays).join(', ')}
Employees with less than 10 hours between shifts: ${Array.from(lessThan10Hours).join(', ')}
    `;
    fs.writeFileSync('output.txt', outputText);

  } catch (error) {
    console.error('Error:', error);
  }
}

function analyzeEmployeeShifts(name, shifts, consecutiveDays, lessThan10Hours, moreThan14Hours) {
  // Implement the logic for analyzing consecutive days, less than 10 hours, and more than 14 hours here
  for (let i = 0; i < shifts.length; i++) {
    const currentShift = shifts[i];
    const startDateTime = new Date(currentShift.timeIn);
    const endDateTime = new Date(currentShift.timeOut);
    const shiftHours = currentShift.timecardHours;

    if (shiftHours > 14) {
      moreThan14Hours.add(name);
    }

    if (i < shifts.length - 1) {
      const nextShift = shifts[i + 1];
      const currentEndTime = new Date(currentShift.timeOut);
      const nextStartTime = new Date(nextShift.timeIn);
      const hoursBetweenShifts = (nextStartTime - currentEndTime) / (60 * 60 * 1000);

      if (hoursBetweenShifts < 10 && hoursBetweenShifts > 1) {
        lessThan10Hours.add(name);
      }
    }
  }
}

// Usage: Call the function with the input XLSX file
analyzeShifts('input.xlsx');
