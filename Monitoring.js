const monitoringSheetName = 'Resource Overview';

function prepareMonitoringSheet() {
  createNewSheet(monitoringSheetName);
  protectSheet(monitoringSheetName);
  moveSheetToIndex(monitoringSheetName, 0)
  setHeaders();
}

function refreshMonitoring() {
  copyTasksReference();
  createPivotTableCr();
  createPivotTablePic();
}

function copyTasksReference() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheet = ss.getSheetByName(monitoringSheetName);
  const activeSheet = ss.getActiveSheet(); // Store the current active sheet

  // Create the target sheet if it doesn't exist
  if (!targetSheet) {
    targetSheet = ss.insertSheet(monitoringSheetName);
  }

  // Start copying from row 1000 in the target sheet
  let targetRow = 1000;

  // Iterate through all sheets in the spreadsheet
  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();

    // Check if the sheet name starts with a number prefix
    if (SHEET_NAME_PREFIX_REGEX.test(sheetName)) {
      // Get the data range of the current sheet
      const dataRange = sheet.getDataRange();
      const values = dataRange.getValues();
      
      // Get the value from cell C2
      const valueFromC2 = sheet.getRange('C2').getValue();

      // Iterate through each row in the sheet
      values.forEach(row => {
        // Check if the first cell in the row is a valid number
        if (!isNaN(row[0]) && row[0] !== '') {
          // Get the status from column D (index 3)
          let status = row[3];
          if (status) {
            status = status.toString().trim().toLowerCase();
          } else {
            status = '';
          }

          // Check if status is not 'Done' or 'Backlog'
          if (status !== 'done' && status !== 'backlog') {
            // Append the value from C2 to the row
            row.push(valueFromC2);
            
            // Copy the row to the target sheet
            targetSheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
            targetRow++;
          }
        }
      });
    }
  });

  // Return the focus to the active sheet
  ss.setActiveSheet(activeSheet);

  Logger.log(`Rows excluding 'Done' and 'Backlog' statuses have been copied to "${monitoringSheetName}" starting from row 1000.`);
}


function setHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheet = ss.getSheetByName(monitoringSheetName);
  
  // Create the target sheet if it doesn't exist
  if (!targetSheet) {
    targetSheet = ss.insertSheet(monitoringSheetName);
  }
  
  // Define the headers
  const headers = [
    'No', 'Task', 'PIC', 'Status', 'Duration', 
    'Start Date', 'End Date', 'Skip', 
    'Skip Reason', 'Parallel', 'Issue Link', 'CR'
  ];

  // Set the headers in row 999
  targetSheet.getRange(999, 1, 1, headers.length).setValues([headers]);

  Logger.log('Headers have been set on row 999 in the sheet: ' + monitoringSheetName);
}

function createPivotTableCr() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(monitoringSheetName);

  // Define the range to place the pivot table (you can change 'M1' if needed)
  const pivotRange = sheet.getRange('A1');

  // Remove the pivot table starting at cell A1
  const existingPivotTables = sheet.getPivotTables();
  existingPivotTables.forEach(pivot => {
    const anchorCell = pivot.getAnchorCell();
    if (anchorCell.getA1Notation() === 'A1') {
      pivot.remove();
    }
  });


  // Define the data range starting from row 999 to the last row and the last column
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const dataRange = sheet.getRange(999, 1, lastRow - 998, lastColumn);

  // Create the pivot table
  const pivotTable = pivotRange.createPivotTable(dataRange);

  // Group by "CR" (12th column), "PIC" (3rd column), and "Task" (2nd column)
  pivotTable.addRowGroup(12); // "CR" is in the 12th column (1-based index 12)
  pivotTable.addRowGroup(3);  // "PIC" is in the 3rd column (1-based index 3)
  pivotTable.addRowGroup(2);  // "Task" is in the 2nd column (1-based index 2)

  // Add values to display in the pivot table and set custom display names
  // For Duration (Sum) with custom header name "Total Duration"
  const durationPivotValue = pivotTable.addPivotValue(
    5, // "Duration" is in the 5th column
    SpreadsheetApp.PivotTableSummarizeFunction.SUM
  );
  durationPivotValue.setDisplayName('Duration'); // Set custom header name

  // For Start Date (Minimum) with custom header name "Earliest Start Date"
  const startDatePivotValue = pivotTable.addPivotValue(
    6, // "Start Date" is in the 6th column
    SpreadsheetApp.PivotTableSummarizeFunction.MIN
  );
  startDatePivotValue.setDisplayName('Start Date');

  // For End Date (Maximum) with custom header name "Latest End Date"
  const endDatePivotValue = pivotTable.addPivotValue(
    7, // "End Date" is in the 7th column
    SpreadsheetApp.PivotTableSummarizeFunction.MAX
  );
  endDatePivotValue.setDisplayName('End Date');

  Logger.log('Pivot table created successfully from row 999.');
}

function createPivotTablePic() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(monitoringSheetName);

  // Define the range to place the pivot table (you can change 'M1' if needed)
  const pivotRange = sheet.getRange('H1');

  // Remove any existing pivot tables from the sheet
  const existingPivotTables = sheet.getPivotTables();
  existingPivotTables.forEach(pivot => {
    const anchorCell = pivot.getAnchorCell();
    if (anchorCell.getA1Notation() === 'H1') {
      pivot.remove();
    }
  });


  // Define the data range starting from row 999 to the last row and the last column
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const dataRange = sheet.getRange(999, 1, lastRow - 998, lastColumn);

  // Create the pivot table
  const pivotTable = pivotRange.createPivotTable(dataRange);

  // Group by "CR" (12th column), "PIC" (3rd column), and "Task" (2nd column)
  pivotTable.addRowGroup(3);  // "PIC" is in the 3rd column (1-based index 3)
  pivotTable.addRowGroup(12); // "CR" is in the 12th column (1-based index 12)
  pivotTable.addRowGroup(2);  // "Task" is in the 2nd column (1-based index 2)

  // Add values to display in the pivot table and set custom display names
  // For Duration (Sum) with custom header name "Total Duration"
  const durationPivotValue = pivotTable.addPivotValue(
    5, // "Duration" is in the 5th column
    SpreadsheetApp.PivotTableSummarizeFunction.SUM
  );
  durationPivotValue.setDisplayName('Duration'); // Set custom header name

  // For Start Date (Minimum) with custom header name "Earliest Start Date"
  const startDatePivotValue = pivotTable.addPivotValue(
    6, // "Start Date" is in the 6th column
    SpreadsheetApp.PivotTableSummarizeFunction.MIN
  );
  startDatePivotValue.setDisplayName('Start Date');

  // For End Date (Maximum) with custom header name "Latest End Date"
  const endDatePivotValue = pivotTable.addPivotValue(
    7, // "End Date" is in the 7th column
    SpreadsheetApp.PivotTableSummarizeFunction.MAX
  );
  endDatePivotValue.setDisplayName('End Date');

  Logger.log('Pivot table created successfully from row 999.');
}

function setDataValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // Get the named range 'WorkflowLabels'
  const namedRange = ss.getRangeByName(labelNamedRange);

  if (!namedRange) {
    Logger.log('Named range "WorkflowLabels" not found.');
    return;
  }

    // Regular expression to match sheet names starting with a number prefix
  const SHEET_NAME_PREFIX_REGEX = /^\d+\./;

  // Iterate over all sheets in the spreadsheet
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();

    // Check if the sheet name starts with a number prefix
    if (SHEET_NAME_PREFIX_REGEX.test(sheetName)) {

      // Get the last row with data in the sheet
      const lastRow = sheet.getLastRow();

      // If the sheet is empty beyond the header row, skip it
      if (lastRow < 6) return;

      // Define the range in column D (column index 4), starting from row 6 to the last row
      const range = sheet.getRange(6, 4, lastRow - 1);

      // Create the data validation rule using the named range
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(namedRange, true)
        .setAllowInvalid(false)
        .build();

      // Set the data validation rule for the range
      range.setDataValidation(rule);
    }
  });

  Logger.log('Data validation has been set for column D in all sheets.');
}

function setConditionalFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // Regular expression to match sheet names starting with a number prefix
  const SHEET_NAME_PREFIX_REGEX = /^\d+\./;

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();

    if (!SHEET_NAME_PREFIX_REGEX.test(sheetName)) {
      return;
    }

    // Get the last row with data in the sheet
    const lastRow = sheet.getLastRow();

    // If the sheet is empty beyond the header row, skip it
    if (lastRow < 6) return;

    // Calculate the number of rows to include in the range
    const numRows = lastRow - 6 + 1; // number of rows from row 6 to lastRow inclusive

    // Define the range in column G (column index 7), starting from row 6 to the last row
    const range = sheet.getRange(6, 7, numRows);

    // Define the range in column D (column index 4), starting from row 6 to the last row
    const rangeStatus = sheet.getRange(6, 4, numRows);

    // Get existing conditional format rules
    const rules = sheet.getConditionalFormatRules();

    // Remove existing rules that apply to column G
    const newRules = rules.filter(rule => {
      const ruleRanges = rule.getRanges();
      // Check if the rule applies to column G
      const appliesToColumnG = ruleRanges.some(ruleRange => {
        const ruleStartCol = ruleRange.getColumn();
        const ruleEndCol = ruleRange.getLastColumn();
        // Return true if the rule applies to column G
        return ruleStartCol <= 7 && ruleEndCol >= 7;
      });
      // Keep the rule only if it does not apply to column G
      return !appliesToColumnG;
    });

    // Define the statuses to exclude
    const statusesToExclude = ['Done', 'Ready to Merge', 'Ready for Deployement', 'Ready to Test', 'Ready to Implement']; // Add more statuses as needed

    // Convert the array of statuses to a string that can be used in the formula
    const statusesList = statusesToExclude.map(status => `"${status}"`).join(",");

    // Build the formula using MATCH and ISNA
    const formula = `=AND(ISNA(MATCH($D6, {${statusesList}}, 0)), $G6<TODAY(), NOT(ISBLANK($G6)))`;

    // Create the conditional formatting rule for red background (overdue dates)
    const redRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground('#FF0000') // Red color
      .setRanges([range])
      .build();

    // Modify the formula for yellow background (dates within next 3 days)
    const formulaYellow = `=AND(ISNA(MATCH($D6, {${statusesList}}, 0)), $G6>=TODAY(), $G6<=TODAY()+3, NOT(ISBLANK($G6)))`;

    // Create the conditional formatting rule for yellow background
    const yellowRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formulaYellow)
      .setBackground('#FFFF00') // Yellow color
      .setRanges([range])
      .build();

    // Create the conditional formatting rule for status Done
    const doneRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Done')
      .setBackground('#d9ead3')
      .setRanges([rangeStatus])
      .build();

    // Create the conditional formatting rule for status Testing Notes
    const testingNotesRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Testing Notes')
      .setBackground('#ea9999')
      .setRanges([rangeStatus])
      .build();

    // Create the conditional formatting rule for status Ready to Implement
    const readyToImplementRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Ready to Implement')
      .setBackground('#ffe599')
      .setRanges([rangeStatus])
      .build();
    
    // Create the conditional formatting rule for status Ready to Test
    const readyToTestRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Ready to Test')
      .setBackground('#f1c232')
      .setRanges([rangeStatus])
      .build();
    
    // Create the conditional formatting rule for status In Progress
    const inProgressRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('In Progress')
      .setBackground('#f9cb9c')
      .setRanges([rangeStatus])
      .build();

    // Add the new rules to the list
    newRules.push(redRule);
    newRules.push(yellowRule);
    newRules.push(doneRule);
    newRules.push(testingNotesRule);
    newRules.push(readyToImplementRule);
    newRules.push(readyToTestRule);
    newRules.push(inProgressRule);

    // Set the updated rules back to the sheet
    sheet.setConditionalFormatRules(newRules);
  });

  Logger.log('Conditional formatting has been set for column G in all sheets.');
}
