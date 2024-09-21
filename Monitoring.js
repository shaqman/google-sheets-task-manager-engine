const monitoringSheetName = 'Resource Overview';

function prepareMonitoringSheet() {
  createNewSheet(monitoringSheetName);
  protectSheet(monitoringSheetName);
  moveSheetToIndex(monitoringSheetName, 0)
  setHeaders();
  createPivotTableCr();
  createPivotTablePic();
}

function refreshMonitoring() {
  copyTasksReference();
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

  // Clear existing data from row 1000 to the last row
  if (targetSheet.getLastRow() >= 1000) {
    const rangeToClear = targetSheet.getRange(1000, 1, targetSheet.getLastRow() - 999, targetSheet.getLastColumn());
    rangeToClear.clearContent();
  }

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
      const currentCrName = sheet.getRange('C2').getValue();

      // Iterate through each row in the sheet
      values.forEach(row => {
        // Check if the first cell in the row is a valid number
        if (!isNaN(row[0]) && row[0] !== '') {
          // Get the status from column D (index 3)
          let status = row[3].toString().trim();
          // Print each column value for debugging
          for (let i = 0; i < row.length && i < 14; i++) { // Up to column N (index 13)
            Logger.log(`Column ${String.fromCharCode(65 + i)}: ${row[i]}`);
          }
          if(status === "") return;
          
          // Define an array of statuses to exclude
          const excludedStatuses = ['done', 'backlog', 'ready to implement', 'ready to test']; // Add more statuses to exclude as needed

          // Check if the status is not in the excludedStatuses array
          if (!excludedStatuses.includes(status.toLowerCase())) {
              // Append the value from C2 to the row
              row.push(currentCrName);
              
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

  let extendedHeaders = [...COLUMN_HEADERS, 'CR'];
  
  // Set the headers in row 999
  targetSheet.getRange(999, 1, 1, extendedHeaders.length).setValues([extendedHeaders]);

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

  const dataRange = sheet.getRange(999, 1, DATA_END_ROW, sheet.getLastColumn());

  // Create the pivot table
  const pivotTable = pivotRange.createPivotTable(dataRange);

  // Group by "CR" (12th column), "PIC" (3rd column), and "Task" (2nd column)
  pivotTable.addRowGroup(sheet.getLastColumn()); // "CR" is in the 12th column (1-based index 12)
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
  const dataRange = sheet.getRange(999, 1, DATA_END_ROW, sheet.getLastColumn());

  // Create the pivot table
  const pivotTable = pivotRange.createPivotTable(dataRange);

  // Group by "CR" (12th column), "PIC" (3rd column), and "Task" (2nd column)
  pivotTable.addRowGroup(3);  // "PIC" is in the 3rd column (1-based index 3)
  pivotTable.addRowGroup(sheet.getLastColumn()); // "CR" is in the 12th column (1-based index 12)
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
      const appliesToColumnGorD = ruleRanges.some(ruleRange => {
        const ruleStartCol = ruleRange.getColumn();
        const ruleEndCol = ruleRange.getLastColumn();
        // Return true if the rule applies to column G or column D
        return (ruleStartCol <= 7 && ruleEndCol >= 7) || (ruleStartCol <= 4 && ruleEndCol >= 4);
      });
      // Keep the rule only if it does not apply to column G
      return !appliesToColumnGorD;
    });
    
    // Create and add the date-based rules
    TASKS_DATE_STATUS_RULES.forEach(({ formula, color }) => {
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(formula)
        .setBackground(color)
        .setRanges([range])
        .build();
      newRules.push(rule);
    });

    // Create and add the status-based rules
    STATUS_RULES.forEach(({ status, color }) => {
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(status)
        .setBackground(color)
        .setRanges([rangeStatus])
        .build();
      newRules.push(rule);
    });

    // Set the updated rules back to the sheet
    sheet.setConditionalFormatRules(newRules);
  });

  // Apply STATUS_RULES to Resource Overview sheet, columns F and M until row 900
  const resourceOverviewSheet = ss.getSheetByName(monitoringSheetName);

  if (resourceOverviewSheet) {
    // Get existing conditional format rules
    let rules = resourceOverviewSheet.getConditionalFormatRules();

    // Remove existing rules that apply to columns G and M
    rules = rules.filter(rule => {
      const ruleRanges = rule.getRanges();
      // Check if the rule applies to column G or M
      const appliesToColumnGorM = ruleRanges.some(ruleRange => {
        const ruleStartCol = ruleRange.getColumn();
        const ruleEndCol = ruleRange.getLastColumn();
        // Return true if the rule applies to column G or column M
        return (ruleStartCol <= 6 && ruleEndCol >= 6) || (ruleStartCol <= 13 && ruleEndCol >= 13);
      });
      // Keep the rule only if it does not apply to column G or M
      return !appliesToColumnGorM;
    });

    const rangeFStatus = resourceOverviewSheet.getRange('F2:F900');
    const rangeMStatus = resourceOverviewSheet.getRange('M2:M900');

    RESOURCE_DATE_RULES.forEach(({ formula, color }) => {
      const ruleF = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(formula.replace(/\$G6/g, '$F2'))
        .setBackground(color)
        .setRanges([rangeFStatus])
        .build();
      rules.push(ruleF);

      const ruleM = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(formula.replace(/\$G6/g, '$M2'))
        .setBackground(color)
        .setRanges([rangeMStatus])
        .build();
      rules.push(ruleM);
    });

    resourceOverviewSheet.setConditionalFormatRules(rules);
    Logger.log('Conditional formatting has been set for columns F and M in the Resource Overview sheet.');
  } else {
    Logger.log('Resource Overview sheet not found.');
  }

  Logger.log('Conditional formatting has been set for column G in all sheets.');
}
