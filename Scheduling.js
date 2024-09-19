function calculateSchedule() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if(data[i][3]=="Done") {
      continue;
    }

    if(isNumeric(data[i][0]) && data[i][2].toString().trim().length>0) {
      lastAssigneeIndex = getLastTaskIndexFor(data[i][2], i, 2, data);
      
      if(lastAssigneeIndex == -1) {
          var daysColumn = 8;     // Column H
          var targetColumn = 5 + 1;   // Column F
          var currentRow = i + 1; // Adjusting for 1-based indexing

          var daysOffset = daysColumn - targetColumn;  // 2 (H to F)

          // Construct the formula using R1C1 notation with absolute reference to C3 (R3C3)
          var formula = '=R3C3+IF(ISBLANK(R[0]C[' + daysOffset + ']),0,R[0]C[' + daysOffset + '])';

          // Set the formula into the target cell
          sheet.getRange(currentRow, targetColumn).setFormulaR1C1(formula);

          // Format the cell as a date
          // sheet.getRange(currentRow, targetColumn).setNumberFormat('yyyy-MM-dd');
      } else {
        indexDifference = (i - lastAssigneeIndex) * -1;
        // Set the formula in the cell, checking if R[0]C[2] is empty before using it
        sheet.getRange(i + 1, 5 + 1).setFormulaR1C1(
          "=R[" + indexDifference + "]C[1] + IF(ISBLANK(R[0]C[2]), 0, R[0]C[2]) + 1"
        );
      }

      sheet.getRange(i+1, 6+1).setFormulaR1C1("=WORKDAY(R[0]C[-1],R[0]C[-2],Harilibur)-1");

      Logger.log('Number: ' + data[i][0] + ' Task: ' + data[i][1] + ' Last Index: ' + lastAssigneeIndex);
    }
  }
}

function getLastTaskIndexFor(assignee, rowIndex, columnIndex, data) {
  // Slice the data array from the start to the specified rowIndex and reverse it
  const slicedData = data.slice(0, rowIndex).reverse();
  
  // Find the first occurrence of the assignee in the reversed data
  const reversedIndex = slicedData.findIndex(row => {
    // Check if the assignee matches and the value in column J (index 9) is not 1
    return row[columnIndex] === assignee && row[9] !== 1;
  });
  
  // If assignee was found, return the original index, else return -1
  return reversedIndex === -1 ? -1 : rowIndex - reversedIndex - 1;
}

function doCalculate(e) {
  var range = e.range;//The range of cells edited
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  const columnOfCellEdited = range.getColumn();//Get column number
  const rowOfCellEdited = e.range.getRow();

    // Regular expression to match sheet names starting with a number prefix
  const sheetNamePrefixRegex = /^\d+\./;

  if (!sheetNamePrefixRegex.test(sheetName)) {
    return;
  }

  const monitoredColumns = [1, 3, 5, 8, 10]; // Columns to be monitored

  // Check if the edited column is in the monitored columns array
  if (monitoredColumns.includes(columnOfCellEdited)) {
    calculateSchedule();
  }

  // Update status to gitlab
  if (columnOfCellEdited === 4) { // Column D
    const status = e.range.getValue(); // Get the new status from column D
    const issueUrlCell = sheet.getRange(rowOfCellEdited, 11); // Column K is column 11
    const issueUrl = issueUrlCell.getValue();

    if (issueUrl && status) {
      // Call the function to update the GitLab issue label
      updateGitLabIssueLabel(issueUrl, status);
    }
  }

  if (columnOfCellEdited === 11) { // Column K
    const issueUrl = e.range.getValue(); // Get the issue URL from column K
    if (issueUrl) {
      // Get the latest matching label from the issue
      const latestLabel = getLatestMatchingLabelFromIssue(issueUrl);

      if (latestLabel) {
        // Set the label into column D
        const statusCell = sheet.getRange(rowOfCellEdited, 4); // Column D is column 4
        statusCell.setValue(latestLabel);
      }
    }
  }


}
