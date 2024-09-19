function initializeSheetHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets(); // Get all sheets

  // Loop through each sheet
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();

    // Check if the sheet name starts with a number prefix
    if (SHEET_NAME_PREFIX_REGEX.test(sheetName)) {
      // Set headers and format
      setHeader(sheet);
    }
  });
}

function setHeader(sheet) {
  // Set values in header
  sheet.getRange('B2').setValue('Project Title');
  sheet.getRange('B3').setValue('Project Start');
  sheet.getRange('B4').setValue('Project End');
  sheet.getRange('C4').setFormula(`=MAX(G6:G10000)`);

  // Set column headers for row 5
  const columnHeaders = ['No', 'Task', 'PIC', 'Status', 'Duration', 'Start Date', 'End Date', 'Skip', 'Skip Reason', 'Parallel', 'Issue Link', 'Skip Holidays', 'Dependency'];
  
  sheet.getRange('A5:M5')
    .setValues([columnHeaders])
    .setFontWeight('bold')
    .setFontColor('white')
    .setBackground('#4a86e8') // Cornflower Blue color
    .setBorder(
      false, false, true, false, false, false,  // Top, Left, Bottom, Right, Vertical, Horizontal
      'black',                                  // Border color
      SpreadsheetApp.BorderStyle.MEDIUM       // Border style
    ).setHorizontalAlignment('center');

  sheet.getRange('A6:M6').setBackground('#a4c2f4'); // Light cornflower blue 2 color
}

