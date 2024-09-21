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
  sheet.getRange('C4')
    .setFormula('=MAX(G'+DATA_START_ROW+':G'+DATA_END_ROW+')')
    .setNumberFormat('dd/mm/yyyy');

  sheet.getRange('A'+ (DATA_START_ROW - 1) + ':M'+ (DATA_START_ROW - 1))
    .setValues([COLUMN_HEADERS])
    .setFontWeight('bold')
    .setFontColor('white')
    .setBackground('#4a86e8') // Cornflower Blue color
    .setBorder(
      false, false, true, false, false, false,  // Top, Left, Bottom, Right, Vertical, Horizontal
      'black',                                  // Border color
      SpreadsheetApp.BorderStyle.MEDIUM       // Border style
    ).setHorizontalAlignment('center');

  sheet.getRange('A'+DATA_START_ROW+':M'+DATA_START_ROW+'').setBackground('#a4c2f4'); // Light cornflower blue 2 color
}

