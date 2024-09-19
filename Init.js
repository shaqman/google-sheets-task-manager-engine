const holidaySheetName = 'Holidays';
const labelSheetName = 'Labels';
const labelNamedRange = 'WorkflowLabels';

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp, SlidesApp or FormApp.
  ui.createMenu('Peentar')
      .addItem('Update Holiday Data', 'importHolidays')
      .addItem('Update Labels', 'populateLabels')
      .addSeparator()
      .addItem('Update Monitoring', 'refreshMonitoring')
      // .addSubMenu(ui.createMenu('Sub-menu')
      //     .addItem('Second item', 'menuItem2'))
      .addToUi();

  defaultState(holidaySheetName);
  defaultState(labelSheetName);
  prepareMonitoringSheet();
  setDataValidation();
  setConditionalFormatting();
}

function importHolidays() {
  const calendarId = 'en.indonesian#holiday@group.v.calendar.google.com'; // Calendar ID from the link you provided
  const timeZone = 'Asia/Jakarta';
  const startDate = new Date(); // You can adjust the range of dates here if needed
  const endDate = new Date(new Date().setFullYear(startDate.getFullYear() + 1)); // Get holidays for one year ahead

  // Get events from the calendar
  const events = CalendarApp.getCalendarById(calendarId).getEvents(startDate, endDate);

  // Create a new sheet or clear the "Holidays" sheet if it exists
  const sheetName = holidaySheetName;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (sheet) {
    sheet.clear(); // Clear existing data
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  // Set headers
  sheet.getRange('A1').setValue('Date');
  sheet.getRange('B1').setValue('Event');

  // Loop through events and add to sheet
  const data = [];
  for (const event of events) {
    const eventDate = event.getStartTime();
    const eventTitle = event.getTitle();
    data.push([eventDate, eventTitle]);
  }

  // Insert the data into the sheet starting from row 2
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, 2).setValues(data);
  }

  // Format the date column
  sheet.getRange(2, 1, data.length, 1).setNumberFormat('yyyy-mm-dd');
  
  // Auto resize columns
  sheet.autoResizeColumns(1, 2);

  // Define the named range "Harilibur" for the date column (from A2 to the last row with data)
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(2, 1, lastRow - 1, 1); // The date column (A2:A)
  ss.setNamedRange('Harilibur', range); // Create the named range

  defaultState(sheetName);
}

function defaultState(sheetName) {
  protectSheet(sheetName);
  hideSheet(sheetName);
}


function populateLabels() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(labelSheetName);
  
  // Create the sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(labelSheetName);
  } else {
    sheet.clear(); // Clear existing content
  }

  // Set the header in cell A1
  sheet.getRange('A1').setValue('Workflow Label');

  // Write labels to the sheet starting from cell A2
  sheet.getRange(2, 1, labels.length, 1).setValues(labels.map(label => [label]));

  // Create a named range for the labels (from A2 to the last label)
  const range = sheet.getRange(2, 1, labels.length, 1);
  ss.setNamedRange(labelNamedRange, range);

  defaultState(labelSheetName);

  Logger.log('Labels have been populated and named range "WorkflowLabels" has been created.');
}
