function protectSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    // Protect the sheet
    const protection = sheet.protect();
    protection.setDescription('This sheet is protected from editing.');
    
    // Allow only the owner of the spreadsheet to edit the sheet
    const me = Session.getEffectiveUser();
    protection.addEditor(me);
    protection.removeEditors(protection.getEditors());
    
    // Ensure the owner is not removed
    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
  }
}

// Function to unprotect the "Holidays" sheet
function unprotectSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    // Remove protection from the sheet
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (const protection of protections) {
      protection.remove();
    }
  }
}

// Function to hide the specified sheet
function hideSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    sheet.hideSheet();
  }
}

// Function to show the specified sheet
function showSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    sheet.showSheet();
  }
}

// Helper function to create a new sheet without changing focus
function createNewSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet(); // Store the current active sheet
  
  // Check if the sheet already exists
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    return sheet; // Return the existing sheet
  }

  // Create a new sheet if it doesn't exist
  sheet = ss.insertSheet(sheetName);

  // Switch back to the previously active sheet to keep focus there
  ss.setActiveSheet(activeSheet);
  
  return sheet;
}

// Helper function to move the specified sheet to a specific index without changing active sheet
function moveSheetToIndex(sheetName, targetIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet(); // Store the current active sheet

  const sheet = ss.getSheetByName(sheetName);
  ss.setActiveSheet(sheet);
  
  if (sheet) {
    const sheetIndex = sheet.getIndex(); // Get the current index of the sheet
    
    // Only move if it's not already at the target index
    if (sheetIndex !== targetIndex) {
      ss.moveActiveSheet(targetIndex); // Move the sheet to the target index
    }

  }

  ss.setActiveSheet(activeSheet);
}

function isNumeric(value) {
    return /^\d+$/.test(value);
}