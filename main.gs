var COOKIE='m5qoh833rr0iceb3imptb46far';
var SHEET_ID='1ExhjTg4mEgDspsbTgppD2OYq2juSnumnQJStyF5RWps';

function fetchAndProcessMusicalList() {
  var endpointUrl = 'https://musical.congregacao.org.br/grp_musical/listagem';

  var endpointOptions = {
    'method': 'post',
    'headers': {
      'Cookie': `PHPSESSID=${COOKIE}` // Set the session cookie value
    }
  };

  // Make a request to the other endpoint using the session cookie
  var endpointResponse = UrlFetchApp.fetch(endpointUrl, endpointOptions);

  // Process the response from the other endpoint
  var musiciansData = JSON.parse(endpointResponse.getContentText());

  // Call the function to add the data to the spreadsheet
  addOrUpdateDataToSheet(musiciansData);
}

function addOrUpdateDataToSheet(data) {
  // Open the spreadsheet by ID
  var spreadsheet = SpreadsheetApp.openById(SHEET_ID);

  // Get the current month and year
  var today = new Date();
  var sheetName = Utilities.formatDate(today, spreadsheet.getSpreadsheetTimeZone(), 'MMM-yyyy'); // Format: JAN-2024

  // Check if the sheet already exists
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet) {
    // Clear content from the second row onward
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
  } else {
    // Create a new sheet with the name of the current month and year
    sheet = spreadsheet.insertSheet(sheetName);
  }

  // Add data to the sheet starting from the second row
  sheet.getRange(2, 1, data.data.length, data.data[0].length).setValues(data.data);
  
  // Sort the data by the ID column (column 1), starting from the second row
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).sort({column: 1, ascending: true});

  // Check if a filter already exists before creating a new one
  var existingFilter = sheet.getFilter();
  if (!existingFilter) {
    // Add a filter to the sheet
    sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).createFilter();
  }
}
