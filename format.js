function formatSheetOnNewFormEntry(e) {
  // Get the sheet that the function is being run on.
  var sheet = e.range.getSheet();

  // Calculate age and apply to column D
  calculateAge();

  // Get the range of cells that the function should format.
  var rangeToFormat = sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn());

  // Apply the desired formatting to the range of cells.
  rangeToFormat.setFontSize(12);
  rangeToFormat.setFontFamily("Arial");

  // Change the text wrapping of all cells to wrap.
  changeAllCellTextWrappingToWrap();

  // Align all cells to middle and center.
  alignAllCellsToMiddleAndCenter();

  // Save the changes to the sheet.
  SpreadsheetApp.flush();
}

function calculateAge() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getRange("D2:D" + sheet.getLastRow());
  var values = range.getValues();
  
  for (var i = 0; i < values.length; i++) {
    var dob = values[i][0];
    if (dob instanceof Date) {
      var today = new Date();
      var age = today.getFullYear() - dob.getFullYear();
      if (today.getMonth() < dob.getMonth() || (today.getMonth() == dob.getMonth() && today.getDate() < dob.getDate())) {
        age--;
      }
      values[i][0] = age;
    }
  }

  var outputRange = sheet.getRange("E2:E" + (values.length + 1));
  outputRange.setValues(values);
}

function onFormSubmit(e) {
  calculateAge();
}


function alignAllCellsToMiddleAndCenter() {
  // Get the range of all cells in the active sheet.
  var range = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange();

  // Set the horizontal alignment of all cells to center.
  range.setHorizontalAlignment('center');

  // Set the vertical alignment of all cells to middle.
  range.setVerticalAlignment('middle');
}

function changeAllCellTextWrappingToWrap() {
  // Get the active sheet.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Get all of the cells in the spreadsheet.
  var range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());

  // Apply a custom format to all cells to set the text wrapping to wrap.
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
}
