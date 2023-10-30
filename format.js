function formatSheetOnNewFormEntry(e) {
    // Get the sheet that the function is being run on.
    var sheet = e.range.getSheet();
  
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
    range.setFontSize(12);
    range.setFontFamily("Arial");
    range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }
  