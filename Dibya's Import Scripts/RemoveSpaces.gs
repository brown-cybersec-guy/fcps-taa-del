function RemoveSpaces() {
  // Open the active spreadsheet and get the "Resource List" sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resourceListSheet = ss.getSheetByName("Resource List"); // This is the sheet we will be working on

  // Get the range of all cells with data in the "Resource List" sheet
  var range = resourceListSheet.getDataRange(); // This retrieves the entire range of data
  var values = range.getValues(); // Get the values from the range as a 2D array

  // Initialize counters and arrays for logging
  var totalSpacesRemoved = 0; // Counter for the total number of spaces removed
  var cellsModified = []; // Array to store logs of cells that were modified

  // Loop through each cell in the data range
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      var originalValue = values[i][j].toString(); // Convert the cell value to a string (in case it's a different data type)
      var trimmedValue = originalValue.trim(); // Trim spaces from the start and end of the cell value
      
      // Calculate the number of spaces removed by comparing lengths
      var spacesRemoved = originalValue.length - trimmedValue.length;

      if (spacesRemoved > 0) {
        // If spaces were removed, update the counters and logs
        totalSpacesRemoved += spacesRemoved; // Add to the total count of spaces removed
        cellsModified.push("Row " + (i + 1) + ", Column " + (j + 1) + ": " + spacesRemoved + " spaces removed"); // Log the modification
        values[i][j] = trimmedValue; // Update the value in the array with the trimmed value
      }
    }
  }

  // Update the sheet with the trimmed values
  range.setValues(values); // Set the modified values back into the sheet

  // Log the results of the operation
  if (cellsModified.length > 0) {
    // If any spaces were removed, log the total and details of each modified cell
    Logger.log("Total spaces removed: " + totalSpacesRemoved);
    cellsModified.forEach(function(cellLog) {
      Logger.log(cellLog); // Log each cell modification
    });
  } else {
    // If no spaces were removed, log that no changes were made
    Logger.log("No spaces were removed.");
  }
}
