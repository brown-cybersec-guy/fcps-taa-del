function TAAURL_DELMod() {
  // Open the active spreadsheet and get the necessary sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var delSheet = ss.getSheetByName("DEL Extract Aug 21"); // Sheet containing the DEL data
  var taaSheet = ss.getSheetByName("To remove from TAA"); // Sheet containing the TAA data to check for duplicates

  // Get the data ranges
  var delRange = delSheet.getRange(2, 29, delSheet.getLastRow() - 1, 1); // Range for Column AC (Tech Assess Report Link) in DEL sheet
  var taaValues = taaSheet.getRange(2, 6, taaSheet.getLastRow() - 1, 1).getValues(); // Get values from Column F in TAA sheet

  // Get the values from DEL sheet's Column AC
  var delValues = delRange.getValues();
  
  // Initialize counters for logging the number of deletions
  var noTechlabReportCount = 0; // Counter for "noTechlabReport.html" deletions
  var tpaappsCount = 0; // Counter for "tpaapps.fcps.edu" deletions
  var taaDuplicateCount = 0; // Counter for duplicates found in TAA

  // Step 1: Delete values with "noTechlabReport.html" or "tpaapps.fcps.edu" in the URL
  for (var i = 0; i < delValues.length; i++) {
    var cellValue = delValues[i][0]; // Get the current cell value in Column AC

    // Check if the cell contains "noTechlabReport.html"
    if (cellValue.indexOf("noTechlabReport.html") !== -1) {
      delValues[i][0] = ""; // Delete the value by setting it to an empty string
      noTechlabReportCount++; // Increment the count for "noTechlabReport.html"
    } 
    // Check if the cell contains "tpaapps.fcps.edu"
    else if (cellValue.indexOf("tpaapps.fcps.edu") !== -1) {
      delValues[i][0] = ""; // Delete the value by setting it to an empty string
      tpaappsCount++; // Increment the count for "tpaapps.fcps.edu"
    }
  }
  
  // Update the DEL sheet with the modified values (after deletions)
  delRange.setValues(delValues);

  // Change all blank cells in Column AC to yellow
  var colorRange = delRange.getBackgrounds(); // Get the current background colors
  for (var j = 0; j < delValues.length; j++) {
    if (delValues[j][0] === "") {
      colorRange[j][0] = "#FFFF00"; // Set background to yellow for blank cells
    }
  }
  delRange.setBackgrounds(colorRange); // Apply the new background colors

  // Log the deletion counts for specific URLs
  Logger.log("Deleted " + noTechlabReportCount + " entries containing 'noTechlabReport.html'.");
  Logger.log("Deleted " + tpaappsCount + " entries containing 'tpaapps.fcps.edu'.");

  // Step 2: Check for duplicates with "To remove from TAA" and delete them
  var delACValues = delRange.getValues(); // Reload the values after the first deletion process
  for (var k = 0; k < delACValues.length; k++) {
    for (var l = 0; l < taaValues.length; l++) {
      // Check if there is a matching value between DEL and TAA
      if (delACValues[k][0] === taaValues[l][0]) {
        delACValues[k][0] = ""; // Delete the duplicate value by setting it to an empty string
        colorRange[k][0] = "#FF0000"; // Set the background color to red for deleted duplicates
        taaDuplicateCount++; // Increment the duplicate count
        break; // Exit the inner loop once a match is found to avoid unnecessary checks
      }
    }
  }

  // Update the DEL sheet with the modified values and new background colors after duplicate removal
  delRange.setValues(delACValues);
  delRange.setBackgrounds(colorRange);

  // Log the deletion count for duplicates found in TAA
  Logger.log("Deleted " + taaDuplicateCount + " duplicate entries found in 'To remove from TAA'.");
}
