function ResourceID() {
  // Open the active spreadsheet and get the specific sheets needed for processing
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resourceListSheet = ss.getSheetByName("Resource List"); // Sheet containing the resource IDs and names
  var delExtractSheet = ss.getSheetByName("DEL Extract Aug 21"); // Sheet containing the extracted DEL data
  var issuesSheet = ss.getSheetByName("ISSUES"); // Sheet where unmatched entries will be logged

  // Clear any previous issues logged in the ISSUES sheet
  issuesSheet.clear();

  // Get data from the "Resource List" sheet, starting from the 2nd row to avoid headers
  // The range includes columns A (Resource ID), B (Resource Name), and C (Publish Status)
  var resourceListData = resourceListSheet.getRange(2, 1, resourceListSheet.getLastRow() - 1, 3).getValues();

  // Get data from the "DEL Extract Aug 21" sheet, starting from the 2nd row to avoid headers
  // The range includes columns A (Resource ID), B (Resource Name), and D (Publish Status)
  var delExtractData = delExtractSheet.getRange(2, 1, delExtractSheet.getLastRow() - 1, 4).getValues();

  // Create an object (map) to store Resource IDs, using a combination of Resource Name and Publish Status as the key
  var resourceMap = {};
  for (var i = 0; i < resourceListData.length; i++) {
    var resourceId = resourceListData[i][0]; // Resource ID from "Resource List"
    var resourceName = resourceListData[i][1]; // Resource Name from "Resource List"
    var publishStatus = resourceListData[i][2]; // Publish Status from "Resource List"
    var key = resourceName + "|" + publishStatus;  // Create a unique key by combining Resource Name and Publish Status
    resourceMap[key] = resourceId; // Store the Resource ID in the map with the key
  }

  // Initialize counters and arrays to track issues and unmatched entries
  var unmatchedCount = 0;
  var unmatchedEntries = [];

  // Iterate through each row in the "DEL Extract Aug 21" data to find matches in the resource map
  for (var j = 0; j < delExtractData.length; j++) {
    var delResourceName = delExtractData[j][1]; // Resource Name from "DEL Extract Aug 21"
    var delPublishStatus = delExtractData[j][3]; // Publish Status from "DEL Extract Aug 21"
    var key = delResourceName + "|" + delPublishStatus;  // Create a unique key for the DEL data

    // Check if the key exists in the resource map
    if (resourceMap[key]) {
      // If a match is found, set the corresponding Resource ID in the "DEL Extract Aug 21" sheet
      delExtractSheet.getRange(j + 2, 1).setValue(resourceMap[key]);
    } else {
      // If no match is found, log the issue in the "ISSUES" sheet
      issuesSheet.appendRow([delResourceName, delPublishStatus]); // Log the Resource Name and Publish Status
      unmatchedCount++; // Increment the counter for unmatched entries
      unmatchedEntries.push(delResourceName + " with Publish Status: " + delPublishStatus); // Add details to the array for logging
    }
  }

  // Log the results in the console
  if (unmatchedCount > 0) {
    Logger.log(unmatchedCount + " Resource Names were not matched:");
    unmatchedEntries.forEach(function(entry) {
      Logger.log(entry); // Log each unmatched entry for reference
    });
  } else {
    Logger.log("All Resource Names and Publish Status were successfully matched."); // Indicate that no issues were found
  }
}
