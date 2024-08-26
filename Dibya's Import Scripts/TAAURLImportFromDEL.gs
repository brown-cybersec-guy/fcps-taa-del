function TAAURLImportFromDEL() {
  // Open the active spreadsheet and get the necessary sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var delSheet = ss.getSheetByName("DEL Extract Aug 21"); // Sheet containing the DEL data
  var resourceListSheet = ss.getSheetByName("Resource List"); // Sheet containing the Resource List data

  // Get the data from the DEL Extract Aug 21 sheet, specifically Columns A (ResourceID) and AC (Tech Assess Report Link)
  var delData = delSheet.getRange(2, 1, delSheet.getLastRow() - 1, 29).getValues(); // Get all the necessary columns (A through AC)

  // Get the data from the Resource List sheet, specifically Columns A (ResourceID), I (TAA report link), and J (IT Approver)
  var resourceListData = resourceListSheet.getRange(2, 1, resourceListSheet.getLastRow() - 1, 10).getValues(); // Get columns A through J

  var copyCount = 0; // Counter to track how many URLs are copied

  // Create a map to associate Resource IDs with their row numbers in the Resource List sheet
  var resourceMap = {};
  for (var i = 0; i < resourceListData.length; i++) {
    var resourceId = resourceListData[i][0]; // ResourceID from the Resource List
    if (resourceId) {
      resourceMap[resourceId] = i + 2; // Store the row number (offset by 2 to account for header row) for later use
    }
  }

  // Iterate through the DEL Extract Aug 21 data to find matching Resource IDs in the Resource List
  for (var j = 0; j < delData.length; j++) {
    var delResourceId = delData[j][0]; // ResourceID from DEL Extract Aug 21
    var delTechAssessLink = delData[j][28]; // Tech Assess Report Link from DEL Extract Aug 21

    // Check if the ResourceID exists, the Tech Assess Report Link is not empty, and there is a corresponding entry in the Resource List
    if (delResourceId && delTechAssessLink && resourceMap[delResourceId]) {
      // If a match is found, copy the Tech Assess Report Link to the Resource List sheet in Column I
      resourceListSheet.getRange(resourceMap[delResourceId], 9).setValue(delTechAssessLink);
      // Set the IT Approver to "DAWatson" in Column J of the Resource List
      resourceListSheet.getRange(resourceMap[delResourceId], 10).setValue("DAWatson");
      copyCount++; // Increment the counter for each successful copy
    }
  }

  // Log the number of URLs that were successfully copied to the Resource List and the IT Approver updates
  Logger.log("Copied " + copyCount + " URLs to the Resource List and set IT Approver to 'DAWatson'.");
}
