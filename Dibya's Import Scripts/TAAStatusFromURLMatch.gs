function TAAStatusFromURLMatch() {
  // Open the active spreadsheet and get the necessary sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resourceListSheet = ss.getSheetByName("Resource List"); // Sheet containing the Resource List data
  var taaExtractSheet = ss.getSheetByName("TAA extract Aug 21"); // Sheet containing the TAA extract data
  var itStatusLookupSheet = ss.getSheetByName("IT Status Lookup"); // Sheet containing the IT Status lookup data

  // Get the data from Resource List (Columns E to I)
  var resourceListData = resourceListSheet.getRange(2, 5, resourceListSheet.getLastRow() - 1, 5).getValues(); 

  // Get the data from TAA extract Aug 21 (Columns E and F)
  var taaExtractData = taaExtractSheet.getRange(2, 5, taaExtractSheet.getLastRow() - 1, 2).getValues(); 

  // Get the data from IT Status Lookup (Columns A and B)
  var itStatusLookupData = itStatusLookupSheet.getRange(2, 1, itStatusLookupSheet.getLastRow() - 1, 2).getValues(); 

  // Mapping of status descriptions to status IDs
  var statusMap = {
    "Meets": 1,
    "Complies": 2,
    "Cautiously Meets": 3,
    "Conditionally Meets": 3,
    "Does not Meet": 4,
    "Withheld": 8,
    "Fails": 9
  };

  // Create a map for IT Approval Status ID to IT Approval Status Description
  var itApprovalStatusMap = {};
  for (var i = 0; i < itStatusLookupData.length; i++) {
    itApprovalStatusMap[itStatusLookupData[i][0]] = itStatusLookupData[i][1]; // Map ID to description
  }

  var changesCount = 0; // Counter for the number of updates made
  var notFoundCount = 0; // Counter for the number of URLs not found
  var notFoundLog = []; // Array to log URLs that were not found

  // Loop through each row in the Resource List data
  for (var j = 0; j < resourceListData.length; j++) {
    var taaReportLink = resourceListData[j][4]; // Column I (TAA report link)
    var itApprovalStatusID = resourceListData[j][0]; // Column E (IT Approval Status ID)
    var matchFound = false; // Flag to check if a match was found

    // Process only if there's a TAA report link and Column E (IT Approval Status ID) is empty
    if (taaReportLink && itApprovalStatusID === "") { 
      for (var k = 0; k < taaExtractData.length; k++) {
        if (taaReportLink === taaExtractData[k][1]) { // Check if the report link matches in TAA extract
          var status = taaExtractData[k][0]; // Status from TAA extract (Column E)
          var statusID = statusMap[status]; // Get the corresponding status ID from the map

          if (statusID) {
            // Update IT Approval Status ID in Column E
            resourceListSheet.getRange(j + 2, 5).setValue(statusID);

            // Update IT Approval Status Description in Column F based on the lookup
            var itApprovalStatus = itApprovalStatusMap[statusID];
            if (itApprovalStatus) {
              resourceListSheet.getRange(j + 2, 6).setValue(itApprovalStatus);
            }

            changesCount++; // Increment the count of changes made
          }
          matchFound = true; // Set match found flag to true
          break; // Exit the inner loop once a match is found
        }
      }

      if (!matchFound) {
        // Highlight the cells in red if no match is found in TAA extract
        resourceListSheet.getRange(j + 2, 5, 1, 1).setBackground("#FF0000"); // Column E (IT Approval Status ID)
        resourceListSheet.getRange(j + 2, 6, 1, 1).setBackground("#FF0000"); // Column F (IT Approval Status)
        resourceListSheet.getRange(j + 2, 9, 1, 1).setBackground("#FF0000"); // Column I (TAA report link)
        notFoundLog.push(taaReportLink); // Log the TAA report link that was not found
        notFoundCount++; // Increment the count of not found entries
      }
    }
  }

  // Log the total number of entries updated
  Logger.log("Total entries updated: " + changesCount);
  
  // Log the URLs that were not found in TAA extract
  if (notFoundCount > 0) {
    Logger.log("Total entries not found: " + notFoundCount);
    notFoundLog.forEach(function(logEntry) {
      Logger.log("Not found: " + logEntry); // Log each URL that was not found
    });
  }
}
