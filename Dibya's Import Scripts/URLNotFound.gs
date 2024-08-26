function URLNotFound() {
  // Open the active spreadsheet and get the "Resource List" sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resourceListSheet = ss.getSheetByName("Resource List");

  // Get the data from the Resource List sheet, specifically Columns A to I
  var resourceListData = resourceListSheet.getRange(2, 1, resourceListSheet.getLastRow() - 1, 9).getValues(); 

  // Create a new sheet named "TAA Report Link Invalid" to store invalid entries
  var invalidSheetName = "TAA Report Link Invalid";
  var invalidSheet = ss.getSheetByName(invalidSheetName);
  
  // If the "TAA Report Link Invalid" sheet already exists, delete it to avoid duplication
  if (invalidSheet) {
    ss.deleteSheet(invalidSheet); 
  }
  
  // Insert a new sheet with the name "TAA Report Link Invalid"
  invalidSheet = ss.insertSheet(invalidSheetName);

  // Set up the header row in the new sheet
  invalidSheet.getRange(1, 1).setValue("Resource ID"); // Column A
  invalidSheet.getRange(1, 2).setValue("Resource Name"); // Column B
  invalidSheet.getRange(1, 3).setValue("TAA Report Link"); // Column C

  var row = 2; // Initialize the row counter for writing to the new sheet

  // Iterate through the Resource List data
  for (var i = 0; i < resourceListData.length; i++) {
    var itApprovalStatusID = resourceListData[i][4]; // Column E (IT Approval Status ID)
    var itApprovalStatus = resourceListData[i][5];   // Column F (IT Approval Status)
    var taaReportLink = resourceListData[i][8];      // Column I (TAA Report Link)

    // Check if both Column E and F are blank, but Column I has content
    if (!itApprovalStatusID && !itApprovalStatus && taaReportLink) {
      // If the condition is met, add the entry to the "TAA Report Link Invalid" sheet
      invalidSheet.getRange(row, 1).setValue(resourceListData[i][0]); // Resource ID (Column A)
      invalidSheet.getRange(row, 2).setValue(resourceListData[i][1]); // Resource Name (Column B)
      invalidSheet.getRange(row, 3).setValue(taaReportLink); // TAA Report Link (Column C)
      row++; // Move to the next row for the next invalid entry
    }
  }
  
  // Log a message indicating that the script has completed successfully
  Logger.log("Script completed. 'TAA Report Link Invalid' sheet created with invalid entries.");
}
