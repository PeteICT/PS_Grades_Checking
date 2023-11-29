function categorizeRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Raw_Data'); // Replace with your main sheet name
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Remove header row

  // Create or get sheets for categorized data
  var missingSheets = {};
  for (var i = 1; i <= 5; i++) {
    var sheetName = 'Missing ' + i;
    missingSheets[i] = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    missingSheets[i].clear(); // Clear existing data
    missingSheets[i].appendRow(headers); // Add headers
  }

  // Check each row and categorize
  data.forEach(function(row) {
    var missingCount = 0;
    [3, 4, 5, 7, 8].forEach(function(index) { // Column indexes for D, E, F, H, I
      if (!row[index] || isNaN(row[index])) {
        missingCount++;
      }
    });

    if (missingCount > 0 && missingCount <= 5) {
      missingSheets[missingCount].appendRow(row);
    }
  });
}
