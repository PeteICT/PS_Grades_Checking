function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Refresh Data', 'categorizeRows')
    .addToUi();
}

function categorizeRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName('Raw_Data');
  var excludeCourses = ss.getSheetByName('Exclude Courses').getDataRange().getValues();
  var excludeStudents = ss.getSheetByName('Exclude Students').getDataRange().getValues();

  // Process exclude lists
  excludeCourses.shift(); // Remove header
  excludeStudents.shift(); // Remove header
  var courseList = excludeCourses.map(function(r) { return r[0].toString().toLowerCase(); });
  var studentList = excludeStudents.map(function(r) { return r[0]; });

  var mainData = mainSheet.getDataRange().getValues();
  var headers = mainData.shift(); // Remove header row
  headers.push('Exclude Student'); // Add new column for checkbox

  // Create or get sheets for categorized data
  var missingSheets = {};
  for (var i = 1; i <= 5; i++) {
    var sheetName = 'Missing ' + i;
    missingSheets[i] = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    missingSheets[i].clear(); // Clear existing data
    missingSheets[i].appendRow(headers); // Add headers
  }

  // Check each row and categorize
  mainData.forEach(function(row) {
    if (courseList.some(function(course) { return row[9].toString().toLowerCase().includes(course); }) || studentList.includes(row[0])) {
      return; // Skip excluded courses and students
    }

    if (row[3] === '--') return; // Skip rows with "--" in Semester Grade

    var missingCount = 0;
    [3, 4, 5, 7, 8].forEach(function(index) { // Column indexes for Semester Grade, E, F, H, I
      if (!row[index] || isNaN(row[index])) {
        missingCount++;
      }
    });

    if (missingCount > 0 && missingCount <= 5) {
      var newRow = row.slice(); // Clone the row
      missingSheets[missingCount].appendRow(newRow);
      var lastRow = missingSheets[missingCount].getLastRow();
      var checkboxRange = missingSheets[missingCount].getRange(lastRow, headers.length); // Identify range for checkbox
      checkboxRange.insertCheckboxes(); // Insert checkbox
    }
  });
}

function checkCheckboxes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var excludeSheet = ss.getSheetByName('Exclude Students');
  for (var i = 1; i <= 5; i++) {
    var sheet = ss.getSheetByName('Missing ' + i);
    var data = sheet.getDataRange().getValues();
    var headers = data.shift(); // Remove header row
    var studentIndex = headers.length - 1; // Index of 'Exclude Student' column

    data.forEach(function(row, rowIndex) {
      if (row[studentIndex] === true) { // Checkbox is checked
        excludeSheet.appendRow([row[0]]); // Append Student No. to Exclude Students
        sheet.deleteRow(rowIndex + 2); // Delete the row, +2 due to header and 0-indexing
      }
    });
  }
}
