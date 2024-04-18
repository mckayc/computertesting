// This is a benchmark script that populates random numbers then adds it together.
// Instructions for installing and running the script:
// 1. Open your Google Spreadsheet.
// 2. Go to Extensions > Apps Script.
// 3. Paste the following code into the script editor.
// 4. Save the project with a name.
// 5. Close the script editor.
// 6. Refresh the spreadsheet, and a new menu item "Run Benchmark" should appear.
// 7. Click on "Run Benchmark" to execute the benchmark test.

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Run Benchmark')
      .addItem('Start Benchmark', 'startBenchmark')
      .addToUi();
}

function startBenchmark() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Setting up initial data
  sheet.getRange("A1").setValue("Start Time:");
  sheet.getRange("B1").setValue(new Date());
  sheet.getRange("A2").setValue("Progress...");
  sheet.getRange("B2").setValue("0 %");
  sheet.getRange("A3").setValue("Stop Time:");
  sheet.getRange("B3").setValue("In Progress");
  sheet.getRange("A4").setValue("Total Time:");
  sheet.getRange("B4").setValue("In Progress");
  sheet.getRange("D1").setValue("Test Time");
  sheet.getRange("E1").setValue("All Tests Ave");

  // Loop 100 times
  for (var loop = 1; loop <= 100; loop++) {
    // Populating random numbers
    for (var i = 1; i <= 5; i++) {
      for (var j = 7; j <= 11; j++) {
        sheet.getRange(i, j).setValue(Math.floor(Math.random() * 1000));
      }
    }

    // Calculating row sums
    for (var i = 1; i <= 5; i++) {
      var sum = 0;
      for (var j = 7; j <= 11; j++) {
        sum += sheet.getRange(i, j).getValue();
      }
      sheet.getRange(i, 13).setValue(sum);
    }

    // Increase progress by 1%
    var currentProgress = parseInt(sheet.getRange("B2").getValue());
    sheet.getRange("B2").setValue((currentProgress + 1) + " %");
  }

  // Calculate and set average in E2
  sheet.getRange("E2").setFormula('=AVERAGE(E4:E)');

  // Set stop time
  var stopTime = new Date();
  sheet.getRange("B3").setValue(stopTime);
  var lastRow = sheet.getLastRow();

  // Copy stop time to next empty cell underneath D3
  sheet.getRange(lastRow + 1, 4).setValue(stopTime);

  // Calculate total time
  var startTime = sheet.getRange("B1").getValue();
  var totalTime = (stopTime - startTime) / 1000; // in seconds
  sheet.getRange("B4").setValue(totalTime + " seconds");

  // Copy total time to next empty cell underneath E3
  sheet.getRange(lastRow + 1, 5).setValue(totalTime);

  // Rename spreadsheet
  var formattedStartTime = Utilities.formatDate(startTime, Session.getScriptTimeZone(), "yyyy-MM-dd - HH:mm:ss");
  var spreadsheetName = formattedStartTime + " - Benchmark - " + totalTime + " seconds";
  SpreadsheetApp.getActiveSpreadsheet().rename(spreadsheetName);
}
