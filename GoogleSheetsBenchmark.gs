/*
What is this?
This is a simple Google App Script that you can run in a Google sheet. 
It populates random numbers and multiplies them together. It does this 
100 times then gives a time to complete. 
This is useful for comparing computers to see how fast the CPU is for 
Calculating Google sheets processes.

Instructions:
1. Open your Google Spreadsheet.
2. Go to Extensions > Apps Script.
3. Delete any code in the script editor and replace it with the code below.
4. Save the script.
5. Run the script. Go to your Spreadsheet to check the progress.
*/

function benchmark() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Set Start Time
  sheet.getRange("B1").setValue(new Date()).setNumberFormat("yyyy-MM-dd - HH:mm:ss");
  sheet.getRange("B2").setValue("0 %");
  sheet.getRange("B3").setValue("In Progress");
  sheet.getRange("B4").setValue("In Progress");

  for (var i = 0; i < 100; i++) {
    // Populate D1:K30 with random numbers
    populateRandomNumbers(sheet);

    // Calculate product in column M
    calculateProduct(sheet);

    // Increase Progress
    increaseProgress(sheet);

    // Delete data in D1:M30
    clearData(sheet);
  }

  // Set Stop Time
  sheet.getRange("B3").setValue(new Date()).setNumberFormat("yyyy-MM-dd - HH:mm:ss");

  // Calculate Total Time
  calculateTotalTime(sheet);

  // Rename Spreadsheet
  renameSpreadsheet(sheet);
}

function populateRandomNumbers(sheet) {
  var range = sheet.getRange("D1:K30");
  var values = range.getValues();

  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      values[i][j] = Math.floor(Math.random() * 89999999) + 9999999;
    }
  }

  range.setValues(values);
}

function calculateProduct(sheet) {
  var range = sheet.getRange("M1:M30");
  range.setFormulaR1C1("=PRODUCT(RC[-9]:RC[-1])");
}

function increaseProgress(sheet) {
  var progress = parseInt(sheet.getRange("B2").getValue());
  sheet.getRange("B2").setValue((progress + 1) + " %");
}

function clearData(sheet) {
  var range = sheet.getRange("D1:M30");
  range.clear();
}

function calculateTotalTime(sheet) {
  var startTime = sheet.getRange("B1").getValue();
  var stopTime = sheet.getRange("B3").getValue();
  var totalTime = (stopTime - startTime) / 1000; // in seconds

  sheet.getRange("B4").setValue(totalTime + " seconds");
}

function renameSpreadsheet(sheet) {
  var startTime = sheet.getRange("B1").getValue();
  var totalTime = sheet.getRange("B4").getValue().split(" ")[0];
  var newName = Utilities.formatDate(startTime, Session.getScriptTimeZone(), "yyyy-MM-dd - HH:mm") + " - Benchmark - " + totalTime + " seconds";
  sheet.getParent().rename(newName);
}
