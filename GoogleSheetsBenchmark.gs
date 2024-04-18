/*
Instructions:
1. Open your Google Spreadsheet.
2. Go to Extensions > Apps Script.
3. Delete any code in the script editor and replace it with the code below.
4. Save the script.
5. Close the script editor.
6. You can now run the benchmark function by selecting it from the toolbar and clicking the play button.

This script will execute the benchmarking process as described below:

A1 should say "Start Time:" - B1 should show the date and time that the script starts running
A2 should say "Progress..." - B2 should say "0 %"
A3 should say "Stop Time:" - B3 should say "In Progress"
A4 should say "Total Time:" - B4 should say "In Progress"

All of this should be written before the rest of the script runs.

D1 through K30 should be populated with random numbers between 9999999 and 99999999
Column M should multiply all of the results in the columns D through K in each row

Once this has been done, B2 should increase by 1 %
After this, all of the data from D1 through M30 as well as should be deleted.
This process should be repeated an additional 99 times.

Once this has been done, the current date and time should be displayed in B3

Set A4 to say "Total time:" - B4 should be how much time in seconds from the start time to the stop time.

Change the name of the spreadsheet to say the date and time then "Benchmark" then the total time shown in B4.
Here is an example of how it should Look: "2024-12-12 - 13:04 - Benchmark - 45 seconds"

All dates and time should follow this format: YYYY-MM-DD - HH:MM:SS
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
