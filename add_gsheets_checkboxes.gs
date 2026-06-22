/**
 * Personal Life Tracker 2026 — Google Sheets Checkbox Setup
 *
 * Run this script ONCE after uploading the xlsx to Google Sheets.
 * Go to: Extensions > Apps Script → paste this code → click Run addCheckboxes
 *
 * What it does: adds native click-to-toggle checkboxes (1=checked, 0=unchecked)
 * to all habit columns — Salat prayers, Gym, Supplements, and Running.
 */
function addCheckboxes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Checkbox data validation: click to toggle, 1 = checked, 0 = unchecked
  var rule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .withCriteria(SpreadsheetApp.DataValidationCriteria.CHECKBOX, [1, 0])
    .build();

  var sheets = {
    "Salat Tracker":   "C4:G368",   // Fajr, Dhuhr, Asr, Maghrib, Isha
    "Gym Routine":     "D4:D368",   // Workout Done
    "Supplements":     "C4:D368",   // Multivitamin + Omega
    "Running Tracker": "C4:C368",   // Ran Today
  };

  var count = 0;
  for (var name in sheets) {
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      sheet.getRange(sheets[name]).setDataValidation(rule);
      count++;
    } else {
      Logger.log("Sheet not found: " + name);
    }
  }

  SpreadsheetApp.getUi().alert(
    "Done! Native checkboxes added to " + count + " sheets.\n\n" +
    "Tap any red cell to check it green. Tap again to uncheck."
  );
}
