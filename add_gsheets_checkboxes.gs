/**
 * ════════════════════════════════════════════════════════════
 *  Personal Life Tracker 2026 — Google Sheets Checkbox Setup
 * ════════════════════════════════════════════════════════════
 *
 *  HOW TO USE (one-time setup, takes ~30 seconds):
 *
 *  1. Upload Personal_Life_Tracker.xlsx to Google Drive
 *  2. Right-click the file → Open with → Google Sheets
 *     (it will convert the file automatically)
 *  3. In the new Sheets file: click  Extensions → Apps Script
 *  4. Delete everything in the editor, paste THIS entire file
 *  5. Click the ▶ Run button  (you may need to grant permission once)
 *  6. Click "Review permissions" → Advanced → Allow
 *  7. Run it again — a popup confirms checkboxes are added
 *  8. Close the script editor and enjoy tap-to-toggle checkboxes!
 *
 *  What it does:
 *  - Adds native click-to-toggle checkboxes to Salat, Gym,
 *    Supplements, Running, and Task Tracker sheets
 *  - Checked = 1 (green),  Unchecked = 0 (red)
 *  - Runs ONCE — checkboxes are saved permanently in the file
 * ════════════════════════════════════════════════════════════
 */

function addCheckboxes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Native checkbox: click to toggle between 1 (checked) and 0 (unchecked)
  var rule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .withCriteria(SpreadsheetApp.DataValidationCriteria.CHECKBOX, [1, 0])
    .build();

  var sheets = {
    "Salat Tracker":   "C4:G368",   // Fajr, Dhuhr, Asr, Maghrib, Isha
    "Gym Routine":     "D4:D368",   // Workout Done
    "Supplements":     "C4:D368",   // Multivitamin + Omega
    "Running Tracker": "C4:C368",   // Ran Today
    "Task Tracker":    "C4:C53",    // Check column (50 tasks)
  };

  var added = [];
  var missing = [];

  for (var name in sheets) {
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      sheet.getRange(sheets[name]).setDataValidation(rule);
      added.push(name);
    } else {
      missing.push(name);
    }
  }

  var msg = "✅  Native checkboxes added to:\n• " + added.join("\n• ");
  if (missing.length > 0) {
    msg += "\n\n⚠️  Sheets not found (check names):\n• " + missing.join("\n• ");
  }
  msg += "\n\nTap any red cell to check it green. Tap again to uncheck.";

  SpreadsheetApp.getUi().alert("Done!", msg, SpreadsheetApp.getUi().ButtonSet.OK);
}
