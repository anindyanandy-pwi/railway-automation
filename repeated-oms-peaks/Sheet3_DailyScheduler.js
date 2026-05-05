// ============================================================
//  Sheet3_DailyScheduler.gs — Sheet 3
//  Individual daily triggers for OMS and TRC reports.
//  Non-blocking — each runs in its own 6-minute window.
// ============================================================

function installPeaksMenuTrigger() { installDailyTrigger_Sheet3(); }

function installDailyTrigger_Sheet3() {
  var managed = [
    "s3_OMS_trigger","s3_TRC_trigger","buildPeaksMenu",
    "_s3_runNextInQueue","runDailyReports_Sheet3",
    "installDailyTrigger_Sheet3","installPeaksMenuTrigger"
  ];
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (managed.indexOf(t.getHandlerFunction()) > -1)
      try { ScriptApp.deleteTrigger(t); } catch(_) {}
  });
  var ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger("buildPeaksMenu").forSpreadsheet(ss).onOpen().create();
  ["s3_OMS_trigger","s3_TRC_trigger"].forEach(function(fn) {
    ScriptApp.newTrigger(fn).timeBased().atHour(S3_HOUR).everyDays(1)
      .inTimezone("Asia/Kolkata").create();
  });
  console.log("Sheet 3 triggers installed at " + S3_HOUR + ":00 IST.");
  try {
    SpreadsheetApp.getUi().alert(
      "Sheet 3 triggers installed!\n\n" +
      "OMS Peaks and TRC UML reports will run automatically at " + S3_HOUR + ":00 AM IST daily.\n\n" +
      "Ensure Sheet3_authorizeSheet1() has been run to enable cache writing to Sheet 1."
    );
  } catch(_) {}
}