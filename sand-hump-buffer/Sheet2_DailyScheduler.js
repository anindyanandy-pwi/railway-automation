// ============================================================
//  Sheet2_DailyScheduler.gs — Sheet 2
//  Individual daily triggers for each Sheet 2 report.
//  Non-blocking — each runs in its own 6-minute window.
// ============================================================

// Individual runners called by daily time triggers
function s2_SandHump_trigger()  { s2_SandHump();  }
function s2_Buffer_trigger()    { s2_Buffer();     }
function s2_ROB_trigger()       { s2_ROB();        }
function s2_FOB_trigger()       { s2_FOB();        }
function s2_Bridge_trigger()    { s2_Bridge();     }

function installDailyTrigger_Sheet2() {
  // Remove all existing managed triggers
  var managed = [
    "s2_SandHump_trigger","s2_Buffer_trigger","s2_ROB_trigger",
    "s2_FOB_trigger","s2_Bridge_trigger","buildExceptionMenu",
    "_s2_runNextInQueue","runDailyReports_Sheet2","installDailyTrigger_Sheet2"
  ];
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (managed.indexOf(t.getHandlerFunction()) > -1)
      try { ScriptApp.deleteTrigger(t); } catch(_) {}
  });

  var ss = SpreadsheetApp.getActive();

  // Menu trigger
  ScriptApp.newTrigger("buildExceptionMenu")
    .forSpreadsheet(ss).onOpen().create();

  // Individual report triggers at S2_HOUR (9 AM) — each gets own window
  var fns = ["s2_SandHump_trigger","s2_Buffer_trigger",
             "s2_ROB_trigger","s2_FOB_trigger","s2_Bridge_trigger"];
  fns.forEach(function(fn) {
    ScriptApp.newTrigger(fn)
      .timeBased().atHour(S2_HOUR).everyDays(1)
      .inTimezone("Asia/Kolkata").create();
  });

  console.log("Sheet 2 triggers installed at " + S2_HOUR + ":00 IST.");
  try {
    SpreadsheetApp.getUi().alert(
      "Sheet 2 triggers installed!\n\n" +
      "Reports will run automatically at " + S2_HOUR + ":00 AM IST daily.\n\n" +
      "Each report runs independently — no timeout risk.\n\n" +
      "Ensure Sheet2_authorizeSheet1() has been run to enable cache writing."
    );
  } catch(_) {}
}