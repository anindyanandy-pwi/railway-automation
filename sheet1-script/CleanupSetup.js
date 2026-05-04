// ============================================================
//  CleanupSetup.gs — Sheet 1 ONLY
//  Run cleanupAndSetup() ONCE after setupAllSystemTabs().
//  Removes old triggers, installs fresh triggers for all reports.
// ============================================================

// All trigger handler functions managed by this system (Sheet 1)
var S1_MANAGED_TRIGGERS = [
  "addExceptionReportMenu","addBadRoadMenu","addExceptionReportMenu",
  "buildExceptionMenu","buildPeaksMenu",
  "runDailyReports_Sheet1","runDailyReports_Sheet2","runDailyReports_Sheet3",
  "installDailyTrigger_Sheet1","installBadRoadMenuTrigger","installPeaksMenuTrigger",
  "s1_BadRoadSurface","s1_TWSTieBar","s1_GaplessCMS","s1_GJoint",
  "_s1_runNextInQueue","_s2_runNextInQueue","_s3_runNextInQueue",
  "addCombinedReportMenu","cm_refreshAllStatuses","generateWeeklySummary",
  "onEditDailyReport","_dr_continueChain","_dr_sendAENEmails"
];

// ── Master setup function ─────────────────────────────────────────────────────
function cleanupAndSetup() {
  console.log("Starting cleanup and setup...");
  _deleteOldTriggers();
  _installAllSheet1Triggers();
  console.log("Cleanup and setup complete.");
  try {
    SpreadsheetApp.getUi().alert(
      "Setup complete!\n\n" +
      "All old triggers removed.\n" +
      "New triggers installed:\n" +
      "  • Menu trigger (on open)\n" +
      "  • Individual report triggers (9 AM daily)\n" +
      "  • Status refresh trigger (every 6 hours)\n" +
      "  • Weekly summary trigger (Friday 9 AM)\n" +
      "  • Column H timestamp trigger (on edit)\n\n" +
      "Sender email: " + Session.getActiveUser().getEmail() + "\n\n" +
      "Check SYSTEM STATUS tab to verify all is working."
    );
  } catch(_) { console.log("Setup complete. Check execution log."); }
}

// ── Delete all old/managed triggers ──────────────────────────────────────────
function _deleteOldTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var deleted  = 0;
  triggers.forEach(function(t) {
    var fn = t.getHandlerFunction();
    // Only delete triggers owned by current user and managed by this system
    if (S1_MANAGED_TRIGGERS.indexOf(fn) > -1) {
      try { ScriptApp.deleteTrigger(t); deleted++; } catch(_) {}
    }
  });
  console.log("Deleted " + deleted + " old triggers.");
}

// ── Install all Sheet 1 triggers ──────────────────────────────────────────────
function _installAllSheet1Triggers() {
  var ss = SpreadsheetApp.getActive();

  // 1. Menu trigger (on open)
  ScriptApp.newTrigger("addCombinedReportMenu")
    .forSpreadsheet(ss).onOpen().create();

  // 2. Individual report triggers — each at 9 AM, staggered
  // All scheduled at same hour; Apps Script fires them independently
  // so each gets its own 6-minute window
  var reportFns = ["s1_BadRoadSurface","s1_TWSTieBar","s1_GaplessCMS","s1_GJoint"];
  reportFns.forEach(function(fn) {
    ScriptApp.newTrigger(fn)
      .timeBased().atHour(S1_HOUR).everyDays(1).inTimezone("Asia/Kolkata").create();
  });

  // 3. Status refresh every 6 hours
  ScriptApp.newTrigger("cm_refreshAllStatuses")
    .timeBased().everyHours(6).create();

  // 4. Weekly summary — Friday at 9 AM
  ScriptApp.newTrigger("generateWeeklySummary")
    .timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(9).inTimezone("Asia/Kolkata").create();

  // 5. Column H timestamp trigger (on edit)
  ScriptApp.newTrigger("onEditDailyReport")
    .forSpreadsheet(ss).onEdit().create();

  console.log("Installed " + (reportFns.length + 4) + " triggers.");
}

// ── Re-run if schedule time changes ──────────────────────────────────────────
// Call this after changing S1_HOUR in Sheet1_DailyScheduler.gs
function reinstallReportTriggers() {
  // Remove only time-based report triggers
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    var reportFns = ["s1_BadRoadSurface","s1_TWSTieBar","s1_GaplessCMS","s1_GJoint"];
    if (reportFns.indexOf(fn) > -1) {
      try { ScriptApp.deleteTrigger(t); } catch(_) {}
    }
  });
  var ss = SpreadsheetApp.getActive();
  ["s1_BadRoadSurface","s1_TWSTieBar","s1_GaplessCMS","s1_GJoint"].forEach(function(fn) {
    ScriptApp.newTrigger(fn)
      .timeBased().atHour(S1_HOUR).everyDays(1).inTimezone("Asia/Kolkata").create();
  });
  console.log("Report triggers reinstalled at hour " + S1_HOUR);
  try { SpreadsheetApp.getUi().alert("Report triggers reinstalled at "+S1_HOUR+":00 IST."); }
  catch(_) {}
}

// ── Column H onEdit trigger ───────────────────────────────────────────────────
// Watches column G of DAILY EXCEPTION REPORT tab
// When someone types in G, auto-fills timestamp in H
function onEditDailyReport(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (sheet.getName() !== CM_TAB_DAILY_RPT) return;
    var col = e.range.getColumn();
    if (col !== 7) return; // Only watch column G
    var row = e.range.getRow();
    var cellA = sheet.getRange(row, 1).getValue();
    if (!cellA) return; // skip empty rows

    // Identify editor by email → AEN name
    var editorEmail = "";
    var aenName     = "";
    try { editorEmail = Session.getActiveUser().getEmail().toLowerCase().trim(); } catch(_) {}

    if (editorEmail) {
      // Look up email in _CONFIG_ AEN_EMAILS section
      try {
        var ss  = SpreadsheetApp.getActiveSpreadsheet();
        var cfg = ss.getSheetByName(CM_TAB_CONFIG);
        if (cfg) {
          var vals = cfg.getRange(1,1,cfg.getLastRow(),6).getValues();
          var inSection = false;
          outer:
          for (var i = 0; i < vals.length; i++) {
            if (vals[i][0] === "AEN_EMAILS") { inSection = true; continue; }
            if (inSection && (!vals[i][0] || String(vals[i][0]).indexOf("_") === 0)) break;
            if (inSection && vals[i][0]) {
              for (var c = 1; c < 6; c++) {
                if (String(vals[i][c]).toLowerCase().trim() === editorEmail) {
                  aenName = String(vals[i][0]).trim();
                  break outer;
                }
              }
            }
          }
        }
      } catch(_) {}
    }

    var ts  = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd.MM.yyyy HH:mm");
    var ackText = aenName
      ? aenName + " acknowledged: " + ts
      : "Acknowledged: " + ts;

    sheet.getRange(row, 8).setValue(ackText)
         .setBackground(CM_C.OK_BG).setFontColor(CM_C.OK_FG).setFontSize(9);
  } catch(err) {
    console.log("onEditDailyReport error: " + err.message);
  }
}

// ── Menu ──────────────────────────────────────────────────────────────────────
// This is the SINGLE onOpen menu for Sheet 1
// ALL menu items for ALL scripts go here
function addCombinedReportMenu() {
  SpreadsheetApp.getUi()
    .createMenu("Exception Report")
    .addItem("Bad Road Surface Report",          "generateBadRoadReport")
    .addItem("TWS Tie Bar Report",               "generateTWSTieBarReport")
    .addItem("Gapless Joint CMS Report",         "generateGaplessReport")
    .addItem("Physically Damaged G/Joint",       "generateGJointReport")
    .addSeparator()
    .addItem("Generate & Email All Reports",      "manualSendAll_Sheet1")
    .addItem("Generate Combined Daily Report",    "generateCombinedDailyReport")
    .addSeparator()
    .addItem("Refresh Status Dashboard",          "cm_refreshAllStatuses")
    .addItem("Open Daily Report Sheet",           "openDailyReportSheet")
    .addItem("Fetch TMS Compliance Data", "pwiStartTMSFetch")
    .addItem("Generate PWI Compliance Exceptions", "generateCompliancePWIReport")
    .addItem("Import TMS Files from Drive", "pwi_processUploadedFiles")
    .addToUi();
}

function openDailyReportSheet() {
  var url = "https://docs.google.com/spreadsheets/d/"+CM_SHEET1_ID+"/edit#gid=855312642";
  try {
    var html = HtmlService.createHtmlOutput('<script>window.open("'+url+'","_blank");google.script.host.close();</script>');
    SpreadsheetApp.getUi().showModalDialog(html, "Opening...");
  } catch(_) {}
}