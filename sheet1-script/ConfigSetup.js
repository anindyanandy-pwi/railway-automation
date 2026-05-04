// ============================================================
//  ConfigSetup.gs — Sheet 1 ONLY
//  Run setupAllSystemTabs() ONCE to create all hidden tabs,
//  the config tab, status dashboard, weekly summary, and SOP.
// ============================================================

function setupAllSystemTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  _createConfigTab(ss);
  _createCacheTab(ss);
  _createWeeklyCacheTab(ss);
  _createStatusTab(ss);
  _createWeeklySummaryTab(ss);
  _createSOPTab(ss);
  try { SpreadsheetApp.getUi().alert(
    "Setup complete!\n\n" +
    "All system tabs created.\n\n" +
    "Next steps:\n" +
    "1. Fill in AEN email addresses in the _CONFIG_ tab\n" +
    "2. Fill in PWI names under each AEN in the _CONFIG_ tab\n" +
    "3. Run cleanupAndSetup() to install all triggers\n\n" +
    "See SYSTEM SOP tab for full instructions."
  ); } catch(_) { console.log("Setup complete."); }
}

function _createConfigTab(ss) {
  var sh = ss.getSheetByName(CM_TAB_CONFIG);
  if (sh) { sh.clearContents(); sh.clearFormats(); }
  else sh = ss.insertSheet(CM_TAB_CONFIG);

  var rows = [
    // AEN Master
    ["AEN_MASTER", "(header — do not edit this row)"],
    ["LLH",       "AEN/LLH"],
    ["AEN/LLH",   "AEN/LLH"],
    ["BWN",       "AEN/BWN"],
    ["AEN/BWN",   "AEN/BWN"],
    ["BDC",       "AEN/BDC"],
    ["AEN/BDC",   "AEN/BDC"],
    ["DKAE",      "AEN/DKAE"],
    ["AEN/DKAE",  "AEN/DKAE"],
    ["TR/HWH",    "AEN/TR/HWH"],
    ["AEN/TR/HWH","AEN/TR/HWH"],
    ["BHP",       "AEN/BHP"],
    ["AEN/BHP",   "AEN/BHP"],
    ["RPH",       "AEN/RPH"],
    ["AEN/RPH",   "AEN/RPH"],
    ["NDAE",      "AEN/NDAE"],
    ["AEN/NDAE",  "AEN/NDAE"],
    ["KWAE",      "AEN/KWAE"],
    ["AEN/KWAE",  "AEN/KWAE"],
    ["AZ",        "AEN/AZ"],
    ["AEN/AZ",    "AEN/AZ"],
    [],
    // PWI Master — user fills this in
    // Format: AEN code in col A, then PWI codes in B, C, D...
    // Example: AEN/LLH | PWI/HGY | PWI/TAK | PWI/LLH
    ["PWI_MASTER","(Fill in PWI codes for each AEN — use separate columns per PWI)"],
    ["AEN/LLH",  "","","",""],
    ["AEN/BWN",  "","","",""],
    ["AEN/BDC",  "","","",""],
    ["AEN/DKAE", "","","",""],
    ["AEN/TR/HWH","","","",""],
    ["AEN/BHP",  "","","",""],
    ["AEN/RPH",  "","","",""],
    ["AEN/NDAE", "","","",""],
    ["AEN/KWAE", "","","",""],
    ["AEN/AZ",   "","","",""],
    [],
    // AEN Emails — user fills this in
    // Multiple emails: add in columns B, C, D...
    ["AEN_EMAILS","(Fill in email addresses — add multiple in same row across columns)"],
    ["AEN/TR/HWH",""],
    ["AEN/LLH",  "aen1llh@gmail.com"],
    ["AEN/BDC",  ""],
    ["AEN/DKAE", "aendkaeminicontrol@gmail.com","dkaeaen@gmail.com"],
    ["AEN/BWN",  ""],
    ["AEN/BHP",  "aenbhper@gmail.com"],
    ["AEN/RPH",  "er.aenrph@gmail.com"],
    ["AEN/NDAE", "aenndae@gmail.com","adenndaeminicontrol@gmail.com"],
    ["AEN/KWAE", "aenkatwa.er@gmail.com"],
    ["AEN/AZ",   "aenazimganj@gmail.com","adenazminicontrol@gmail.com"],
    [],
    // SL Config
    ["SL_CONFIG","SL No","Item Name in Daily Report"],
    ["generateBadRoadReport",    13, "Bad Road Surface"],
    ["generateGaplessReport",    29, "Gapless joint of CMS Crossing"],
    ["generateTWSTieBarReport",  35, "TWS Tie Bar"],
    ["generateGJointReport",     37, "Glued Jt Suspension"],
    ["generateSandHumpReport",   41, "Sand Hump"],
    ["generateBufferReport",     42, "Buffer"],
    ["generateROBReport",        43, "ROB Guard Rail"],
    ["generateFOBReport",        43, "FOB Guard Rail"],
    ["generateBridgeReport",     43, "Bridge Guard Rail"],
    ["generateOMSReport",        51, "OMS/TRC UML Peaks"],
    ["generateTRCReport",        51, "OMS/TRC UML Peaks"],
    []
  ];

  sh.getRange(1,1,rows.length, 6).setValues(
    rows.map(function(r){ while(r.length<6) r.push(""); return r; })
  );

  // Styling
  var H = CM_C.HEADER_BG, T = CM_C.TITLE_BG;
  [[1,"AEN_MASTER"],[23,"PWI_MASTER"],[35,"AEN_EMAILS"],[47,"SL_CONFIG"]].forEach(function(p){
    sh.getRange(p[0],1,1,6).setBackground(T).setFontColor(CM_C.TITLE_FG).setFontWeight("bold");
  });

  sh.setColumnWidth(1,200); sh.setColumnWidth(2,220); sh.setColumnWidth(3,180);
  sh.hideSheet();
  console.log("_CONFIG_ tab created.");
}

function _createCacheTab(ss) {
  var sh = ss.getSheetByName(CM_TAB_CACHE);
  if (sh) return; // preserve existing cache
  sh = ss.insertSheet(CM_TAB_CACHE);
  sh.appendRow(["Script Function","Date","SL No","Summary JSON","Consecutive Days","Timestamp"]);
  sh.getRange(1,1,1,6).setBackground(CM_C.HEADER_BG).setFontColor(CM_C.HEADER_FG).setFontWeight("bold");
  sh.setColumnWidth(4,400);
  sh.hideSheet();
  console.log("_EXCEPTION_CACHE_ tab created.");
}

function _createWeeklyCacheTab(ss) {
  var sh = ss.getSheetByName(CM_TAB_WEEKLY_CACHE);
  if (sh) return;
  sh = ss.insertSheet(CM_TAB_WEEKLY_CACHE);
  sh.appendRow(["Script Function","MON","TUE","WED","THU","FRI","SAT","Week Start"]);
  sh.getRange(1,1,1,8).setBackground(CM_C.HEADER_BG).setFontColor(CM_C.HEADER_FG).setFontWeight("bold");
  sh.setColumnWidth(2,300); sh.setColumnWidth(3,300);
  sh.hideSheet();
  console.log("_WEEKLY_CACHE_ tab created.");
}

function _createStatusTab(ss) {
  var sh = ss.getSheetByName(CM_TAB_STATUS);
  if (sh) { sh.clearContents(); sh.clearFormats(); }
  else sh = ss.insertSheet(CM_TAB_STATUS);

  sh.appendRow(["SYSTEM STATUS — HOWRAH DIVISION EXCEPTION REPORT SYSTEM"]);
  sh.getRange(1,1,1,4).merge()
    .setBackground(CM_C.TITLE_BG).setFontColor(CM_C.TITLE_FG).setFontWeight("bold")
    .setHorizontalAlignment("center");
  sh.appendRow([" "," "," "," "]);
  sh.appendRow(["Script / Report","Last Run","Exceptions Found","Status"]);
  sh.getRange(3,1,1,4).setBackground(CM_C.HEADER_BG).setFontColor(CM_C.HEADER_FG).setFontWeight("bold");

  var scripts = [
    ["Bad Road Surface","—","—"],["TWS Tie Bar","—","—"],
    ["Gapless joint of CMS Crossing","—","—"],["Glued Jt Suspension","—","—"],
    ["Sand Hump","—","—"],["Buffer","—","—"],
    ["ROB Guard Rail","—","—"],["FOB Guard Rail","—","—"],
    ["Bridge Guard Rail","—","—"],["OMS/TRC UML Peaks","—","—"]
  ];
  scripts.forEach(function(s){
    sh.appendRow([s[0],s[1],s[2],"✗ NEVER"]);
  });
  var lr = sh.getLastRow();
  sh.getRange(4,1,lr-3,3).setBackground(CM_C.DATA_BG).setFontColor(CM_C.DATA_FG);
  sh.getRange(4,4,lr-3,1).setBackground(CM_C.ST_EXP_BG).setFontColor(CM_C.ST_EXP_FG).setFontWeight("bold");
  sh.setColumnWidth(1,260); sh.setColumnWidth(2,180); sh.setColumnWidth(3,160); sh.setColumnWidth(4,180);
  console.log("SYSTEM STATUS tab created.");
}

function _createWeeklySummaryTab(ss) {
  var sh = ss.getSheetByName(CM_TAB_WEEKLY_SUM);
  if (sh) return; // preserve existing data
  sh = ss.insertSheet(CM_TAB_WEEKLY_SUM);
  sh.appendRow(["WEEKLY EXCEPTION SUMMARY — HOWRAH DIVISION"]);
  sh.getRange(1,1,1,8).merge()
    .setBackground(CM_C.TITLE_BG).setFontColor(CM_C.TITLE_FG).setFontWeight("bold")
    .setHorizontalAlignment("center");
  sh.appendRow(["Weekly summaries are prepended above each week. Generated every Friday at 9:00 AM IST."]);
  sh.setColumnWidth(1,280); for (var c=2;c<=8;c++) sh.setColumnWidth(c,80);
  console.log("WEEKLY SUMMARY tab created.");
}

function _createSOPTab(ss) {
  var sh = ss.getSheetByName(CM_TAB_SOP);
  if (sh) { sh.clearContents(); sh.clearFormats(); }
  else sh = ss.insertSheet(CM_TAB_SOP);

  var sop = [
    ["SYSTEM SOP — HOWRAH DIVISION EXCEPTION REPORT SYSTEM","","","",""],
    ["Last Updated: "+cm_today()+" | Version: 1.0","","","",""],
    ["","","","",""],
    ["SECTION 1: INITIAL SETUP","","","",""],
    ["Step","Action","Where","Notes",""],
    ["1","Run setupAllSystemTabs()","Sheet 1 → Apps Script → ConfigSetup.gs","One time only",""],
    ["2","Fill AEN emails in _CONFIG_ tab","Sheet 1 → _CONFIG_ tab → AEN_EMAILS section","Add multiple emails per row",""],
    ["3","Fill PWI names in _CONFIG_ tab","Sheet 1 → _CONFIG_ tab → PWI_MASTER section","For future use",""],
    ["4","Run cleanupAndSetup()","Sheet 1 → Apps Script → CleanupSetup.gs","One time only",""],
    ["5","Run Sheet2_authorizeSheet1()","Sheet 2 → Apps Script → Sheet2_Auth.gs","One time only",""],
    ["6","Run Sheet3_authorizeSheet1()","Sheet 3 → Apps Script → Sheet3_Auth.gs","One time only",""],
    ["","","","",""],
    ["SECTION 2: DAILY OPERATION","","","",""],
    ["Time","Action","Result","Notes",""],
    ["9:00 AM","Automatic triggers run","All reports generated, PDFs saved, emails sent","No action needed",""],
    ["Anytime","Exception Report menu → Generate & Email All Reports","Runs all reports for current sheet","Background, non-blocking",""],
    ["Anytime","Exception Report menu → Generate Combined Daily Report","Populates DAILY EXCEPTION REPORT tab","Two-click process",""],
    ["","","","",""],
    ["SECTION 3: ADDING A NEW EXCEPTION REPORT SCRIPT","","","",""],
    ["Step","Action","File","Notes",""],
    ["1","Create new .gs file in the appropriate sheet's Apps Script","New file","Name it clearly",""],
    ["2","Add cm_writeSummaryToCache() call at end of your exception scan","Your new script","See existing scripts for examples",""],
    ["3","Add cm_updateStatusDashboard() call","Your new script","Keeps status tab updated",""],
    ["4","Add SL number row in _CONFIG_ tab → SL_CONFIG section","_CONFIG_ tab","Format: functionName | SL# | Item Name",""],
    ["5","Add menu item in addExceptionReportMenu() in bad_road_surface_apps_script.gs","bad_road_surface_apps_script.gs","Add .addItem() line",""],
    ["6","Add new function to S1_CHAIN_FNS in Sheet1_DailyScheduler.gs","Sheet1_DailyScheduler.gs","Enables daily trigger",""],
    ["","","","",""],
    ["SECTION 4: UPDATING SL NUMBERS","","","",""],
    ["If the SL number of an item changes in the daily report:","","","",""],
    ["1","Open _CONFIG_ tab in Sheet 1","","",""],
    ["2","Find the row in SL_CONFIG section with the script name","","",""],
    ["3","Update the SL number in column B","","",""],
    ["4","No code changes required","","",""],
    ["","","","",""],
    ["SECTION 5: CHANGING SCHEDULE TIME","","","",""],
    ["Option A (no code): Apps Script → clock icon → edit trigger → change hour","","","",""],
    ["Option B (via code): Change S1_HOUR in Sheet1_DailyScheduler.gs → re-run installDailyTrigger_Sheet1()","","","",""],
    ["","","","",""],
    ["SECTION 6: ADDING/REMOVING EMAIL RECIPIENTS","","","",""],
    ["Open Sheet1_DailyScheduler.gs → edit the S1_RECIPIENTS array","","","",""],
    ["For AEN-specific emails: update _CONFIG_ tab → AEN_EMAILS section","","","",""],
    ["","","","",""],
    ["SECTION 7: TROUBLESHOOTING","","","",""],
    ["Problem","Likely Cause","Solution","",""],
    ["Reports not running at 9 AM","Trigger not installed","Run installDailyTrigger_Sheet1() / Sheet2 / Sheet3","",""],
    ["Email quota exceeded","Too many test runs","Wait for midnight quota reset. Use test mode with 1 recipient","",""],
    ["Cache stale warning","Reports not run today","Run individual reports or Generate & Email All","",""],
    ["Section not found in PDF","Script ran but exception report tab empty","Run the individual report first","",""],
    ["Combined report timed out","Too many reports running in series","Use background chain — click once and wait for email","",""],
    ["_CONFIG_ tab missing","Setup not run","Run setupAllSystemTabs()","",""]
  ];

  sop.forEach(function(row,i){
    while(row.length<5) row.push("");
    sh.getRange(i+1,1,1,5).setValues([row]);
  });

  // Section headers styling
  [1,2,4,13,21,31,37,43,51].forEach(function(r){
    try { sh.getRange(r,1,1,5).setBackground(CM_C.TITLE_BG).setFontWeight("bold"); } catch(_) {}
  });
  [5,14,22,32,38].forEach(function(r){
    try { sh.getRange(r,1,1,5).setBackground(CM_C.HEADER_BG).setFontWeight("bold"); } catch(_) {}
  });

  sh.setColumnWidth(1,350); sh.setColumnWidth(2,280); sh.setColumnWidth(3,280);
  console.log("SYSTEM SOP tab created.");
}