// ============================================================
//  Sheet3_Auth.gs — Sheet 3 ONLY
//  Run Sheet3_authorizeSheet1() ONCE to authorize access to
//  Sheet 1 for cache and status dashboard writing.
// ============================================================

var S3_SHEET1_ID = "1-EWkp8nq5aL_BKs3MQymexS-1gjrrXtPbLu1ZCZ1_FI";

function Sheet3_authorizeSheet1() {
  try {
    var ss1 = SpreadsheetApp.openById(S3_SHEET1_ID);
    console.log("Sheet 1 access confirmed: " + ss1.getName());
    try { SpreadsheetApp.getUi().alert(
      "Authorization successful!\n\nSheet 3 can now read/write to Sheet 1.\n\nSheet 1 name: " + ss1.getName()
    ); } catch(_) {}
  } catch(e) {
    console.log("Authorization failed: " + e.message);
    try { SpreadsheetApp.getUi().alert("Authorization failed: " + e.message); } catch(_) {}
  }
}

// ── Cross-sheet cache writer (Sheet 3 version) ────────────────────────────────
function s3_writeSummaryToSheet1Cache(scriptFn, summaryObj) {
  try {
    var ss1      = SpreadsheetApp.openById(S3_SHEET1_ID);
    var cacheTab = ss1.getSheetByName("_EXCEPTION_CACHE_");
    if (!cacheTab) {
      console.log("_EXCEPTION_CACHE_ not found in Sheet 1. Run setupAllSystemTabs() in Sheet 1.");
      return;
    }
    var cfgTab = ss1.getSheetByName("_CONFIG_");
    var slNo   = "";
    if (cfgTab) {
      var cfgVals = cfgTab.getRange(1,1,cfgTab.getLastRow(),3).getValues();
      var inSL = false;
      for (var i = 0; i < cfgVals.length; i++) {
        if (cfgVals[i][0] === "SL_CONFIG") { inSL = true; continue; }
        if (inSL && (!cfgVals[i][0] || String(cfgVals[i][0]).indexOf("_") === 0)) break;
        if (inSL && cfgVals[i][0] === scriptFn) { slNo = cfgVals[i][1]; break; }
      }
    }
    var dateStr = s3_fmtDate(new Date());
    var jsonStr = JSON.stringify(summaryObj);
    var consec  = s3_calcConsecutive(cacheTab, scriptFn, summaryObj.totalExceptions);
    var ts      = Utilities.formatDate(new Date(),"Asia/Kolkata","dd.MM.yyyy HH:mm");
    var lr = Math.max(cacheTab.getLastRow(),1), found = -1;
    if (lr > 1) {
      var cv = cacheTab.getRange(2,1,lr-1,2).getValues();
      for (var j = 0; j < cv.length; j++) {
        if (cv[j][0] === scriptFn && cv[j][1] === dateStr) { found = j+2; break; }
      }
    }
    if (found > 0) cacheTab.getRange(found,1,1,6).setValues([[scriptFn,dateStr,slNo,jsonStr,consec,ts]]);
    else cacheTab.appendRow([scriptFn,dateStr,slNo,jsonStr,consec,ts]);

    s3_updateSheet1Status(ss1, scriptFn, summaryObj.totalExceptions, ts);
    s3_writeSheet1WeeklyCache(ss1, scriptFn, summaryObj);
  } catch(e) { console.log("s3_writeSummaryToSheet1Cache error: " + e.message); }
}

function s3_calcConsecutive(cacheTab, scriptFn, todayEx) {
  if (todayEx === 0) return 0;
  try {
    var lr = cacheTab.getLastRow(); if (lr < 2) return 1;
    var yesterday = s3_fmtDate(new Date(new Date().getTime()-86400000));
    var vals = cacheTab.getRange(2,1,lr-1,5).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (vals[i][0] === scriptFn && vals[i][1] === yesterday) return (parseInt(vals[i][4])||0)+1;
    }
    return 1;
  } catch(e) { return 1; }
}

function s3_updateSheet1Status(ss1, scriptFn, exceptionsCount, ts) {
  try {
    var sh = ss1.getSheetByName("SYSTEM STATUS"); if (!sh) return;
    var cfgTab = ss1.getSheetByName("_CONFIG_"), name = scriptFn;
    if (cfgTab) {
      var vals = cfgTab.getRange(1,1,cfgTab.getLastRow(),3).getValues();
      var inSL = false;
      for (var i = 0; i < vals.length; i++) {
        if (vals[i][0] === "SL_CONFIG") { inSL = true; continue; }
        if (inSL && vals[i][0] === scriptFn) { name = String(vals[i][2]).trim(); break; }
      }
    }
    var lr = sh.getLastRow(), found = -1;
    // Data rows start at row 4
    if (lr >= 4) {
      var svals = sh.getRange(4,1,lr-3,1).getValues();
      for (var j = 0; j < svals.length; j++) if (svals[j][0] === name) { found = j+4; break; }
    }
    var row = [name, ts, exceptionsCount, "✓ Fresh"];
    if (found > 0) sh.getRange(found,1,1,4).setValues([row]);
    else { sh.appendRow(row); found = sh.getLastRow(); }
    sh.getRange(found,4).setBackground("#DCFCE7").setFontColor("#000000").setFontWeight("bold");
  } catch(e) { console.log("s3_updateSheet1Status error: " + e.message); }
}

function s3_writeSheet1WeeklyCache(ss1, scriptFn, summaryObj) {
  try {
    var sh = ss1.getSheetByName("_WEEKLY_CACHE_"); if (!sh) return;
    var d = new Date(), day = ["SUN","MON","TUE","WED","THU","FRI","SAT"][d.getDay()];
    var dayColMap = {MON:1,TUE:2,WED:3,THU:4,FRI:5}; var dayCol = dayColMap[day]; if (!dayCol) return;
    var weekStart = s3_getWeekStart(), jsonStr = JSON.stringify(summaryObj);
    var lr = Math.max(sh.getLastRow(),1), found = -1;
    if (lr > 1) {
      var vals = sh.getRange(2,1,lr-1,8).getValues();
      for (var i = 0; i < vals.length; i++) {
        if (vals[i][0] === scriptFn && vals[i][7] === weekStart) { found = i+2; break; }
      }
    }
    if (found > 0) sh.getRange(found, dayCol+1).setValue(jsonStr);
    else { var row = new Array(8).fill(""); row[0]=scriptFn; row[dayCol]=jsonStr; row[7]=weekStart; sh.appendRow(row); }
  } catch(e) { console.log("s3_writeSheet1WeeklyCache error: " + e.message); }
}

function s3_fmtDate(d) {
  return ('0'+d.getDate()).slice(-2)+'.'+('0'+(d.getMonth()+1)).slice(-2)+'.'+d.getFullYear();
}
function s3_getWeekStart() {
  var d=new Date(), day=d.getDay(), diff=(day===0)?-6:1-day;
  return s3_fmtDate(new Date(d.getFullYear(),d.getMonth(),d.getDate()+diff));
}

// ── AEN normalization (Sheet 3) ───────────────────────────────────────────────
var S3_AEN_MAP = (function() {
  var codes=["TR/HWH","LLH","BDC","DKAE","BWN","BHP","RPH","NDAE","KWAE","AZ"];
  var m={};
  codes.forEach(function(c){
    m[c.toLowerCase()]="AEN/"+c;
    m[("aen/"+c).toLowerCase()]="AEN/"+c;
  });
  return m;
}());

function s3_normalizeAEN(raw) {
  if (!raw) return null;
  var s = String(raw).trim().replace(/\s*\([^)]*\)/g, '').trim();
  var lower = s.toLowerCase();
  if (S3_AEN_MAP[lower]) return S3_AEN_MAP[lower];
  var slashed = lower.replace(/\s+/g, '/');
  if (S3_AEN_MAP[slashed]) return S3_AEN_MAP[slashed];
  if (S3_AEN_MAP['aen/' + slashed]) return S3_AEN_MAP['aen/' + slashed];
  return null;
}
function s3_extractAENFromLabel(label) {
  var s = String(label||"");
  var m = s.match(/\(AEN:\s*([^)]+)\)/i);
  if (m) return s3_normalizeAEN(m[1].trim());
  m = s.match(/(?:^|\|\s*)AEN:\s*([^\|\(]+)/i);
  if (m) return s3_normalizeAEN(m[1].trim());
  return null;
}
function s3_buildSummary(exTypesObj) {
  var result={date:s3_fmtDate(new Date()),exceptions:{},totalExceptions:0};
  for (var exType in exTypesObj) {
    var items=exTypesObj[exType]||[];
    for (var i=0;i<items.length;i++) {
      var item=items[i];
      var label=typeof item==='string'?item:(item.label||String(item));
      var aen=s3_extractAENFromLabel(label)||"UNASSIGNED";
      if (!result.exceptions[aen]) result.exceptions[aen]={};
      result.exceptions[aen][exType]=(result.exceptions[aen][exType]||0)+1;
      result.totalExceptions++;
    }
  }
  return result;
}

// ── Color constants (Sheet 3) ─────────────────────────────────────────────────
var S3_C = {
  TITLE_BG:"#BFDBFE",  TITLE_FG:"#000000",
  HEADER_BG:"#DBEAFE", HEADER_FG:"#000000",
  OK_BG:"#DCFCE7",     OK_FG:"#000000",
  EXCEPT_BG:"#FEE2E2", EXCEPT_FG:"#000000",
  DATA_BG:"#FFFFFF",   DATA_FG:"#1F2937",
  WHITE:"#FFFFFF",     BLACK:"#000000"
};

// ── Menu shortcut to Sheet 1 ──────────────────────────────────────────────────
function openSheet1DailyReport_S3() {
  var url = "https://docs.google.com/spreadsheets/d/"+S3_SHEET1_ID+"/edit#gid=855312642";
  try {
    var html = HtmlService.createHtmlOutput('<script>window.open("'+url+'","_blank");google.script.host.close();</script>');
    SpreadsheetApp.getUi().showModalDialog(html, "Opening Sheet 1 Daily Report...");
  } catch(_) { console.log("Open Sheet 1: "+url); }
}