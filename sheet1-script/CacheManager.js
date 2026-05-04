// ============================================================
//  CacheManager.gs — Sheet 1 ONLY
//  Foundation file. Contains color constants, AEN normalization,
//  cache read/write, status dashboard, and weekly cache helpers.
//  All Sheet 1 scripts depend on this file.
// ============================================================

// ── Core IDs ─────────────────────────────────────────────────────────────────
var CM_SHEET1_ID    = "1-EWkp8nq5aL_BKs3MQymexS-1gjrrXtPbLu1ZCZ1_FI";
var CM_FOLDER_ID    = "1InCD6mlBHmvuvtgp1ycTY-bNkZxwqbld";
var CM_PARENT_EMAIL = "enggplanningcellhwh@gmail.com";

// ── Tab names ─────────────────────────────────────────────────────────────────
var CM_TAB_CACHE        = "_EXCEPTION_CACHE_";
var CM_TAB_WEEKLY_CACHE = "_WEEKLY_CACHE_";
var CM_TAB_CONFIG       = "_CONFIG_";
var CM_TAB_STATUS       = "SYSTEM STATUS";
var CM_TAB_WEEKLY_SUM   = "WEEKLY SUMMARY";
var CM_TAB_DAILY_RPT    = "DAILY EXCEPTION REPORT";
var CM_TAB_SOP          = "SYSTEM SOP";
var CM_TAB_EXCEPTION    = "Exception Report";

// ── Color scheme ─────────────────────────────────────────────────────────────
var CM_C = {
  TITLE_BG    : "#BFDBFE",  TITLE_FG    : "#000000",
  HEADER_BG   : "#DBEAFE",  HEADER_FG   : "#000000",
  OK_BG       : "#DCFCE7",  OK_FG       : "#000000",
  EXCEPT_BG   : "#FEE2E2",  EXCEPT_FG   : "#000000",
  DATA_BG     : "#FFFFFF",  DATA_FG     : "#1F2937",
  ST_FRESH_BG : "#DCFCE7",  ST_FRESH_FG : "#000000",
  ST_YEST_BG  : "#FEF9C3",  ST_YEST_FG  : "#000000",
  ST_STALE_BG : "#FEE2E2",  ST_STALE_FG : "#000000",
  ST_EXP_BG   : "#991B1B",  ST_EXP_FG   : "#FFFFFF",
  OVERWRITE   : "#FEFCE8",
  WHITE       : "#FFFFFF",  BLACK       : "#000000"
};

// ── AEN canonical map ─────────────────────────────────────────────────────────
var CM_AEN_MAP = (function() {
  var codes = ["TR/HWH","LLH","BDC","DKAE","BWN","BHP","RPH","NDAE","KWAE","AZ"];
  var m = {};
  codes.forEach(function(c) {
    m[c.toLowerCase()]          = "AEN/" + c;
    m[("aen/"+c).toLowerCase()] = "AEN/" + c;
    m[("AEN/"+c).toLowerCase()] = "AEN/" + c;
  });
  return m;
}());

// ── Date helpers ──────────────────────────────────────────────────────────────
function cm_fmtDate(d) {
  if (!d) return "";
  return ('0'+d.getDate()).slice(-2)+'.'+('0'+(d.getMonth()+1)).slice(-2)+'.'+d.getFullYear();
}
function cm_today() { return cm_fmtDate(new Date()); }
function cm_parseDate(s) {
  if (!s) return null;
  if (s instanceof Date && !isNaN(s.getTime())) return s;
  var str = String(s).split(' ')[0].trim();
  var p = str.indexOf('.') > -1 ? str.split('.') : str.split('/');
  if (p.length !== 3) return null;
  var d=parseInt(p[0],10), mo=parseInt(p[1],10)-1, y=parseInt(p[2],10);
  if (y < 100) y += 2000;
  var dt = new Date(y,mo,d);
  return isNaN(dt.getTime()) ? null : dt;
}
function cm_daysBetween(d1, d2) { return Math.floor((d2-d1)/86400000); }
function cm_getWeekStart() {
  var d=new Date(), day=d.getDay(), diff=(day===0)?-6:1-day;
  return cm_fmtDate(new Date(d.getFullYear(),d.getMonth(),d.getDate()+diff));
}

// ── AEN normalization ─────────────────────────────────────────────────────────
function cm_normalizeAEN(raw) {
  if (!raw) return null;
  var s = String(raw).trim();

  // Step 1: Strip anything in parentheses — handles "AEN/TR/HWH (AEN/1/HWH)" → "AEN/TR/HWH"
  s = s.replace(/\s*\([^)]*\)/g, '').trim();

  // Step 2: Lowercase for lookup
  var lower = s.toLowerCase();

  // Step 3: Direct map lookup
  if (CM_AEN_MAP[lower]) return CM_AEN_MAP[lower];

  // Step 4: Replace spaces with / and try again — handles "TR HWH" → "TR/HWH"
  var slashed = lower.replace(/\s+/g, '/');
  if (CM_AEN_MAP[slashed]) return CM_AEN_MAP[slashed];
  if (CM_AEN_MAP['aen/' + slashed]) return CM_AEN_MAP['aen/' + slashed];

  // Step 5: Fallback — read config tab for any extra variants user added
  try {
    var cfg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CM_TAB_CONFIG);
    if (cfg) {
      var vals = cfg.getRange(1,1,cfg.getLastRow(),2).getValues();
      var inSection = false;
      for (var i=0; i<vals.length; i++) {
        if (vals[i][0]==="AEN_MASTER") { inSection=true; continue; }
        if (inSection && !vals[i][0]) continue;
        if (inSection && String(vals[i][0]).indexOf("_")===0) break;
        var v = String(vals[i][0]).trim().toLowerCase();
        if (v === lower || v === slashed || v.replace(/\s+/g,'/') === slashed)
          return String(vals[i][1]).trim();
      }
    }
  } catch(_) {}
  return null;
}

function cm_extractAENFromLabel(label) {
  var s = String(label||"");
  // Format 1: "(AEN: XXX)" — used by Bad Road Surface, TWS Tie Bar
  var m = s.match(/\(AEN:\s*([^)]+)\)/i);
  if (m) return cm_normalizeAEN(m[1].trim());
  // Format 2: "AEN: XXX" plain — used by Gapless, GJoint, Sand Hump, Buffer, Guard Rails, OMS, TRC
  // Matches "AEN: XXX" at start of label or immediately after a pipe separator
  m = s.match(/(?:^|\|\s*)AEN:\s*([^\|\(]+)/i);
  if (m) return cm_normalizeAEN(m[1].trim());
  return null;
}

// ── Exception summary builder ─────────────────────────────────────────────────
// exTypesObj: { "Exception type": [labels or {label:...}] }
// Returns canonical summary object for cache storage
function cm_buildSummary(exTypesObj) {
  var result = { date:cm_today(), exceptions:{}, totalExceptions:0 };
  for (var exType in exTypesObj) {
    var items = exTypesObj[exType] || [];
    for (var i=0; i<items.length; i++) {
      var item  = items[i];
      var label = typeof item==='string' ? item : (item.label||String(item));
      var aen   = cm_extractAENFromLabel(label) || "UNASSIGNED";
      if (!result.exceptions[aen]) result.exceptions[aen] = {};
      result.exceptions[aen][exType] = (result.exceptions[aen][exType]||0)+1;
      result.totalExceptions++;
    }
  }
  return result;
}

// ── Cache tab ─────────────────────────────────────────────────────────────────
function cm_getCacheSheet() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName(CM_TAB_CACHE);
  if (!sh) {
    sh=ss.insertSheet(CM_TAB_CACHE);
    sh.appendRow(["Script Function","Date","SL No","Summary JSON","Consecutive Days","Timestamp"]);
    ss.setActiveSheet(ss.getSheetByName(CM_TAB_EXCEPTION)||ss.getSheets()[0]);
    sh.hideSheet();
  }
  return sh;
}

function cm_writeSummaryToCache(scriptFn, summaryObj) {
  try {
    var sh      = cm_getCacheSheet();
    var slInfo  = cm_getSlInfoForScript(scriptFn);
    var slNo    = slInfo ? slInfo.sl : "";
    var dateStr = cm_today();
    var jsonStr = JSON.stringify(summaryObj);
    var consec  = cm_calcConsecutiveDays(scriptFn, summaryObj.totalExceptions);
    var ts      = Utilities.formatDate(new Date(),"Asia/Kolkata","dd.MM.yyyy HH:mm");
    var lr      = sh.getLastRow();
    var found   = -1;
    if (lr > 1) {
      var cv = sh.getRange(2,1,lr-1,2).getValues();
      for (var i=0;i<cv.length;i++) {
        if (cv[i][0]===scriptFn && cv[i][1]===dateStr) { found=i+2; break; }
      }
    }
    if (found>0) {
      sh.getRange(found,1,1,6).setValues([[scriptFn,dateStr,slNo,jsonStr,consec,ts]]);
    } else {
      sh.appendRow([scriptFn,dateStr,slNo,jsonStr,consec,ts]);
    }
    cm_cleanOldEntries(sh);
    // Also write to weekly cache
    cm_writeWeeklyCache(scriptFn, summaryObj);
    // Update status dashboard
    cm_updateStatusDashboard(scriptFn, summaryObj.totalExceptions);
  } catch(e) { console.log("cm_writeSummaryToCache error: "+e.message); }
}

function cm_readCacheForScript(scriptFn, dateStr) {
  try {
    var sh=cm_getCacheSheet(), lr=sh.getLastRow();
    if (lr<2) return null;
    var vals=sh.getRange(2,1,lr-1,6).getValues();
    for (var i=0;i<vals.length;i++) {
      if (vals[i][0]===scriptFn && vals[i][1]===dateStr) {
        return { scriptFn:vals[i][0], date:vals[i][1], slNo:vals[i][2],
                 summary:JSON.parse(vals[i][3]||"{}"),
                 consecutiveDays:vals[i][4], timestamp:vals[i][5] };
      }
    }
    return null;
  } catch(e) { return null; }
}

function cm_readAllCacheForDate(dateStr) {
  try {
    var sh=cm_getCacheSheet(), lr=sh.getLastRow();
    if (lr<2) return {};
    var vals=sh.getRange(2,1,lr-1,6).getValues(), result={};
    for (var i=0;i<vals.length;i++) {
      if (vals[i][1]===dateStr) {
        result[vals[i][0]]={ summary:JSON.parse(vals[i][3]||"{}"),
                              consecutiveDays:vals[i][4], timestamp:vals[i][5] };
      }
    }
    return result;
  } catch(e) { return {}; }
}

function cm_getCacheDateForScript(scriptFn) {
  try {
    var sh=cm_getCacheSheet(), lr=sh.getLastRow();
    if (lr<2) return null;
    var vals=sh.getRange(2,1,lr-1,2).getValues();
    var latest=null;
    for (var i=0;i<vals.length;i++) {
      if (vals[i][0]===scriptFn) {
        var d=cm_parseDate(vals[i][1]);
        if (!latest || (d && d>latest)) latest=d;
      }
    }
    return latest;
  } catch(e) { return null; }
}

function cm_calcConsecutiveDays(scriptFn, todayExceptions) {
  try {
    if (todayExceptions===0) return 0;
    var sh=cm_getCacheSheet(), lr=sh.getLastRow();
    if (lr<2) return 1;
    var vals=sh.getRange(2,1,lr-1,5).getValues();
    var yesterday=cm_fmtDate(new Date(new Date().getTime()-86400000));
    for (var i=0;i<vals.length;i++) {
      if (vals[i][0]===scriptFn && vals[i][1]===yesterday)
        return (parseInt(vals[i][4])||0)+1;
    }
    return 1;
  } catch(e) { return 1; }
}

function cm_cleanOldEntries(sh) {
  try {
    var lr=sh.getLastRow(); if (lr<2) return;
    var cutoff=new Date(new Date().getTime()-7*86400000);
    var vals=sh.getRange(2,1,lr-1,2).getValues();
    for (var i=vals.length-1;i>=0;i--) {
      var d=cm_parseDate(vals[i][1]);
      if (d && d<cutoff) sh.deleteRow(i+2);
    }
  } catch(e) {}
}

// ── SL Config ─────────────────────────────────────────────────────────────────
function cm_readSlConfig() {
  try {
    var cfg=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CM_TAB_CONFIG);
    if (!cfg) return [];
    var vals=cfg.getRange(1,1,cfg.getLastRow(),3).getValues();
    var result=[], inSection=false;
    for (var i=0;i<vals.length;i++) {
      var a=String(vals[i][0]).trim();
      if (a==="SL_CONFIG") { inSection=true; continue; }
      if (inSection && (!a || a.indexOf("_")===0)) break;
      if (inSection && a) result.push({fn:a, sl:vals[i][1], name:String(vals[i][2]).trim()});
    }
    return result;
  } catch(e) { return []; }
}

function cm_getSlInfoForScript(scriptFn) {
  var cfg=cm_readSlConfig();
  for (var i=0;i<cfg.length;i++) if (cfg[i].fn===scriptFn) return cfg[i];
  return null;
}

function cm_getScriptDisplayName(scriptFn) {
  var info=cm_getSlInfoForScript(scriptFn);
  return info ? info.name : scriptFn;
}

// ── Status Dashboard ──────────────────────────────────────────────────────────
function cm_updateStatusDashboard(scriptFn, exceptionsCount) {
  try {
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sh=ss.getSheetByName(CM_TAB_STATUS); if (!sh) return;
    var name=cm_getScriptDisplayName(scriptFn);
    var ts=Utilities.formatDate(new Date(),"Asia/Kolkata","dd.MM.yyyy HH:mm");
    var lr=sh.getLastRow(), found=-1;
    // Data rows start at row 4 (row1=title, row2=spacer, row3=header)
    if (lr>=4) {
      var vals=sh.getRange(4,1,lr-3,1).getValues();
      for (var i=0;i<vals.length;i++) if (vals[i][0]===name) { found=i+4; break; }
    }
    var row=[name,ts,exceptionsCount,"✓ Fresh"];
    if (found>0) sh.getRange(found,1,1,4).setValues([row]);
    else { sh.appendRow(row); found=sh.getLastRow(); }
    sh.getRange(found,1,1,3).setBackground(CM_C.DATA_BG).setFontColor(CM_C.DATA_FG);
    sh.getRange(found,4).setBackground(CM_C.ST_FRESH_BG).setFontColor(CM_C.ST_FRESH_FG).setFontWeight("bold");
  } catch(e) { console.log("cm_updateStatusDashboard error: "+e.message); }
}

function cm_refreshAllStatuses() {
  try {
    var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CM_TAB_STATUS);
    if (!sh) return;
    var lr=sh.getLastRow();
    // Data rows start at row 4 (row1=title, row2=spacer, row3=header)
    if (lr<4) return;
    var dataRows=lr-3;
    var vals=sh.getRange(4,1,dataRows,4).getValues(), now=new Date();
    for (var i=0;i<vals.length;i++) {
      // Skip blank rows and the header row if accidentally included
      if (!vals[i][0] || vals[i][0]==="Script / Report") continue;
      var d=cm_parseDate(String(vals[i][1]).split(' ')[0]);
      var status, bg, fg;
      if (!d) { status="✗ NEVER"; bg=CM_C.ST_EXP_BG; fg=CM_C.ST_EXP_FG; }
      else {
        var days=cm_daysBetween(d,now);
        if (days===0)      { status="✓ Fresh";             bg=CM_C.ST_FRESH_BG; fg=CM_C.ST_FRESH_FG; }
        else if (days===1) { status="⚠ Yesterday";         bg=CM_C.ST_YEST_BG;  fg=CM_C.ST_YEST_FG;  }
        else if (days<=3)  { status="✗ Stale ("+days+"d)"; bg=CM_C.ST_STALE_BG; fg=CM_C.ST_STALE_FG; }
        else               { status="✗ EXPIRED";           bg=CM_C.ST_EXP_BG;   fg=CM_C.ST_EXP_FG;   }
      }
      // i=0 → row 4, so actual row = i+4
      sh.getRange(i+4,4).setValue(status).setBackground(bg).setFontColor(fg).setFontWeight("bold");
    }
  } catch(e) { console.log("cm_refreshAllStatuses error: "+e.message); }
}

// ── Weekly Cache ──────────────────────────────────────────────────────────────
function cm_writeWeeklyCache(scriptFn, summaryObj) {
  try {
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sh=ss.getSheetByName(CM_TAB_WEEKLY_CACHE);
    if (!sh) {
      sh=ss.insertSheet(CM_TAB_WEEKLY_CACHE);
      sh.appendRow(["Script Function","MON","TUE","WED","THU","FRI","SAT","Week Start"]);
      sh.hideSheet();
    }
    var day=["SUN","MON","TUE","WED","THU","FRI","SAT"][new Date().getDay()];
    var weekStart=cm_getWeekStart(), jsonStr=JSON.stringify(summaryObj);
    var dayColMap={MON:1,TUE:2,WED:3,THU:4,FRI:5};
    var dayCol=dayColMap[day]; if (!dayCol) return;
    var lr=Math.max(sh.getLastRow(),1), found=-1;
    if (lr>1) {
      var vals=sh.getRange(2,1,lr-1,8).getValues();
      for (var i=0;i<vals.length;i++) {
        if (vals[i][0]===scriptFn && vals[i][7]===weekStart) { found=i+2; break; }
      }
    }
    if (found>0) {
      sh.getRange(found,dayCol+1).setValue(jsonStr);
    } else {
      var row=new Array(8).fill("");
      row[0]=scriptFn; row[dayCol]=jsonStr; row[7]=weekStart;
      sh.appendRow(row);
    }
  } catch(e) { console.log("cm_writeWeeklyCache error: "+e.message); }
}

function cm_readWeeklyCache() {
  try {
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sh=ss.getSheetByName(CM_TAB_WEEKLY_CACHE); if (!sh) return [];
    var lr=sh.getLastRow(); if (lr<2) return [];
    var weekStart=cm_getWeekStart();
    var vals=sh.getRange(2,1,lr-1,8).getValues(), result=[];
    for (var i=0;i<vals.length;i++) {
      if (vals[i][7]===weekStart) {
        result.push({
          scriptFn : vals[i][0],
          MON : vals[i][1] ? JSON.parse(vals[i][1]) : null,
          TUE : vals[i][2] ? JSON.parse(vals[i][2]) : null,
          WED : vals[i][3] ? JSON.parse(vals[i][3]) : null,
          THU : vals[i][4] ? JSON.parse(vals[i][4]) : null,
          FRI : vals[i][5] ? JSON.parse(vals[i][5]) : null
        });
      }
    }
    return result;
  } catch(e) { return []; }
}

function cm_clearWeeklyCache() {
  try {
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sh=ss.getSheetByName(CM_TAB_WEEKLY_CACHE); if (!sh) return;
    var lr=sh.getLastRow(); if (lr<2) return;
    var weekStart=cm_getWeekStart();
    var vals=sh.getRange(2,1,lr-1,8).getValues();
    for (var i=vals.length-1;i>=0;i--) {
      if (vals[i][7]===weekStart) sh.deleteRow(i+2);
    }
  } catch(e) {}
}

// ── Remarks / Under AEN / Action builders ─────────────────────────────────────
function cm_buildRemarks(summaryObj, scriptName) {
  if (!summaryObj || summaryObj.totalExceptions===0) return "OK";
  var exMap={};
  for (var aen in summaryObj.exceptions) {
    for (var exType in summaryObj.exceptions[aen]) {
      if (!exMap[exType]) exMap[exType]=[];
      exMap[exType].push({aen:aen, count:summaryObj.exceptions[aen][exType]});
    }
  }
  var parts=[];
  for (var t in exMap) {
    var aenParts=exMap[t].map(function(x){ return x.aen+" ("+x.count+" case"+(x.count>1?"s":"")+")" });
    parts.push(t+" under "+aenParts.join(", "));
  }
  return parts.join(". ");
}

function cm_buildUnderAEN(summaryObj) {
  if (!summaryObj || summaryObj.totalExceptions===0) return "";
  var aens=Object.keys(summaryObj.exceptions);
  // Deduplicate and sort
  var unique=[]; aens.forEach(function(a){ if (unique.indexOf(a)<0) unique.push(a); });
  return unique.sort().join(", ");
}

function cm_buildActionTaken(summaryObj) {
  if (!summaryObj || summaryObj.totalExceptions===0) return "";
  var aens=[];
  for (var aen in summaryObj.exceptions) {
    if (aens.indexOf(aen)<0) {
      // Check for repeat defaulter
      aens.push(aen);
    }
  }
  aens.sort();
  return aens.map(function(a){ return a+"-"; }).join("\n");
}

// ── Repair damaged status tab ─────────────────────────────────────────────────
// Run this ONCE if rows 2-3 of SYSTEM STATUS got corrupted
function repairStatusTab() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName(CM_TAB_STATUS); if (!sh) { console.log("SYSTEM STATUS tab not found."); return; }
  // Row 1: title — leave alone
  // Row 2: spacer — clear all content and format
  sh.getRange(2,1,1,4).clearContent().clearFormat().setBackground(CM_C.WHITE);
  // Row 3: header — restore
  sh.getRange(3,1,1,4)
    .setValues([["Script / Report","Last Run","Exceptions Found","Status"]])
    .setBackground(CM_C.HEADER_BG).setFontColor(CM_C.HEADER_FG).setFontWeight("bold");
  // Remove any stray rows that were incorrectly created before row 4
  // (nothing to do — rows 4+ are data and should be fine)
  try { SpreadsheetApp.getUi().alert("SYSTEM STATUS tab repaired. Rows 2-3 restored."); }
  catch(_) { console.log("Status tab repaired."); }
}
function cm_getConsecutiveDaysFromCache(scriptFn) {
  try {
    var sh=cm_getCacheSheet(), lr=sh.getLastRow();
    if (lr<2) return 0;
    var vals=sh.getRange(2,1,lr-1,5).getValues(), today=cm_today();
    for (var i=0;i<vals.length;i++) {
      if (vals[i][0]===scriptFn && vals[i][1]===today) return parseInt(vals[i][4])||0;
    }
    return 0;
  } catch(e) { return 0; }
}