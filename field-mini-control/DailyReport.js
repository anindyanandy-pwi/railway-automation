// ============================================================
//  DailyReport.gs — Sheet 1 ONLY
//  Generates the combined daily exception report block in the
//  DAILY EXCEPTION REPORT tab.
//
//  Two-click process:
//    Click 1: Check cache; if stale, schedule scripts in bg
//    Click 2: Cache fresh; overwrite check; populate report
//
//  Report date = yesterday (reportDate)
//  Cache date  = today (scripts ran today)
// ============================================================

var DR_STALE_LIMIT_DAYS = 3;

// ── Color scheme — exact match to Excel target format ────────────────────────
var DR_C = {
  TITLE_BG      : "#EAD1DC",  // Pink rose — title row
  HEADER_BG     : "#B6D7A8",  // Light green — header row
  SL_ITEM_BG    : "#00FF00",  // Bright green — cols A & B (daily items)
  CD_BG         : "#C9DAF8",  // Light blue — cols C & D
  WEEKLY_BG     : "#F1C232",  // Yellow — weekly items
  EXCEPTION_F_BG: "#FF0000",  // Red — Under AEN cell when exception
  NO_FILL       : "#FFFFFF",  // White — Remarks & Action columns
  OVERWRITE     : "#FEFCE8",  // Yellow highlight — overwritten cells
  BLACK         : "#000000",
  FONT          : "Arial",
  DATA_SIZE     : 10,
  TITLE_SIZE    : 13
};

// Script display names for combined SL summaries
var DR_SCRIPT_SHORT_NAMES = {
  "generateROBReport"    : "ROB",
  "generateFOBReport"    : "FOB",
  "generateBridgeReport" : "Bridge",
  "generateOMSReport"    : "OMS Peaks",
  "generateTRCReport"    : "TRC UML Peaks"
};

// ── Helper to apply DR_C border + font consistently ───────────────────────────
function _drFormat(range, bg, bold, hAlign, vAlign, fontSize) {
  range.setBackground(bg || DR_C.NO_FILL)
       .setFontColor(DR_C.BLACK)
       .setFontFamily(DR_C.FONT)
       .setFontSize(fontSize || DR_C.DATA_SIZE)
       .setFontWeight(bold ? "bold" : "normal")
       .setWrap(true);
  if (hAlign) range.setHorizontalAlignment(hAlign);
  if (vAlign) range.setVerticalAlignment(vAlign);
  range.setBorder(true, true, true, true, true, true,
    DR_C.BLACK, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

// ── Main entry point (menu click) ────────────────────────────────────────────
function generateCombinedDailyReport() {
  var ui         = SpreadsheetApp.getUi();
  var today      = cm_today();  // actual today — for cache checks
  var reportDate = cm_fmtDate(new Date(new Date().getTime() - 86400000)); // yesterday — block title

  // ── Check Sheet 1 cache freshness (today's data) ──────────────────────────
  var s1Scripts = ["generateBadRoadReport","generateTWSTieBarReport",
                   "generateGaplessReport","generateGJointReport"];
  var s1Fresh = s1Scripts.every(function(fn) {
    var d = cm_getCacheDateForScript(fn);
    return d && cm_fmtDate(d) === today;
  });

  if (!s1Fresh) {
    ui.alert(
      "Data Not Ready",
      "Sheet 1 exception data is not yet generated for today.\n\n" +
      "All 4 Sheet 1 reports will now run in the background.\n" +
      "This will take approximately 8-10 minutes.\n\n" +
      "You will receive an email at " + CM_PARENT_EMAIL + " when ready.\n\n" +
      "Then click 'Generate Combined Daily Report' again.\n\n" +
      "Report will be dated: " + reportDate,
      ui.ButtonSet.OK
    );
    _dr_scheduleCombinedChain();
    return;
  }

  // ── Click 2: cache is fresh ───────────────────────────────────────────────
  var cacheStatus = _dr_checkCrossSheetCache();
  var proceedMsg  = _dr_buildCacheStatusMessage(cacheStatus, today);
  if (proceedMsg) {
    var resp = ui.alert("Cache Status Check", proceedMsg, ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
  }

  var dailySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CM_TAB_DAILY_RPT);
  if (!dailySheet) {
    ui.alert("Error", "Tab '" + CM_TAB_DAILY_RPT + "' not found.", ui.ButtonSet.OK);
    return;
  }

  var existingBlock = _dr_findBlockForDate(dailySheet, reportDate);
  var overwriteApproved = false;

  if (existingBlock && existingBlock.hasData) {
    var resp2 = ui.alert(
      "Data Already Exists",
      "Report block for " + reportDate + " already has data:\n\n" +
      existingBlock.existingFields.join("\n") + "\n\n" +
      "Overwrite? (Overwritten cells will be highlighted yellow)",
      ui.ButtonSet.YES_NO
    );
    if (resp2 !== ui.Button.YES) return;
    overwriteApproved = true;
  }

  PropertiesService.getScriptProperties()
    .setProperty("DR_OVERWRITE_APPROVED", overwriteApproved ? "YES" : "NO");
  PropertiesService.getScriptProperties()
    .setProperty("DR_TARGET_DATE", reportDate);

  _dr_populateDailyReport(dailySheet, reportDate, overwriteApproved, cacheStatus);

  ui.alert("Done!",
    "Combined Daily Exception Report populated for " + reportDate + ".\n\n" +
    "AEN-specific emails will be sent in the background shortly.",
    ui.ButtonSet.OK);
}

// ── Background chain for Click 1 ─────────────────────────────────────────────
var DR_CHAIN_FNS = ["s1_BadRoadSurface","s1_TWSTieBar","s1_GaplessCMS","s1_GJoint"];

function _dr_scheduleCombinedChain() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "_dr_continueChain") ScriptApp.deleteTrigger(t);
  });
  PropertiesService.getScriptProperties()
    .setProperty("DR_COMBINED_QUEUE", JSON.stringify(DR_CHAIN_FNS.slice()));
  ScriptApp.newTrigger("_dr_continueChain")
    .timeBased().at(new Date(Date.now() + 30000)).create();
}

function _dr_continueChain() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "_dr_continueChain") ScriptApp.deleteTrigger(t);
  });
  var prop  = PropertiesService.getScriptProperties();
  var raw   = prop.getProperty("DR_COMBINED_QUEUE");
  if (!raw) return;
  var queue = JSON.parse(raw);
  if (!queue || queue.length === 0) { prop.deleteProperty("DR_COMBINED_QUEUE"); return; }
  var fnName = queue.shift();
  prop.setProperty("DR_COMBINED_QUEUE", JSON.stringify(queue));
  try { _sCallFn(fnName); } catch(e) { console.log("Chain error in "+fnName+": "+e.message); }
  if (queue.length > 0) {
    ScriptApp.newTrigger("_dr_continueChain")
      .timeBased().at(new Date(Date.now() + 90000)).create();
  } else {
    prop.deleteProperty("DR_COMBINED_QUEUE");
    var reportDate = cm_fmtDate(new Date(new Date().getTime() - 86400000));
    try {
      GmailApp.sendEmail(CM_PARENT_EMAIL,
        "Exception Report Data Ready — " + cm_today(),
        "All Sheet 1 exception reports generated for " + cm_today() + ".\n\n" +
        "Report will be dated: " + reportDate + "\n\n" +
        "Ensure Sheet 2 and Sheet 3 reports have also been run today.\n\n" +
        "You can now click 'Generate Combined Daily Report' in Sheet 1.",
        { name: "Exception Report System — Howrah Division" });
    } catch(e) { console.log("Notify email error: "+e.message); }
  }
}

// ── Cross-sheet cache check ───────────────────────────────────────────────────
function _dr_checkCrossSheetCache() {
  var today = cm_today();
  var cacheSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CM_TAB_CACHE);
  if (!cacheSheet) return {};
  var allScripts = cm_readSlConfig().map(function(r){ return r.fn; });
  var status = {}, cacheData = {};
  var lr = Math.max(cacheSheet.getLastRow(), 1);
  if (lr > 1) {
    var vals = cacheSheet.getRange(2,1,lr-1,6).getValues();
    vals.forEach(function(row) {
      var fn = row[0], date = row[1];
      if (!cacheData[fn] || date > cacheData[fn].date)
        cacheData[fn] = { date: date, ts: row[5] };
    });
  }
  allScripts.forEach(function(fn) {
    var entry = cacheData[fn];
    if (!entry) { status[fn] = { date:null, age:999, ok:false }; }
    else {
      var d = cm_parseDate(entry.date);
      var age = d ? cm_daysBetween(d, new Date()) : 999;
      status[fn] = { date:entry.date, ts:entry.ts, age:age, ok:age<=DR_STALE_LIMIT_DAYS };
    }
  });
  return status;
}

function _dr_buildCacheStatusMessage(status, today) {
  var stale=[], expired=[];
  for (var fn in status) {
    var s = status[fn];
    if (s.age > DR_STALE_LIMIT_DAYS) expired.push(fn+" (last: "+(s.date||"never")+")");
    else if (s.date && s.date !== today) stale.push(fn+" (last: "+s.date+", "+s.age+"d ago)");
  }
  var lines=[];
  if (expired.length) lines.push("⚠ EXPIRED (over "+DR_STALE_LIMIT_DAYS+"d limit):\n"+expired.join("\n"));
  if (stale.length)   lines.push("⚠ STALE (not today — using anyway):\n"+stale.join("\n"));
  if (!lines.length)  return null;
  return lines.join("\n\n")+"\n\nProceed with available data? Stale cells will be highlighted yellow.";
}

// ── Find block for a specific date ────────────────────────────────────────────
function _dr_findBlockForDate(sheet, dateStr) {
  var lr = sheet.getLastRow(); if (lr < 1) return null;
  var colA = sheet.getRange(1,1,lr,1).getValues();
  for (var i = 0; i < colA.length; i++) {
    var v = String(colA[i][0]);
    if (v.indexOf(dateStr) > -1 && v.indexOf("Combined") > -1) {
      var blockStart = i + 1;
      var endRow = _dr_findBlockEnd(colA, blockStart, dateStr);
      var dataStart = blockStart + 2;
      var existingFields = [];
      for (var r = dataStart; r <= endRow; r++) {
        var sl = String(colA[r-1][0]).trim();
        if (!sl || isNaN(parseInt(sl))) continue;
        var rowVals = sheet.getRange(r,5,1,3).getValues()[0];
        if (rowVals[0]||rowVals[1]||rowVals[2])
          existingFields.push("  SL "+sl+": "+["Remarks","Under AEN","Action"]
            .filter(function(_,j){ return rowVals[j]; }).join(", "));
      }
      return { startRow:blockStart, endRow:endRow,
               hasData:existingFields.length>0, existingFields:existingFields };
    }
  }
  return null;
}

function _dr_findBlockEnd(colA, startRow, currentDate) {
  for (var i = startRow; i < colA.length; i++) {
    var v = String(colA[i][0]);
    if (v.indexOf("Combined Mini Control Exception Report") > -1 &&
        v.indexOf(currentDate) < 0 && i > startRow) return i;
  }
  return colA.length;
}

// ── Find most recent block (for template) ─────────────────────────────────────
function _dr_findMostRecentBlock(sheet) {
  var lr = sheet.getLastRow(); if (lr < 1) return null;
  var reportDate = cm_fmtDate(new Date(new Date().getTime()-86400000));
  var block = _dr_findBlockForDate(sheet, reportDate);
  if (block) return block;
  var colA = sheet.getRange(1,1,lr,1).getValues();
  var latest = null, latestDate = null;
  for (var i = 0; i < colA.length; i++) {
    var v = String(colA[i][0]);
    if (v.indexOf("Combined Mini Control Exception Report, Date:") > -1) {
      var m = v.match(/Date:\s*(\d{2}\.\d{2}\.\d{4})/);
      if (m) {
        var d = cm_parseDate(m[1]);
        if (!latestDate || d > latestDate) { latestDate=d; latest=i+1; }
      }
    }
  }
  if (!latest) return null;
  var end = _dr_findBlockEnd(colA, latest, latestDate ? cm_fmtDate(latestDate) : "");
  return { startRow:latest, endRow:end };
}

// ── Create new block at top ───────────────────────────────────────────────────
function _dr_createNewBlock(sheet, reportDate) {
  var templateBlock = _dr_findMostRecentBlock(sheet);
  var numRows = templateBlock ? (templateBlock.endRow - templateBlock.startRow + 1) : 80;
  numRows = Math.max(numRows, 5);
  var insertCount = numRows + 2;
  sheet.insertRowsBefore(1, insertCount);

  // Row 1: Title — merged A-G
  var titleRange = sheet.getRange(1,1,1,7);
  titleRange.clearFormat();
  try { titleRange.merge(); } catch(_) {}
  titleRange.setValue("Combined Mini Control Exception Report, Date: " + reportDate);
  _drFormat(titleRange, DR_C.TITLE_BG, true, "center", "center", DR_C.TITLE_SIZE);

  // Row 2: Header
  var hdrRange = sheet.getRange(2,1,1,7);
  hdrRange.clearFormat();
  hdrRange.setValues([["SL","Items","Daily/ Weekly","Looked after by",
                        "Remarks for Exception","Under AEN","Action taken By AEN"]]);
  _drFormat(hdrRange, DR_C.HEADER_BG, true, "center", "center", DR_C.DATA_SIZE);
  // G header: top align
  sheet.getRange(2,7).setVerticalAlignment("top");

  // Copy data rows from template (values + formatting)
  if (templateBlock) {
    var origDataStart = templateBlock.startRow + 2;
    var origDataStart_shifted = origDataStart + insertCount;
    var dataRows = templateBlock.endRow - templateBlock.startRow - 1;

    if (dataRows > 0) {
      try {
        // Copy A-D with formatting (colors from template are correct)
        var srcRange  = sheet.getRange(origDataStart_shifted, 1, dataRows, 4);
        var destRange = sheet.getRange(3, 1, dataRows, 4);
        srcRange.copyTo(destRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

        // Apply correct font & border to A-D (copyTo may not get border right)
        for (var i = 0; i < dataRows; i++) {
          var row = sheet.getRange(3+i, 1, 1, 4);
          row.setFontFamily(DR_C.FONT).setFontSize(DR_C.DATA_SIZE).setWrap(true);
          row.setBorder(true,true,true,true,true,true,
            DR_C.BLACK, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        }
        // Col A: center align
        sheet.getRange(3,1,dataRows,1).setHorizontalAlignment("center")
             .setVerticalAlignment("center");
        // Col B: left align, center vertical
        sheet.getRange(3,2,dataRows,1).setHorizontalAlignment("left")
             .setVerticalAlignment("center");
        // Col C: center align
        sheet.getRange(3,3,dataRows,1).setHorizontalAlignment("center")
             .setVerticalAlignment("center");
        // Col D: center align
        sheet.getRange(3,4,dataRows,1).setHorizontalAlignment("center")
             .setVerticalAlignment("center");

        // E-G: clear and set format (fresh for exception data)
        var efgRange = sheet.getRange(3, 5, dataRows, 3);
        efgRange.clearContent().clearFormat();
        _drFormat(efgRange, DR_C.NO_FILL, false, null, "center", DR_C.DATA_SIZE);
        // G: top align
        sheet.getRange(3, 7, dataRows, 1).setVerticalAlignment("top");

      } catch(e) {
        console.log("Template copy error: "+e.message);
        _dr_buildDefaultDataRows(sheet, reportDate, numRows-2);
      }
    }
  } else {
    _dr_buildDefaultDataRows(sheet, reportDate, 76);
  }

  // Gap row after block
  try {
    sheet.getRange(insertCount, 1, 1, 7).clearContent().clearFormat()
      .setBackground(DR_C.NO_FILL);
  } catch(_) {}

  return { startRow:1, dataStartRow:3 };
}

function _dr_buildDefaultDataRows(sheet, reportDate, count) {
  for (var i = 0; i < count; i++) {
    var rowRange = sheet.getRange(3+i, 1, 1, 7);
    rowRange.clearFormat();
    // A: SL number, green
    sheet.getRange(3+i,1).setValue(i+1);
    _drFormat(sheet.getRange(3+i,1), DR_C.SL_ITEM_BG, true, "center", "center");
    // B: item name, green
    sheet.getRange(3+i,2).setValue("[Item "+(i+1)+" — fill manually]");
    _drFormat(sheet.getRange(3+i,2), DR_C.SL_ITEM_BG, true, "left", "center");
    // C-D: blue
    _drFormat(sheet.getRange(3+i,3,1,2), DR_C.CD_BG, true, "center", "center");
    // E-G: white
    _drFormat(sheet.getRange(3+i,5,1,3), DR_C.NO_FILL, false, null, "center");
    sheet.getRange(3+i,7).setVerticalAlignment("top");
  }
}

// ── Main populate function ────────────────────────────────────────────────────
function _dr_populateDailyReport(dailySheet, reportDate, overwriteApproved, cacheStatus) {
  var block = _dr_findBlockForDate(dailySheet, reportDate);
  if (!block) {
    _dr_createNewBlock(dailySheet, reportDate);
    block = _dr_findBlockForDate(dailySheet, reportDate);
    if (!block) { console.log("Could not create or find block for "+reportDate); return; }
  }

  // ── CRITICAL FIX: read cache using TODAY (when scripts ran), not reportDate ──
  var cacheDate = cm_today();
  var allCache  = cm_readAllCacheForDate(cacheDate);
  var slConfig  = cm_readSlConfig();

  // Group scripts by SL number
  var slGroups = {};
  slConfig.forEach(function(cfg) {
    if (!slGroups[cfg.sl]) slGroups[cfg.sl] = [];
    slGroups[cfg.sl].push({ fn:cfg.fn, name:cfg.name, cache:allCache[cfg.fn]||null });
  });

  var lr    = dailySheet.getLastRow();
  var colA  = dailySheet.getRange(1,1,lr,1).getValues();
  var dataStartRow = block.startRow + 2;

  for (var sl in slGroups) {
    var group    = slGroups[sl];
    var slInt    = parseInt(sl);
    var targetRow = _dr_findSLRow(dailySheet, colA, dataStartRow, block.endRow, slInt, group[0].name);
    if (targetRow < 0) {
      try {
        GmailApp.sendEmail(CM_PARENT_EMAIL,
          "WARNING: SL not found in Daily Report — " + reportDate,
          "SL "+sl+" ("+group[0].name+") not found in block for "+reportDate+
          ".\nFallback also failed. Check _CONFIG_ tab SL numbers.");
      } catch(_) {}
      continue;
    }

    var remarks     = _dr_buildGroupRemarks(group, sl);
    var underAEN    = _dr_buildGroupUnderAEN(group);
    var actionTaken = _dr_buildGroupAction(group);
    var isOK        = (remarks === "OK");
    var isStale     = _dr_isGroupStale(group, cacheStatus);

    var eCell = dailySheet.getRange(targetRow, 5);
    var fCell = dailySheet.getRange(targetRow, 6);
    var gCell = dailySheet.getRange(targetRow, 7);

    var eVal = eCell.getValue(), fVal = fCell.getValue(), gVal = gCell.getValue();
    var hadData = !!(eVal||fVal||gVal);
    var markYellow = hadData && overwriteApproved;

    if (isOK) {
      // ── OK row: "OK" in E only, no merge, no fill ─────────────────────────
      // First unmerge if previously merged
      try { dailySheet.getRange(targetRow,5,1,3).breakApart(); } catch(_) {}
      eCell.setValue("OK");
      _drFormat(eCell, DR_C.NO_FILL, false, "center", "center");
      fCell.clearContent();
      _drFormat(fCell, DR_C.NO_FILL, false, "center", "center");
      gCell.clearContent();
      _drFormat(gCell, DR_C.NO_FILL, false, null, "top");
    } else {
      var eBg = markYellow ? DR_C.OVERWRITE : (isStale ? DR_C.OVERWRITE : DR_C.NO_FILL);
      var fBg = markYellow ? DR_C.OVERWRITE : DR_C.EXCEPTION_F_BG;
      var gBg = markYellow ? DR_C.OVERWRITE : DR_C.NO_FILL;

      // E: Remarks — white fill
      eCell.setValue(remarks);
      _drFormat(eCell, eBg, false, "left", "center");

      // F: Under AEN — RED fill
      fCell.setValue(underAEN);
      _drFormat(fCell, fBg, true, "center", "center");

      // G: Action taken — white fill, top align
      gCell.setValue(actionTaken);
      _drFormat(gCell, gBg, false, "left", "top");
    }
  }

  console.log("Daily report populated for " + reportDate + " using cache from " + cacheDate);
  _dr_scheduleAENEmails();
}

// ── Find SL row with fallback ─────────────────────────────────────────────────
function _dr_findSLRow(sheet, colA, startRow, endRow, slNo, itemName) {
  for (var i = startRow-1; i < Math.min(endRow, colA.length); i++) {
    if (parseInt(colA[i][0]) === slNo) return i+1;
  }
  var nameLower = itemName.toLowerCase();
  for (var j = startRow-1; j < Math.min(endRow, colA.length); j++) {
    var cellB = String(sheet.getRange(j+1,2).getValue()).toLowerCase();
    if (cellB.indexOf(nameLower)>-1 || nameLower.indexOf(cellB)>-1) {
      console.log("Fallback: SL "+slNo+" found by name at row "+(j+1));
      try {
        GmailApp.sendEmail(CM_PARENT_EMAIL,
          "INFO: SL Fallback Used — " + cm_today(),
          "SL "+slNo+" ('"+itemName+"') found by name fallback at row "+(j+1)+
          ".\nVerify SL number in _CONFIG_ tab.");
      } catch(_) {}
      return j+1;
    }
  }
  return -1;
}

// ── Remarks / Under AEN / Action builders ─────────────────────────────────────
function _dr_buildGroupRemarks(group, sl) {
  var allZero = group.every(function(g){ return !g.cache||g.cache.summary.totalExceptions===0; });
  if (allZero) return "OK";
  if (group.length === 1) {
    var s = group[0].cache ? group[0].cache.summary : null;
    if (!s||s.totalExceptions===0) return "OK";
    return _dr_formatRemarks(s, null, group[0].cache.consecutiveDays||0);
  }
  var parts=[];
  group.forEach(function(g) {
    var shortName = DR_SCRIPT_SHORT_NAMES[g.fn]||g.name;
    var s = g.cache ? g.cache.summary : null;
    if (!s||s.totalExceptions===0) parts.push(shortName+": OK");
    else parts.push(shortName+": "+_dr_formatRemarks(s,null,g.cache.consecutiveDays||0));
  });
  return parts.join(". ");
}

function _dr_formatRemarks(summary, prefix, consecutiveDays) {
  if (!summary||summary.totalExceptions===0) return "OK";
  var exMap={};
  for (var aen in summary.exceptions) {
    for (var exType in summary.exceptions[aen]) {
      if (!exMap[exType]) exMap[exType]=[];
      var count = summary.exceptions[aen][exType];
      var repeat = consecutiveDays>=3 ? " ⚠ REPEAT "+consecutiveDays+"th day" : "";
      exMap[exType].push(aen+" ("+count+" case"+(count>1?"s":"")+")"+repeat);
    }
  }
  var parts=[];
  for (var t in exMap) parts.push(t+" under "+exMap[t].join(", "));
  return (prefix?prefix+": ":"")+parts.join(". ");
}

function _dr_buildGroupUnderAEN(group) {
  var aens={};
  group.forEach(function(g) {
    if (!g.cache||g.cache.summary.totalExceptions===0) return;
    for (var aen in g.cache.summary.exceptions) aens[aen]=true;
  });
  return Object.keys(aens).sort().join(", ");
}

function _dr_buildGroupAction(group) {
  var aens={};
  group.forEach(function(g) {
    if (!g.cache||g.cache.summary.totalExceptions===0) return;
    for (var aen in g.cache.summary.exceptions) {
      for (var exType in g.cache.summary.exceptions[aen]) {
        if (!aens[aen]) aens[aen]=[];
        aens[aen].push(exType);
      }
    }
  });
  var lines=[];
  Object.keys(aens).sort().forEach(function(aen) {
    aens[aen].forEach(function(exType){ lines.push(aen+"- ("+exType+")"); });
  });
  return lines.join("\n");
}

function _dr_isGroupStale(group, cacheStatus) {
  return group.some(function(g) {
    var s = cacheStatus[g.fn];
    return s && s.age > 0;
  });
}

// ── AEN-specific email chain ──────────────────────────────────────────────────
function _dr_scheduleAENEmails() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction()==="_dr_sendAENEmails") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("_dr_sendAENEmails")
    .timeBased().at(new Date(Date.now()+120000)).create();
}

function _dr_sendAENEmails() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction()==="_dr_sendAENEmails") ScriptApp.deleteTrigger(t);
  });

  var today      = cm_today();
  var reportDate = cm_fmtDate(new Date(new Date().getTime()-86400000));
  var allCache   = cm_readAllCacheForDate(today);

  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName(CM_TAB_CONFIG);
  var aenEmails = {};
  if (cfg) {
    var vals = cfg.getRange(1,1,cfg.getLastRow(),6).getValues();
    var inSection=false;
    for (var i=0;i<vals.length;i++) {
      if (vals[i][0]==="AEN_EMAILS") { inSection=true; continue; }
      if (inSection&&(!vals[i][0]||String(vals[i][0]).indexOf("_")===0)) break;
      if (inSection&&vals[i][0]) {
        var emails=[];
        for (var c=1;c<6;c++) { if (vals[i][c]) emails.push(String(vals[i][c]).trim()); }
        if (emails.length) aenEmails[String(vals[i][0]).trim()]=emails;
      }
    }
  }

  var aenData={};
  for (var fn in allCache) {
    var summary=allCache[fn].summary;
    if (!summary||summary.totalExceptions===0) continue;
    for (var aen in summary.exceptions) {
      if (!aenData[aen]) aenData[aen]=[];
      var scriptName=cm_getScriptDisplayName(fn);
      for (var exType in summary.exceptions[aen]) {
        aenData[aen].push("• "+scriptName+" — "+exType+
          " ("+summary.exceptions[aen][exType]+" case"+
          (summary.exceptions[aen][exType]>1?"s":"")+")");
      }
    }
  }

  var sent=0;
  for (var aen in aenData) {
    var emailList=aenEmails[aen];
    if (!emailList||emailList.length===0) continue;
    var body="Dear "+aen+",\n\n"+
      "The following exceptions were recorded under your jurisdiction for "+reportDate+":\n\n"+
      aenData[aen].join("\n")+"\n\n"+
      "Please take necessary action and update the remarks in the Daily Exception Report.\n"+
      "Direct link: https://docs.google.com/spreadsheets/d/"+CM_SHEET1_ID+"/edit#gid=855312642\n\n"+
      "Regards,\nException Report System\nHowrah Division, Eastern Railway";
    try {
      GmailApp.sendEmail(emailList.join(","),
        "Exception Report — "+aen+" — "+reportDate, body,
        { name:"Exception Report System — Howrah Division" });
      sent++;
    } catch(e) { console.log("AEN email error for "+aen+": "+e.message); }
  }

  try {
    GmailApp.sendEmail(CM_PARENT_EMAIL,
      "All AEN-specific Exception Emails Sent — "+reportDate,
      "All "+sent+" AEN-specific exception emails sent for "+reportDate+".",
      { name:"Exception Report System — Howrah Division" });
  } catch(_) {}
  console.log("AEN emails sent: "+sent);
}