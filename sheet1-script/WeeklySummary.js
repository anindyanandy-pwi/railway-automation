// ============================================================
//  WeeklySummary.gs — Sheet 1 ONLY
//  Generates weekly exception summary every Friday at 9 AM.
//  Results prepended above previous weeks in WEEKLY SUMMARY tab.
// ============================================================

function generateWeeklySummary() {
  try {
    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var sh       = ss.getSheetByName(CM_TAB_WEEKLY_SUM);
    if (!sh) sh  = ss.insertSheet(CM_TAB_WEEKLY_SUM);

    var weekData = cm_readWeeklyCache();
    if (!weekData || weekData.length === 0) {
      console.log("Weekly cache empty — no summary generated.");
      return;
    }

    var weekStart = cm_getWeekStart();
    var weekEnd   = cm_fmtDate(new Date()); // Friday
    var slConfig  = cm_readSlConfig();
    var days      = ["MON","TUE","WED","THU","FRI"];

    // Build rows: [ [Item, Mon, Tue, Wed, Thu, Fri, Total, Repeat Flag] ]
    var rows = _ws_buildSummaryRows(weekData, slConfig, days);

    // Build repeat defaulter section
    var repeatRows = _ws_buildRepeatDefaulters(weekData, slConfig);

    // Insert block at row 3 (below title row 1 and subtitle row 2)
    var existingRows = sh.getLastRow();
    var blockSize    = 4 + rows.length + 2 + repeatRows.length + 2; // header + data + separator + repeats + gap
    sh.insertRowsBefore(3, blockSize);

    var writeRow = 3;

    // Week title
    var titleRange = sh.getRange(writeRow, 1, 1, 8);
    try { titleRange.merge(); } catch(_) {}
    titleRange.setValue("WEEK: " + weekStart + " — " + weekEnd)
      .setBackground(CM_C.TITLE_BG).setFontColor(CM_C.TITLE_FG)
      .setFontWeight("bold").setHorizontalAlignment("center").setFontSize(11);
    writeRow++;

    // Column headers
    sh.getRange(writeRow, 1, 1, 8).setValues([["ITEM","MON","TUE","WED","THU","FRI","TOTAL","REPEAT FLAG"]])
      .setBackground(CM_C.HEADER_BG).setFontColor(CM_C.HEADER_FG).setFontWeight("bold");
    writeRow++;

    // Data rows
    rows.forEach(function(row) {
      var dataRange = sh.getRange(writeRow, 1, 1, 8);
      dataRange.setValues([row]);
      var total = row[6];
      if (total === 0) {
        dataRange.setBackground(CM_C.OK_BG).setFontColor(CM_C.OK_FG);
      } else {
        dataRange.setBackground(CM_C.DATA_BG).setFontColor(CM_C.DATA_FG);
        if (row[7]) { // repeat flag
          sh.getRange(writeRow, 8).setBackground(CM_C.EXCEPT_BG).setFontColor(CM_C.EXCEPT_FG).setFontWeight("bold");
        }
      }
      writeRow++;
    });

    // Repeat defaulters section
    if (repeatRows.length > 0) {
      writeRow++; // gap
      sh.getRange(writeRow, 1, 1, 8).setValues([["REPEAT DEFAULTERS THIS WEEK","","","","","","",""]])
        .setBackground(CM_C.EXCEPT_BG).setFontColor(CM_C.EXCEPT_FG).setFontWeight("bold");
      try { sh.getRange(writeRow, 1, 1, 8).merge(); } catch(_) {}
      writeRow++;
      repeatRows.forEach(function(row) {
        sh.getRange(writeRow, 1, 1, 8).setValues([row])
          .setBackground(CM_C.DATA_BG).setFontColor(CM_C.DATA_FG);
        writeRow++;
      });
    }

    // Gap row
    sh.getRange(writeRow, 1, 1, 8).clearContent().clearFormat()
      .setBackground(CM_C.WHITE);

    // Column widths
    sh.setColumnWidth(1, 280);
    for (var c=2; c<=7; c++) sh.setColumnWidth(c, 75);
    sh.setColumnWidth(8, 200);

    // Clear weekly cache ONLY after successful write
    cm_clearWeeklyCache();
    console.log("Weekly summary generated for week " + weekStart + " — " + weekEnd);

    // Notification
    try {
      GmailApp.sendEmail(CM_PARENT_EMAIL,
        "Weekly Exception Summary Generated — " + weekEnd,
        "The weekly exception summary for " + weekStart + " to " + weekEnd +
        " has been generated and added to the WEEKLY SUMMARY tab in Sheet 1.",
        { name: "Exception Report System — Howrah Division" });
    } catch(_) {}
  } catch(e) {
    console.log("generateWeeklySummary error: "+e.message);
    try {
      GmailApp.sendEmail(CM_PARENT_EMAIL,
        "ERROR: Weekly Summary Failed — " + cm_today(),
        "Weekly summary generation failed.\n\nError: " + e.message +
        "\n\nWeekly cache has NOT been cleared.");
    } catch(_) {}
  }
}

function _ws_buildSummaryRows(weekData, slConfig, days) {
  // Group by SL
  var slGroups = {};
  slConfig.forEach(function(cfg) {
    if (!slGroups[cfg.sl]) slGroups[cfg.sl] = { name: cfg.name, fns: [] };
    slGroups[cfg.sl].fns.push(cfg.fn);
  });

  // Build cache lookup: fn -> { MON: summary, TUE: summary, ... }
  var fnCache = {};
  weekData.forEach(function(w) { fnCache[w.scriptFn] = w; });

  var rows = [];
  Object.keys(slGroups).sort(function(a,b){return parseInt(a)-parseInt(b);}).forEach(function(sl) {
    var group = slGroups[sl];
    var dayCounts = {};
    days.forEach(function(day) { dayCounts[day] = 0; });
    var maxConsec = 0;

    group.fns.forEach(function(fn) {
      var wData = fnCache[fn];
      if (!wData) return;
      days.forEach(function(day) {
        var s = wData[day];
        if (s && s.totalExceptions) dayCounts[day] += s.totalExceptions;
      });
    });

    var total = days.reduce(function(sum,d){ return sum+dayCounts[d]; }, 0);
    // Get max consecutive days from cache
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CM_TAB_CACHE);
    if (sh && sh.getLastRow()>1) {
      var vals = sh.getRange(2,1,sh.getLastRow()-1,5).getValues();
      group.fns.forEach(function(fn) {
        vals.forEach(function(row) {
          if (row[0]===fn) maxConsec = Math.max(maxConsec, parseInt(row[4])||0);
        });
      });
    }

    var repeatFlag = maxConsec >= 3 ? "⚠ " + maxConsec + " consecutive days" : "";
    rows.push([
      group.name,
      dayCounts.MON, dayCounts.TUE, dayCounts.WED, dayCounts.THU, dayCounts.FRI,
      total, repeatFlag
    ]);
  });
  return rows;
}

function _ws_buildRepeatDefaulters(weekData, slConfig) {
  // Find AENs that appeared in exceptions on 3+ days
  var aenDays = {}; // aen -> Set of days
  var aenScripts = {}; // aen -> [scriptNames]

  weekData.forEach(function(w) {
    var scriptName = cm_getScriptDisplayName(w.scriptFn);
    ["MON","TUE","WED","THU","FRI"].forEach(function(day) {
      var s = w[day];
      if (!s || !s.exceptions) return;
      for (var aen in s.exceptions) {
        if (!aenDays[aen]) aenDays[aen] = {};
        aenDays[aen][day] = true;
        if (!aenScripts[aen]) aenScripts[aen] = {};
        aenScripts[aen][scriptName] = (aenScripts[aen][scriptName]||0)+1;
      }
    });
  });

  var rows = [];
  Object.keys(aenDays).sort().forEach(function(aen) {
    var days = Object.keys(aenDays[aen]).length;
    if (days >= 3) {
      var scripts = Object.keys(aenScripts[aen]).map(function(s){
        return s + " ("+aenScripts[aen][s]+" day"+(aenScripts[aen][s]>1?"s":"")+")" ;
      });
      rows.push([aen + " — " + scripts.join(", "), days + " days",
                 "","","","","",""]);
    }
  });
  return rows;
}