// ============================================================
//  CompliancePWI_Exception.gs — Sheet 1
//  Exception scanner for Compliance by PWI sheet.
//
//  EXCEPTION TYPES:
//  Ex1: Pending > 15 days, TMS compliance date still blank
//  Ex2: TMS date filled, actual work+photo missing after 10 days (day 11+)
//  Ex3: Photo missing — actual date filled but photo column blank
//  Ex4: Actual date filled, TMS date blank (> 3-day grace)
//  Ex5: Data quality — Pending Since date missing but row has data
//
//  SCAN RANGE RULE:
//  If today ≤ 7: scan 1st of previous month to today
//  If today > 7: scan 1st of current month to today
//
//  Integrates with: Cache system → SYSTEM STATUS → SL 40 in daily report
// ============================================================

var PWI_EX_SECTION_ID = "COMPLIANCE_PWI";

// ── Re-use helpers from CompliancePWI_Fetch.gs ────────────────────────────────
// _pwi_parseDate, _pwi_fmtDate, _pwi_today, _pwi_isEmpty defined there

// ── External tab/function refs resolved at runtime ────────────────────────────
var _PWI_EXCEPTION_TAB = "Exception Report"; // fallback if CM_TAB_EXCEPTION undefined

function generateCompliancePWIReport() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PWI_TAB_NAME);
  if (!sheet) {
    try { SpreadsheetApp.getUi().alert("'"+PWI_TAB_NAME+"' tab not found."); }
    catch(_) {}
    return;
  }

  var today      = new Date(); today.setHours(0,0,0,0);
  var scanRange  = _pwi_getScanRange(today);
  var thresholds = pwi_readThresholds();
  var unitMap    = pwi_readUnitMap();

  // Read all date blocks in scan range
  var dateBlocks = _pwi_getDateBlocksInRange(sheet, scanRange.from, scanRange.to);

  // Exception accumulators: keyed by exception type, then AEN
  var ex1 = {}; // Pending > 15 days, no TMS date
  var ex2 = {}; // TMS date filled, no actual work after 10 days
  var ex3 = {}; // Actual date filled, photo missing
  var ex4 = {}; // Actual date filled, TMS date blank > 3 days
  var ex5 = {}; // Data quality: Pending Since date missing

  // Detail lists for Exception Report tab
  var ex1Detail=[], ex2Detail=[], ex3Detail=[], ex4Detail=[], ex5Detail=[];

  var scanned = 0, totalEx = 0;

  // Process each date block
  dateBlocks.forEach(function(block) {
    var blockDate = block.date; // Date object
    var blockCol  = block.col;  // 1-indexed column

    // Read all rows in this block
    var lr = sheet.getLastRow();
    if (lr < PWI_DATA_START) return;

    var colB    = sheet.getRange(PWI_DATA_START, PWI_UNIT_COL, lr-PWI_DATA_START+1, 1).getValues();
    var blockVals = sheet.getRange(PWI_DATA_START, blockCol, lr-PWI_DATA_START+1, PWI_BLOCK_COLS).getValues();

    var curUnit = "", curAEN = "";
    for (var i = 0; i < blockVals.length; i++) {
      var unitVal = String(colB[i][0]||"").trim();
      if (unitVal) {
        curUnit = unitVal;
        curAEN  = (unitMap[curUnit] && unitMap[curUnit].aen) ? unitMap[curUnit].aen : "UNASSIGNED";
      }

      var lineAndLoc = String(blockVals[i][1]||"").trim();
      var defect     = String(blockVals[i][2]||"").trim();
      var pendingDt  = String(blockVals[i][3]||"").trim();
      var tmsDt      = String(blockVals[i][4]||"").trim();
      var actualDt   = String(blockVals[i][5]||"").trim();
      var photo      = String(blockVals[i][6]||"").trim();

      // Skip rows with no meaningful data
      if (!lineAndLoc && !defect) continue;
      scanned++;

      // Parse dates
      var pendingDate = _pwi_parseDate(pendingDt);
      var tmsDate     = _pwi_parseDate(tmsDt);
      var actualDate  = _pwi_parseDate(actualDt);

      // Extract P&T from first line of lineAndLoc
      var lines = lineAndLoc.split("\n");
      var pt    = lines[0] ? lines[0].trim() : "Unknown P&T";

      var label = "["+curUnit+"] P&T: "+pt+" | "+defect.substring(0,50) +
                  " | Block: "+_pwi_fmtDate(blockDate);

      // ── Ex5: Data quality — Pending Since date missing ────────────────────
      if (!pendingDate && (lineAndLoc || defect)) {
        if (!ex5[curAEN]) ex5[curAEN] = {};
        if (!ex5[curAEN][pt]) ex5[curAEN][pt] = 0;
        ex5[curAEN][pt]++;
        ex5Detail.push(label + " | Issue: Pending Since date missing");
        // Flag cell with orange background
        try {
          sheet.getRange(PWI_DATA_START+i, blockCol+3)
            .setBackground(PWI_C.DATA_QUALITY);
        } catch(_) {}
      }

      // ── Ex1: Pending > 15 days, no TMS date ──────────────────────────────
      if (pendingDate && !tmsDate) {
        var pendingDays = Math.floor((today - pendingDate) / 864e5);
        if (pendingDays > thresholds.pendingDays) {
          if (!ex1[curAEN]) ex1[curAEN] = {};
          if (!ex1[curAEN][pt]) ex1[curAEN][pt] = 0;
          ex1[curAEN][pt]++;
          ex1Detail.push(label+" | Pending "+pendingDays+" days, TMS date empty");
        }
      }

      // ── Ex2: TMS date filled, actual work missing after 10 days ──────────
      if (tmsDate && !actualDate) {
        var tmsDays = Math.floor((today - tmsDate) / 864e5);
        if (tmsDays > thresholds.complianceDays) { // day 11+
          if (!ex2[curAEN]) ex2[curAEN] = {};
          if (!ex2[curAEN][pt]) ex2[curAEN][pt] = 0;
          ex2[curAEN][pt]++;
          ex2Detail.push(label+" | TMS: "+tmsDt+", "+tmsDays+" days without actual compliance");
        }
      }

      // ── Ex3: Photo missing (actual date filled, photo blank) ──────────────
      if (actualDate && !photo) {
        if (!ex3[curAEN]) ex3[curAEN] = {};
        if (!ex3[curAEN][pt]) ex3[curAEN][pt] = 0;
        ex3[curAEN][pt]++;
        ex3Detail.push(label+" | Actual compliance date: "+actualDt+", photo missing");
      }

      // ── Ex4: Actual date filled, TMS date blank (> grace days) ───────────
      if (actualDate && !tmsDate) {
        var actualDays = Math.floor((today - actualDate) / 864e5);
        if (actualDays > thresholds.graceDays) {
          if (!ex4[curAEN]) ex4[curAEN] = {};
          if (!ex4[curAEN][pt]) ex4[curAEN][pt] = 0;
          ex4[curAEN][pt]++;
          ex4Detail.push(label+" | Actual date: "+actualDt+", TMS date not entered");
        }
      }
    }
  });

  totalEx = _pwi_countAll([ex1,ex2,ex3,ex4,ex5]);

  // Write to Exception Report tab
  _pwi_writeExceptionReport(ex1Detail, ex2Detail, ex3Detail, ex4Detail, ex5Detail,
                             scanned, totalEx, scanRange, today);

  // Write to cache (cm_writeSummaryToCache from CacheManager.gs if available)
  try {
    var cacheSummary = _pwi_buildExSummary(ex1,ex2,ex3,ex4,ex5);
    if (typeof cm_writeSummaryToCache === 'function') {
      cm_writeSummaryToCache("generateCompliancePWIReport", cacheSummary, PWI_CACHE_STALE_DAYS);
    }
  } catch(ce) { console.log("PWI exception cache error: "+ce.message); }

  try {
    SpreadsheetApp.getUi().alert(
      "Compliance by PWI Exception Report generated.\n\n" +
      "Scan range: "+_pwi_fmtDate(scanRange.from)+" → "+_pwi_fmtDate(scanRange.to)+"\n" +
      "Rows scanned: "+scanned+"\n" +
      "Total exceptions: "+totalEx+"\n\n" +
      "See Exception Report tab for details."
    );
  } catch(_) {}
}

// ── Scan range calculator ─────────────────────────────────────────────────────
function _pwi_getScanRange(today) {
  var day   = today.getDate();
  var month = today.getMonth();
  var year  = today.getFullYear();
  var from, to;
  to = new Date(today); // today is always the upper bound
  if (day <= 7) {
    // Include previous month from its 1st
    var prevMonth = month === 0 ? 11 : month - 1;
    var prevYear  = month === 0 ? year - 1 : year;
    from = new Date(prevYear, prevMonth, 1);
  } else {
    // Only current month from 1st
    from = new Date(year, month, 1);
  }
  from.setHours(0,0,0,0);
  to.setHours(0,0,0,0);
  return { from: from, to: to };
}

// ── Get date blocks in range ──────────────────────────────────────────────────
function _pwi_getDateBlocksInRange(sheet, fromDate, toDate) {
  var lr = sheet.getLastColumn();
  if (lr < PWI_DATA_START_COL) return [];
  var dateRow = sheet.getRange(PWI_DATE_ROW, PWI_DATA_START_COL, 1, lr-PWI_DATA_START_COL+1).getValues()[0];
  var blocks = [];
  for (var i = 0; i < dateRow.length; i++) {
    var dv = String(dateRow[i]||"").trim();
    if (!dv) continue;
    var d = _pwi_parseDate(dv);
    if (!d) continue;
    d.setHours(0,0,0,0);
    if (d >= fromDate && d <= toDate) {
      blocks.push({ date: d, col: PWI_DATA_START_COL + i, dateStr: dv });
    }
  }
  return blocks;
}

// ── Build summary for cache/daily report ─────────────────────────────────────
function _pwi_buildExSummary(ex1, ex2, ex3, ex4, ex5) {
  var exceptions = {};
  function addToEx(exObj, typeLabel) {
    for (var aen in exObj) {
      if (!exceptions[aen]) exceptions[aen] = {};
      var pts = Object.keys(exObj[aen]).length;
      var cnt = Object.values(exObj[aen]).reduce(function(a,b){ return a+b; }, 0);
      exceptions[aen][typeLabel + " ("+cnt+" defects, "+pts+" P&T pts)"] = cnt;
    }
  }
  addToEx(ex1, "Pending TMS entry");
  addToEx(ex2, "Pending physical compliance");
  addToEx(ex3, "Photo missing");
  addToEx(ex4, "Actual done, TMS blank");
  addToEx(ex5, "Data quality");
  var total = _pwi_countAll([ex1,ex2,ex3,ex4,ex5]);
  return { totalExceptions: total, exceptions: exceptions };
}

// ── Write exception report to Exception Report tab ───────────────────────────
function _pwi_writeExceptionReport(ex1D, ex2D, ex3D, ex4D, ex5D, scanned, total, scanRange, today) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var tabName = (typeof CM_TAB_EXCEPTION !== 'undefined') ? CM_TAB_EXCEPTION : _PWI_EXCEPTION_TAB;
  var exSheet = ss.getSheetByName(tabName);
  if (!exSheet) return;

  // Build output rows
  var out = [];
  function row(t,b,c) { out.push([t||"",b||false,c||null]); }
  row("COMPLIANCE BY PWI — EXCEPTION REPORT", true, "#EAD1DC");
  row("Date: "+_pwi_today()+"  |  Howrah Division / Eastern Railway", false, "#EAD1DC");
  row("Scan range: "+_pwi_fmtDate(scanRange.from)+" → "+_pwi_fmtDate(scanRange.to), false, "#EAD1DC");
  row("Rows scanned: "+scanned+"  |  Total exceptions: "+total, false, "#EAD1DC");
  row("");
  function section(title, items) {
    row(title+"  ("+items.length+" entries)", true, "#B6D7A8");
    if (!items.length) row("  No exceptions found", false, "#DCFCE7");
    else items.forEach(function(e){ row("  * "+e, false, "#FEE2E2"); });
    row("");
  }
  section("EXCEPTION 1 — PENDING > 15 DAYS, TMS DATE NOT ENTERED", ex1D);
  section("EXCEPTION 2 — TMS DATE FILLED BUT ACTUAL WORK NOT DONE (> 10 DAYS)", ex2D);
  section("EXCEPTION 3 — ACTUAL DATE FILLED BUT PHOTO MISSING", ex3D);
  section("EXCEPTION 4 — ACTUAL WORK DONE BUT NOT ENTERED IN TMS (> 3 DAYS)", ex4D);
  section("EXCEPTION 5 — DATA QUALITY: PENDING SINCE DATE MISSING", ex5D);
  row("END OF COMPLIANCE BY PWI REPORT — "+_pwi_today(), true, "#EAD1DC");

  // Write to sheet directly (writeReportSection may not be available)
  var sectionHeader = "##START:"+PWI_EX_SECTION_ID+"##";
  var sectionEnd    = "##END:"+PWI_EX_SECTION_ID+"##";
  var lr = exSheet.getLastRow();
  var colA = lr > 0 ? exSheet.getRange(1,1,lr,1).getValues() : [];
  var startRow = -1, endRow = -1;
  for (var i=0; i<colA.length; i++) {
    if (String(colA[i][0]).indexOf(sectionHeader)>-1) startRow = i+1;
    if (startRow>0 && String(colA[i][0]).indexOf(sectionEnd)>-1) { endRow=i+1; break; }
  }
  if (startRow>0 && endRow>0 && endRow>startRow) {
    exSheet.deleteRows(startRow, endRow-startRow+1);
  }
  var insertRow = exSheet.getLastRow()+1;
  exSheet.getRange(insertRow,1).setValue(sectionHeader);
  insertRow++;
  out.forEach(function(r) {
    var cell = exSheet.getRange(insertRow,1);
    cell.setValue(r[0]);
    if (r[2]) cell.setBackground(r[2]);
    if (r[1]) cell.setFontWeight("bold");
    insertRow++;
  });
  exSheet.getRange(insertRow,1).setValue(sectionEnd);
}

// ── Count total exceptions ────────────────────────────────────────────────────
function _pwi_countAll(exArr) {
  var total = 0;
  exArr.forEach(function(ex) {
    for (var aen in ex) {
      for (var pt in ex[aen]) total += ex[aen][pt];
    }
  });
  return total;
}

// ── Scheduler entry (triggered when exception report button clicked) ──────────
function s1_CompliancePWI() {
  generateCompliancePWIReport();
}