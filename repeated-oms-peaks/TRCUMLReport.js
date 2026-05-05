// ============================================================
//  TRCUMLReport.gs — Sheet 3
//  ULTRA-ROBUST VERSION: Fixes Photo Count & Repetitive Entries
//  Architecture: Metadata Forward-Fill + Location Grouping
// ============================================================

var TRC_SECTION_ID = "TRC_UML_PEAKS";

// ── Local Helpers for Data Integrity ──────────────────────────────────────────

function _trcIsJunk(v) {
  var s = String(v || "").trim().toLowerCase();
  return s === "" || s === "---" || s === "--" || s === "nan" || s === "nil" || s === "na" || s === "n/a" || s === "-";
}

function _trcHasPhoto(val, formula) {
  // 1. Detect embedded CellImage objects
  if (val !== null && val !== undefined && typeof val === 'object' && !Array.isArray(val)) return true;
  
  // 2. Detect formula-based links (=HYPERLINK or =IMAGE)
  if (formula) {
    var f = String(formula).toLowerCase();
    if (f.indexOf("http") > -1 || f.indexOf("drive") > -1 || f.indexOf("image") > -1 || f.indexOf("hyperlink") > -1) return true;
  }
  
  // 3. Detect plain text URLs or "CellImage" text markers
  var s = String(val || "").trim().toLowerCase();
  if (_trcIsJunk(s)) return false;
  if (s.indexOf("http") > -1 || s.indexOf("drive") > -1 || s.indexOf("cellimage") > -1 || s.length > 10) return true;
  
  return false;
}

// ── Main Scanner Function ─────────────────────────────────────────────────────

function generateTRCReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = pFindSheet(ss, ["trc","trc uml","trc peak","uml"]);
  if (!dataSheet) return;

  var dataRange   = dataSheet.getDataRange();
  var allData     = dataRange.getValues();
  var allFormulas = dataRange.getFormulas();

  // Find Header Row (Search for Sl.No)
  var hdrIdx = -1;
  for (var i = 0; i < allData.length; i++) {
    var v0 = String(allData[i][0] || "").toLowerCase();
    if (v0.indexOf("sl.no") > -1 || v0.indexOf("sl. no") > -1) { hdrIdx = i; break; }
  }
  if (hdrIdx < 0) hdrIdx = 1; // Default to row 2

  var hdr = allData[hdrIdx];
  var subHdr = (hdrIdx + 1 < allData.length) ? allData[hdrIdx + 1] : [];

  // Column Mapping
  var iPWI     = pCol(hdr, "pwi");
  var iSection = pCol(hdr, "section");
  var iLine    = pCol(hdr, "line");
  var iKM      = (pCol(subHdr, "from") >= 0) ? pCol(subHdr, "from") : pCol(hdr, "km");
  
  // Identify all columns containing "photo" text
  var photoCols = [];
  for (var c = 0; c < subHdr.length; c++) {
    if (String(subHdr[c]).toLowerCase().indexOf("photo") > -1) photoCols.push(c);
  }
  
  var iDone    = (pCol(hdr, "completion") >= 0) ? pCol(hdr, "completion") : pCol(hdr, "date of work");
  var iSpeed   = pCol(hdr, "speed");
  var iTDC     = pCol(hdr, "tdc");
  var iRevTDC  = pCol(hdr, "revised");

  // Logic: Group by Location Key to handle "One Photo for Multiple Peaks"
  var groups = {}; 
  var curPWI = "", curSection = "", curLine = "", curKM = "";
  var today = new Date(); today.setHours(0,0,0,0);
  
  for (var r = hdrIdx + 2; r < allData.length; r++) {
    var row = allData[r], rowF = allFormulas[r] || [];
    if (!row[0] || isNaN(parseInt(row[0]))) continue; // Skip non-data rows

    // Forward-fill merged cells
    var vP = String(row[iPWI] || "").trim();     if (vP) curPWI = vP;
    var vS = String(row[iSection] || "").trim(); if (vS) curSection = vS;
    var vL = String(row[iLine] || "").trim();    if (vL) curLine = vL;
    var vK = String(row[iKM] || "").trim();      if (vK) curKM = vK;

    var locKey = curPWI + "|" + curSection + "|" + curLine + "|" + curKM;
    if (!groups[locKey]) {
      groups[locKey] = {
        label: "PWI: " + curPWI + " | Sec: " + curSection + " | KM: " + curKM + " | Line: " + curLine,
        hasPhoto: false, isDone: false, speedPending: null, tdcLapsed: null, revisedLapsed: null, rowCount: 0
      };
    }
    
    var g = groups[locKey];
    g.rowCount++;

    // If ANY row in this location group has a photo, the whole group is valid
    photoCols.forEach(function(colIdx) {
      if (_trcHasPhoto(row[colIdx], rowF[colIdx])) g.hasPhoto = true;
    });

    // Check if work is done (Date of work / Completion)
    var doneVal = iDone >= 0 ? String(row[iDone] || "").trim() : "";
    if (doneVal && !_trcIsJunk(doneVal)) g.isDone = true;

    // Capture speed restrictions
    var speed = iSpeed >= 0 ? String(row[iSpeed] || "").trim() : "";
    if (speed && !_trcIsJunk(speed)) g.speedPending = speed;

    // TDC Logic
    var tdcDate = iTDC >= 0 ? pDate(row[iTDC]) : null;
    var revStr  = iRevTDC >= 0 ? String(row[iRevTDC] || "").trim() : "";
    if (tdcDate && tdcDate < today && _trcIsJunk(revStr)) {
      g.tdcLapsed = { tdc: pFmt(tdcDate), days: pDays(tdcDate, today) };
    }
    
    var revDate = iRevTDC >= 0 ? pDate(row[iRevTDC]) : null;
    if (revDate && revDate < today) {
      g.revisedLapsed = { tdc: pFmt(revDate), days: pDays(revDate, today) };
    }
  }

  // ── Build final exception lists based on GROUP status ───────────────────────
  
  var ex1=[], ex2=[], ex3=[], ex4=[];
  var totalAnalyzedRows = 0;
  var totalExceptions = 0;

  for (var k in groups) {
    var g = groups[k];
    totalAnalyzedRows += g.rowCount;

    // Skip if completely resolved
    if (g.isDone && g.hasPhoto) continue;

    // Report location once if photo is missing
    if (!g.hasPhoto) ex2.push(g.label);

    // Only report Speed/TDC if not marked as Done
    if (!g.isDone) {
      if (g.speedPending) ex1.push({ label: g.label, speed: g.speedPending });
      if (g.tdcLapsed)    ex3.push({ label: g.label, tdc: g.tdcLapsed.tdc, days: g.tdcLapsed.days });
      if (g.revisedLapsed) ex4.push({ label: g.label, tdc: g.revisedLapsed.tdc, days: g.revisedLapsed.days });
    }
  }

  totalExceptions = ex1.length + ex2.length + ex3.length + ex4.length;

  s3WriteSection(TRC_SECTION_ID, _buildTRCOutput(totalAnalyzedRows, totalExceptions, ex1, ex2, ex3, ex4));

  try {
    var summary = s3_buildSummary({
      "Speed restriction pending": ex1.map(function(e){ return e.label; }),
      "Photo missing": ex2,
      "TDC lapsed": ex3.map(function(e){ return e.label; }),
      "Revised TDC lapsed": ex4.map(function(e){ return e.label; })
    });
    s3_writeSummaryToSheet1Cache("generateTRCReport", summary);
  } catch(e) { console.log("TRC Cache Error: " + e.message); }

  try { SpreadsheetApp.getUi().alert("TRC Report updated.\nRows Analyzed: " + totalAnalyzedRows + "\nUnique Locations with Exceptions: " + ex2.length); } catch(_) {}
}

// ── Shared Automation Helpers ────────────────────────────────────────────────

function s3_TRC()         { s3RunOne(TRC_SECTION_ID,"TRC_UML_Peaks","generateTRCReport",S3_RECIPIENTS); }
function s3_TRC_trigger() { s3_TRC(); }

function _buildTRCOutput(analyzed, totalEx, ex1, ex2, ex3, ex4) {
  var out = [];
  var today = pToday();
  function row(t, b, c) { out.push([t || "", b || false, c || null]); }
  
  row("TRC UML PEAKS — EXCEPTION REPORT", true, S3_C.TITLE_BG);
  row("Date: " + today + "   |   Howrah Division / Eastern Railway", false, S3_C.TITLE_BG);
  row("Total Data Rows: " + analyzed + "   |   Exceptions Found: " + totalEx + " (at " + ex2.length + " unique locations)", false, S3_C.TITLE_BG);
  row("");
  
  row("EXCEPTION 1 — SPEED RESTRICTION PENDING  (" + ex1.length + " entries)", true, S3_C.HEADER_BG);
  if (!ex1.length) row("  No exceptions found", false, S3_C.OK_BG);
  else ex1.forEach(function(e) { row("  * " + e.label, false, S3_C.EXCEPT_BG); row("      Speed Restriction: " + e.speed); });
  row("");
  
  row("EXCEPTION 2 — PHOTO MISSING (" + ex2.length + " locations)", true, S3_C.HEADER_BG);
  row("  (One entry per location — covers multiple peak rows)", true, S3_C.HEADER_BG);
  if (!ex2.length) row("  No exceptions found", false, S3_C.OK_BG);
  else ex2.forEach(function(e) { row("  * " + e, false, S3_C.EXCEPT_BG); });
  row("");
  
  row("EXCEPTION 3 — TDC LAPSED  (" + ex3.length + " locations)", true, S3_C.HEADER_BG);
  if (!ex3.length) row("  No exceptions found", false, S3_C.OK_BG);
  else ex3.forEach(function(e) { row("  * " + e.label, false, S3_C.EXCEPT_BG); row("      TDC: " + e.tdc + "   |   Lapsed by: " + e.days + " days"); });
  row("");
  
  row("EXCEPTION 4 — REVISED TDC LAPSED  (" + ex4.length + " locations)", true, S3_C.HEADER_BG);
  if (!ex4.length) row("  No exceptions found", false, S3_C.OK_BG);
  else ex4.forEach(function(e) { row("  * " + e.label, false, S3_C.EXCEPT_BG); row("      Revised TDC: " + e.tdc + "   |   Lapsed by: " + e.days + " days"); });
  row("");
  
  row("END OF TRC UML PEAKS REPORT — " + today, true, S3_C.TITLE_BG);
  return out;
}