// ============================================================
//  TRCUMLReport.gs — Sheet 3
//  FINAL VERSION: Fixed for Merged Cells & Photo Grouping
// ============================================================

var TRC_SECTION_ID = "TRC_UML_PEAKS";

// ── Local Helpers (Do not remove) ─────────────────────────────────────────────

function _trcCellImage(v) {
  // Detects "Insert image in cell" objects
  return v !== null && v !== undefined && typeof v === 'object' && !Array.isArray(v);
}

function _trcStr(row, idx) {
  if (idx < 0 || idx >= row.length) return "";
  if (_trcCellImage(row[idx])) return "";
  return String(row[idx] || "").trim();
}

function _trcPhotoFull(val, formula) {
  if (_trcCellImage(val)) return true;
  if (formula) {
    var f = String(formula).toLowerCase();
    if (f.indexOf("hyperlink(") > -1 || f.indexOf("image(") > -1 || f.indexOf("http") > -1) return true;
  }
  return pPhoto(val);
}

// ── Main Scanner ─────────────────────────────────────────────────────────────

function generateTRCReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = pFindSheet(ss, ["trc","trc uml","trc peak","uml"]);
  if (!dataSheet) return;

  var dataRange   = dataSheet.getDataRange();
  var allData     = dataRange.getValues();
  var allFormulas = dataRange.getFormulas();

  var hdrIdx = -1;
  for (var i = 0; i < allData.length; i++) {
    var v = String(allData[i][0] || "").toLowerCase();
    if (v.indexOf("sl.no") > -1 || v.indexOf("sl. no") > -1) { hdrIdx = i; break; }
  }
  if (hdrIdx < 0) hdrIdx = 1; // Fallback to standard row 2

  var hdr = allData[hdrIdx];
  var subHdr = (hdrIdx + 1 < allData.length) ? allData[hdrIdx + 1] : [];

  var iPWI     = pCol(hdr, "pwi");
  var iSection = pCol(hdr, "section");
  var iLine    = pCol(hdr, "line");
  var iKM      = pCol(subHdr, "from") >= 0 ? pCol(subHdr, "from") : pCol(hdr, "km");
  var iPhoto1  = pCol(subHdr, "photo of inspection");
  var iPhoto2  = pCol(subHdr, "photo of attention");
  var iDone    = pCol(hdr, "completion") >= 0 ? pCol(hdr, "completion") : pCol(hdr, "date of work");
  var iSpeed   = pCol(hdr, "speed");
  var iTDC     = pCol(hdr, "tdc");
  var iRevTDC  = pCol(hdr, "revised");

  var today = new Date(); today.setHours(0,0,0,0);
  var scannedCount = 0, skippedCount = 0;
  
  // Storage for Grouped Exceptions (Key = Location)
  var missingPhotoMap = {}; 
  var speedPendingMap = {};
  var tdcLapsedMap    = {};
  var revisedLapsedMap = {};

  // Persistence for Merged Cells
  var curPWI = "", curSection = "", curLine = "", curKM = "";
  var carryPhotoActive = false, carryPhotoKey = "";

  var dataStart = hdrIdx + 2;

  for (var r = dataStart; r < allData.length; r++) {
    var row = allData[r], rowF = allFormulas[r] || [];
    var sl  = String(row[0] || "").trim();
    if (!sl || isNaN(parseInt(sl))) continue; // Skip non-data rows

    // 1. Metadata Persistence (Handles Merged Cells)
    var rp = _trcStr(row, iPWI);     if (rp) curPWI = rp;
    var rs = _trcStr(row, iSection); if (rs) curSection = rs;
    var rl = _trcStr(row, iLine);    if (rl) curLine = rl;
    var rk = _trcStr(row, iKM);      if (rk) curKM = rk;

    var locLabel = "PWI: " + curPWI + " | Sec: " + curSection + " | KM: " + curKM + " | Line: " + curLine;
    var locKey   = curPWI + curSection + curLine + curKM; // Unique ID for this location

    // 2. Photo Logic (Value + Formula + Carry-Forward)
    var p1Val = iPhoto1 >= 0 ? row[iPhoto1] : "", p1F = iPhoto1 >= 0 ? rowF[iPhoto1] : "";
    var p2Val = iPhoto2 >= 0 ? row[iPhoto2] : "", p2F = iPhoto2 >= 0 ? rowF[iPhoto2] : "";
    
    var rawHasPhoto = _trcPhotoFull(p1Val, p1F) || _trcPhotoFull(p2Val, p2F);
    
    var hasPhoto = rawHasPhoto;
    if (rawHasPhoto) {
      carryPhotoActive = true;
      carryPhotoKey    = locKey;
    } else if (carryPhotoActive && locKey === carryPhotoKey) {
      hasPhoto = true; 
    } else {
      carryPhotoActive = false;
      carryPhotoKey    = "";
    }

    var doneStr = iDone >= 0 ? _trcStr(row, iDone) : "";
    if (hasPhoto && doneStr) { skippedCount++; continue; }
    scannedCount++;

    // 3. Populate Exceptions (One per Location Key)
    if (!hasPhoto) missingPhotoMap[locKey] = locLabel;

    if (iSpeed >= 0 && !pEmpty(_trcStr(row, iSpeed)) && !doneStr) 
      speedPendingMap[locKey] = { label: locLabel, speed: _trcStr(row, iSpeed) };

    var tdcDate = iTDC >= 0 ? pDate(row[iTDC]) : null;
    var revStr  = iRevTDC >= 0 ? _trcStr(row, iRevTDC) : "";
    if (tdcDate && tdcDate < today && !revStr && !doneStr)
      tdcLapsedMap[locKey] = { label: locLabel, tdc: pFmt(tdcDate), days: pDays(tdcDate, today) };

    var revDate = iRevTDC >= 0 ? pDate(row[iRevTDC]) : null;
    if (revDate && revDate < today && !doneStr)
      revisedLapsedMap[locKey] = { label: locLabel, tdc: pFmt(revDate), days: pDays(revDate, today) };
  }

  // Convert Maps to Arrays for Output
  var ex1 = Object.values(speedPendingMap);
  var ex2 = Object.values(missingPhotoMap);
  var ex3 = Object.values(tdcLapsedMap);
  var ex4 = Object.values(revisedLapsedMap);

  s3WriteSection(TRC_SECTION_ID, _buildTRCOutput(scannedCount, skippedCount, ex1, ex2, ex3, ex4));

  try {
    var summary = s3_buildSummary({
      "Speed restriction pending": ex1.map(function(e){ return e.label; }),
      "Photo missing": ex2,
      "TDC lapsed": ex3.map(function(e){ return e.label; }),
      "Revised TDC lapsed": ex4.map(function(e){ return e.label; })
    });
    s3_writeSummaryToSheet1Cache("generateTRCReport", summary);
  } catch(e) {}

  try { SpreadsheetApp.getUi().alert("TRC UML Report updated.\n\nScanned: " + scannedCount + " | Unique Exceptions: " + (ex1.length + ex2.length + ex3.length + ex4.length)); } catch(_) {}
}

function _buildTRCOutput(scanned, skipped, ex1, ex2, ex3, ex4) {
  var tot = ex1.length + ex2.length + ex3.length + ex4.length, out = [];
  var today = pToday();
  function row(t, b, c) { out.push([t || "", b || false, c || null]); }
  row("TRC UML PEAKS — EXCEPTION REPORT", true, S3_C.TITLE_BG);
  row("Date: " + today + "   |   Howrah Division / Eastern Railway", false, S3_C.TITLE_BG);
  row("Scanned Rows: " + scanned + "   |   Completed: " + skipped + "   |   Unique Exceptions: " + tot, false, S3_C.TITLE_BG);
  row("");
  row("EXCEPTION 1 — SPEED RESTRICTION PENDING  (" + ex1.length + " locations)", true, S3_C.HEADER_BG);
  if (!ex1.length) row("  No exceptions found", false, S3_C.OK_BG);
  else ex1.forEach(function(e) { row("  * " + e.label, false, S3_C.EXCEPT_BG); row("      Speed Restriction: " + e.speed); });
  row("");
  row("EXCEPTION 2 — PHOTO MISSING (" + ex2.length + " locations)", true, S3_C.HEADER_BG);
  if (!ex2.length) row("  No exceptions found", false, S3_C.OK_BG);
  else ex2.forEach(function(e) { row("  * " + e, false, S3_C.EXCEPT_BG); });
  row("");
  row("EXCEPTION 3 — TDC LAPSED  (" + ex3.length + " locations)", true, S3_C.HEADER_BG);
  if (!ex3.length) row("  No exceptions found", false, S3_C.OK_BG);
  else ex3.forEach(function(e) { row("  * " + e.label, false, S3_C.EXCEPT_BG); row("      TDC: " + e.tdc + "   |   Lapsed: " + e.days + " days"); });
  row("");
  row("EXCEPTION 4 — REVISED TDC LAPSED  (" + ex4.length + " locations)", true, S3_C.HEADER_BG);
  if (!ex4.length) row("  No exceptions found", false, S3_C.OK_BG);
  else ex4.forEach(function(e) { row("  * " + e.label, false, S3_C.EXCEPT_BG); row("      Revised TDC: " + e.tdc + "   |   Lapsed: " + e.days + " days"); });
  row("");
  row("END OF TRC UML PEAKS REPORT — " + today, true, S3_C.TITLE_BG);
  return out;
}