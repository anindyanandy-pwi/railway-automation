// ============================================================
//  OMSPeaksReport.gs — Sheet 3
//  OMS/TRC UML Repeated Peaks exception report.
//  FIXES:
//  1. dataStart = hdrIdx+2 (skip sub-header row)
//  2. Strict AEN col match — won't hit "Analysis by AEN"
//  3. CellImage guard on pStr for AEN/PWI (no garbage forward-fill)
//  4. pPhoto now detects CellImage objects (embedded images)
//  5. Row validation: only process rows where col 0 is a positive number
// ============================================================

var OMS_SECTION_ID = "OMS_PEAKS";

// ── Local helpers to fix CellImage issues ─────────────────────────────────────
function _omsCellImage(v) {
  return v !== null && v !== undefined && typeof v === 'object' && !Array.isArray(v);
}
// Use instead of pStr() for AEN/PWI columns — returns "" for embedded images
function _omsStr(row, idx) {
  if (idx < 0 || idx >= row.length) return "";
  if (_omsCellImage(row[idx])) return "";
  return pStr(row, idx);
}
// Use instead of pPhoto() — detects CellImage as a valid photo
function _omsPhoto(v) {
  if (_omsCellImage(v)) return true;
  return pPhoto(v);
}
// Strict column search — only exact keyword match (prevents "Analysis by AEN" matching "aen")
function _omsColStrict(hdr, keyword) {
  var kw = keyword.toLowerCase().trim();
  for (var i = 0; i < hdr.length; i++) {
    if (String(hdr[i]).trim().toLowerCase() === kw) return i;
  }
  return -1;
}

function generateOMSReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = pFindSheet(ss, ["oms","oms peak","oms/trc","repeated oms"]);
  if (!dataSheet) {
    try { SpreadsheetApp.getUi().alert('OMS Peaks sheet not found.'); }
    catch(_) { console.log('OMS Peaks sheet not found.'); }
    return;
  }
  var allData = dataSheet.getDataRange().getValues();

  // ── Find main header row ──────────────────────────────────────────────────
  var hdrIdx = -1;
  for (var i = 0; i < allData.length; i++) {
    for (var j = 0; j < allData[i].length; j++) {
      var v = String(allData[i][j]).toLowerCase();
      if (v.indexOf("km") > -1 || v.indexOf("location") > -1 || v.indexOf("pwi") > -1) {
        hdrIdx = i; break;
      }
    }
    if (hdrIdx > -1) break;
  }
  if (hdrIdx < 0) {
    try { SpreadsheetApp.getUi().alert("Header not found in OMS Peaks sheet."); }
    catch(_) {} return;
  }

  var hdr    = allData[hdrIdx];
  var subHdr = hdrIdx+1 < allData.length ? allData[hdrIdx+1] : [];

  // ── Column detection ──────────────────────────────────────────────────────
  // FIX: Use strict match for AEN — OMS has no dedicated AEN col
  var iAEN    = _omsColStrict(hdr, "aen");   // returns -1 for OMS (correct)
  var iPWI    = pCol(hdr, "pwi");
  var iLine   = pCol(hdr, "line");
  var iLoc    = pCol(hdr, "location");
  // KM may be in sub-header row if main header has merged "Location" cell
  var iKM     = pCol(hdr, "km");
  if (iKM < 0) iKM = pCol(subHdr, "km");
  var iSection = pCol(hdr, "section");
  if (iSection < 0) iSection = pCol(hdr, "major section");
  var iRail    = pCol(hdr, "rail");
  var iPreMes  = pCol(hdr, "pre measurement");
  if (iPreMes < 0) iPreMes = pCol(subHdr, "pre measurement");
  var iPostMes = pCol(hdr, "post measurement");
  if (iPostMes < 0) iPostMes = pCol(subHdr, "post measurement");
  var iPhoto   = pCol(hdr, "photo");
  if (iPhoto < 0) iPhoto = pCol(subHdr, "photo");
  var iTDC     = pCol(hdr, "tdc");
  var iRevTDC  = pCol(hdr, "revised");
  var iDone    = pCol(hdr, "completion");
  if (iDone < 0) iDone = pCol(hdr, "date of tamping");

  var today  = new Date(); today.setHours(0,0,0,0);
  var ex1=[],ex2=[],ex3=[],ex4=[];
  var scanned=0, skipped=0, curAEN="", curPWI="";

  // FIX: dataStart = hdrIdx+2 to skip sub-header row
  var dataStart = hdrIdx + 2;

  for (var r = dataStart; r < allData.length; r++) {
    var row = allData[r];

    // FIX: Skip non-data rows — col 0 (Sl.No.) must be a positive number
    var slVal = row[0];
    if (slVal === null || slVal === undefined || slVal === "") continue;
    if (_omsCellImage(slVal)) continue;
    var slNum = Number(slVal);
    if (isNaN(slNum) || slNum <= 0) continue;

    // FIX: Use _omsStr() — returns "" for CellImage, prevents bad forward-fill
    var ra = _omsStr(row, iAEN);
    var rp = _omsStr(row, iPWI);
    if (iAEN >= 0 && !pEmpty(ra)) curAEN = ra;
    if (!pEmpty(rp)) curPWI = rp;

    var km      = iKM >= 0 ? _omsStr(row, iKM) : "";
    var loc     = iLoc >= 0 ? _omsStr(row, iLoc) : "";
    var section = iSection >= 0 ? _omsStr(row, iSection) : "";
    var line    = _omsStr(row, iLine);

    if (pEmpty(km) && pEmpty(loc) && pEmpty(section)) continue;

    // Build label — only include AEN if we actually have an AEN column
    var label = "";
    if (iAEN >= 0 && !pEmpty(curAEN)) label += "AEN: " + curAEN + " | ";
    if (!pEmpty(curPWI))   label += "PWI: " + curPWI + " | ";
    if (!pEmpty(section))  label += "Sec: " + section + " | ";
    if (!pEmpty(km))       label += "KM: " + km + " | ";
    if (!pEmpty(loc))      label += "Loc: " + loc + " | ";
    if (!pEmpty(line))     label += "Line: " + line + " | ";
    if (iRail >= 0 && !pEmpty(_omsStr(row,iRail))) label += "Rail: " + _omsStr(row,iRail) + " | ";
    label = label.replace(/\s*\|\s*$/, "");  // trim trailing pipe

    var doneStr  = iDone >= 0 ? _omsStr(row, iDone) : "";
    // FIX: _omsPhoto detects CellImage as valid photo
    var photoVal = iPhoto >= 0 ? row[iPhoto] : "";
    if (!pEmpty(doneStr) && _omsPhoto(photoVal)) { skipped++; continue; }
    scanned++;

    // Exception 1: Pre-measurement not done
    if (iPreMes >= 0 && pEmpty(_omsStr(row, iPreMes)) && !_omsCellImage(row[iPreMes]))
      ex1.push(label);

    // Exception 2: Photo missing
    if (!_omsPhoto(photoVal)) ex2.push(label);

    // Exception 3: TDC lapsed (no revised)
    var tdcDate = iTDC >= 0 ? pDate(row[iTDC]) : null;
    var revStr  = iRevTDC >= 0 ? _omsStr(row, iRevTDC) : "";
    if (tdcDate && tdcDate < today && pEmpty(revStr) && pEmpty(doneStr))
      ex3.push({ label:label, tdc:pFmt(tdcDate), days:pDays(tdcDate,today) });

    // Exception 4: Revised TDC lapsed
    var revDate = iRevTDC >= 0 ? pDate(row[iRevTDC]) : null;
    if (revDate && revDate < today && pEmpty(doneStr))
      ex4.push({ label:label, tdc:pFmt(revDate), days:pDays(revDate,today) });
  }

  s3WriteSection(OMS_SECTION_ID, _buildOMSOutput(scanned, skipped, ex1, ex2, ex3, ex4));

  try {
    var summary = s3_buildSummary({
      "Pre measurement overdue": ex1,
      "Photo missing": ex2,
      "TDC lapsed": ex3.map(function(e){ return e.label; }),
      "Revised TDC lapsed": ex4.map(function(e){ return e.label; })
    });
    s3_writeSummaryToSheet1Cache("generateOMSReport", summary);
  } catch(e) { console.log("OMS cache error: " + e.message); }

  var tot = ex1.length+ex2.length+ex3.length+ex4.length;
  try { SpreadsheetApp.getUi().alert("OMS Peaks Report updated.\n\nScanned: "+scanned+
    "  Completed: "+skipped+"\nExceptions: "+tot); }
  catch(_) { console.log("OMS Peaks: "+tot+" exceptions."); }
}

function s3_OMS()         { s3RunOne(OMS_SECTION_ID,"OMS_Peaks","generateOMSReport",S3_RECIPIENTS); }
function s3_OMS_trigger() { s3_OMS(); }

function _buildOMSOutput(scanned, skipped, ex1, ex2, ex3, ex4) {
  var tot=ex1.length+ex2.length+ex3.length+ex4.length, out=[];
  var today=pToday();
  function row(t,b,c){ out.push([t||"",b||false,c||null]); }
  row("OMS/TRC UML REPEATED PEAKS — EXCEPTION REPORT", true, S3_C.TITLE_BG);
  row("Date: "+today+"   |   Howrah Division / Eastern Railway", false, S3_C.TITLE_BG);
  row("Scanned: "+scanned+"   |   Tamping complete (excluded): "+skipped+"   |   Total exceptions: "+tot, false, S3_C.TITLE_BG);
  row("");
  function sect(title, items, fn) {
    row(title+"  ("+items.length+" entries)", true, S3_C.HEADER_BG);
    if (!items.length) row("  No exceptions found", false, S3_C.OK_BG);
    else items.forEach(fn); row("");
  }
  sect("EXCEPTION 1 — PRE MEASUREMENT NOT DONE", ex1, function(e){ row("  * "+e, false, S3_C.EXCEPT_BG); });
  sect("EXCEPTION 2 — PHOTO MISSING", ex2, function(e){ row("  * "+e, false, S3_C.EXCEPT_BG); });
  row("EXCEPTION 3 — TDC LAPSED  ("+ex3.length+" entries)", true, S3_C.HEADER_BG);
  if (!ex3.length) row("  No exceptions found", false, S3_C.OK_BG);
  else ex3.forEach(function(e){
    row("  * "+e.label, false, S3_C.EXCEPT_BG);
    row("      TDC: "+e.tdc+"   |   Lapsed by: "+e.days+" days");
  }); row("");
  row("EXCEPTION 4 — REVISED TDC LAPSED  ("+ex4.length+" entries)", true, S3_C.HEADER_BG);
  if (!ex4.length) row("  No exceptions found", false, S3_C.OK_BG);
  else ex4.forEach(function(e){
    row("  * "+e.label, false, S3_C.EXCEPT_BG);
    row("      Revised TDC: "+e.tdc+"   |   Lapsed by: "+e.days+" days");
  }); row("");
  row("END OF OMS PEAKS REPORT — "+today, true, S3_C.TITLE_BG);
  return out;
}
