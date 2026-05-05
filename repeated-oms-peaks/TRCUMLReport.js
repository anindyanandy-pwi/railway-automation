// ============================================================
//  TRCUMLReport.gs — Sheet 3
//  TRC UML Peaks exception report.
//  FIXES:
//  1. dataStart = hdrIdx+2 (skip sub-header row)
//  2. Strict AEN col match — "Analysis by AEN" won't match
//  3. CellImage guard on pStr for AEN/PWI (no garbage forward-fill)
//  4. Check BOTH photo columns: col11=Photo of inspection, col13=Photo of Attention
//  5. Photo CARRY-FORWARD: merged rows inherit photo from first row of group
//  6. Row validation: col 0 must be positive number
// ============================================================

var TRC_SECTION_ID = "TRC_UML_PEAKS";

// ── Local helpers ─────────────────────────────────────────────────────────────
function _trcCellImage(v) {
  return v !== null && v !== undefined && typeof v === 'object' && !Array.isArray(v);
}
// Returns "" for CellImage objects — prevents "CellImage" being forward-filled as PWI/AEN
function _trcStr(row, idx) {
  if (idx < 0 || idx >= row.length) return "";
  if (_trcCellImage(row[idx])) return "";
  return pStr(row, idx);
}
// Detects CellImage as valid photo
function _trcPhoto(v) {
  if (_trcCellImage(v)) return true;
  return pPhoto(v);
}
// Strict column search — exact keyword match only
function _trcColStrict(hdr, keyword) {
  var kw = keyword.toLowerCase().trim();
  for (var i = 0; i < hdr.length; i++) {
    if (String(hdr[i]).trim().toLowerCase() === kw) return i;
  }
  return -1;
}

function generateTRCReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = pFindSheet(ss, ["trc","trc uml","trc peak","uml"]);
  if (!dataSheet) {
    try { SpreadsheetApp.getUi().alert('TRC UML sheet not found.'); }
    catch(_) { console.log('TRC UML sheet not found.'); }
    return;
  }

  var allData = dataSheet.getDataRange().getValues();

  // ── Find main header row ──────────────────────────────────────────────────
  var hdrIdx = -1;
  for (var i = 0; i < allData.length; i++) {
    for (var j = 0; j < allData[i].length; j++) {
      var v = String(allData[i][j]).toLowerCase();
      if (v.indexOf("km") > -1 || v.indexOf("section") > -1 || v.indexOf("pwi") > -1) {
        hdrIdx = i; break;
      }
    }
    if (hdrIdx > -1) break;
  }
  if (hdrIdx < 0) {
    try { SpreadsheetApp.getUi().alert("Header not found in TRC UML sheet."); }
    catch(_) {} return;
  }

  var hdr    = allData[hdrIdx];
  var subHdr = hdrIdx+1 < allData.length ? allData[hdrIdx+1] : [];

  // ── Column detection ──────────────────────────────────────────────────────
  var iAEN     = _trcColStrict(hdr, "aen");   // -1: no dedicated AEN col in TRC
  var iPWI     = pCol(hdr, "pwi");
  var iSection = pCol(hdr, "section");
  var iLine    = pCol(hdr, "line");
  var iRail    = pCol(hdr, "rail");
  var iDefect  = pCol(hdr, "defect");
  var iSpeed   = pCol(hdr, "speed");

  // KM: check main header then sub-header
  var iKM = pCol(hdr, "km");
  if (iKM < 0) iKM = pCol(subHdr, "from");
  if (iKM < 0) iKM = pCol(subHdr, "km");

  // FIX: TWO photo columns — check both
  var iPhoto1 = pCol(hdr, "photo");
  if (iPhoto1 < 0) iPhoto1 = pCol(subHdr, "photo of inspection");
  if (iPhoto1 < 0) iPhoto1 = pCol(subHdr, "photo");
  // Second photo column: "Photo of Attention"
  var iPhoto2 = pCol(subHdr, "photo of attention");

  var iPreMes = pCol(hdr, "pre measurement");
  if (iPreMes < 0) iPreMes = pCol(subHdr, "pre measurement");
  var iTDC    = pCol(hdr, "tdc");
  var iRevTDC = pCol(hdr, "revised");
  var iDone   = pCol(hdr, "completion");
  if (iDone < 0) iDone = pCol(hdr, "date of work");
  if (iDone < 0) iDone = pCol(subHdr, "post measurement");

  var today  = new Date(); today.setHours(0,0,0,0);
  var ex1=[],ex2=[],ex3=[],ex4=[];
  var scanned=0, skipped=0, curAEN="", curPWI="";

  // FIX: Photo carry-forward state for merged rows
  var carryPhotoKey = "";
  var carryPhotoActive = false;

  // FIX: dataStart = hdrIdx+2 to skip sub-header row
  var dataStart = hdrIdx + 2;

  for (var r = dataStart; r < allData.length; r++) {
    var row = allData[r];

    // FIX: Skip non-data rows — col 0 (Sl.No.) must be a positive number
    var slVal = row[0];
    if (slVal === null || slVal === undefined || slVal === "") continue;
    if (_trcCellImage(slVal)) continue;
    var slNum = Number(slVal);
    if (isNaN(slNum) || slNum <= 0) continue;

    // FIX: _trcStr returns "" for CellImage — no garbage forward-fill
    var ra = _trcStr(row, iAEN);
    var rp = _trcStr(row, iPWI);
    if (iAEN >= 0 && !pEmpty(ra)) curAEN = ra;
    if (!pEmpty(rp)) curPWI = rp;

    var km      = iKM >= 0 ? _trcStr(row, iKM) : "";
    var section = iSection >= 0 ? _trcStr(row, iSection) : "";
    var line    = _trcStr(row, iLine);

    if (pEmpty(km) && pEmpty(section)) continue;

    // Build label
    var label = "";
    if (iAEN >= 0 && !pEmpty(curAEN)) label += "AEN: " + curAEN + " | ";
    if (!pEmpty(curPWI))   label += "PWI: " + curPWI + " | ";
    if (!pEmpty(section))  label += "Sec: " + section + " | ";
    if (!pEmpty(km))       label += "KM: " + km + " | ";
    if (!pEmpty(line))     label += "Line: " + line + " | ";
    if (iRail>=0 && !pEmpty(_trcStr(row,iRail)))   label += "Rail: "   + _trcStr(row,iRail)   + " | ";
    if (iDefect>=0 && !pEmpty(_trcStr(row,iDefect))) label += "Defect: " + _trcStr(row,iDefect) + " | ";
    label = label.replace(/\s*\|\s*$/, "");

    var doneStr = iDone >= 0 ? _trcStr(row, iDone) : "";

    // FIX: Check BOTH photo columns
    var photoVal1 = iPhoto1 >= 0 ? row[iPhoto1] : "";
    var photoVal2 = iPhoto2 >= 0 ? row[iPhoto2] : "";
    var rawHasPhoto = _trcPhoto(photoVal1) || _trcPhoto(photoVal2);

    // FIX: Photo carry-forward for merged rows
    // Group key = same location (PWI + Section + Line + KM)
    var locationKey = curPWI + "|" + section + "|" + line + "|" + km;
    var hasPhoto = rawHasPhoto;
    if (rawHasPhoto) {
      // This row has a photo — activate carry-forward for this group
      carryPhotoActive = true;
      carryPhotoKey = locationKey;
    } else if (carryPhotoActive && locationKey === carryPhotoKey) {
      // Same location group as last photo row — inherit photo (merged cell)
      hasPhoto = true;
    } else {
      // Different location — reset carry-forward
      carryPhotoActive = false;
      carryPhotoKey = "";
    }

    if (!pEmpty(doneStr) && hasPhoto) { skipped++; continue; }
    scanned++;

    // Exception 1: Speed restriction pending
    if (iSpeed >= 0 && !pEmpty(_trcStr(row,iSpeed)) && pEmpty(doneStr))
      ex1.push({ label:label, speed:_trcStr(row,iSpeed) });

    // Exception 2: Photo missing
    if (!hasPhoto) ex2.push(label);

    // Exception 3: TDC lapsed
    var tdcDate = iTDC >= 0 ? pDate(row[iTDC]) : null;
    var revStr  = iRevTDC >= 0 ? _trcStr(row, iRevTDC) : "";
    if (tdcDate && tdcDate < today && pEmpty(revStr) && pEmpty(doneStr))
      ex3.push({ label:label, tdc:pFmt(tdcDate), days:pDays(tdcDate,today) });

    // Exception 4: Revised TDC lapsed
    var revDate = iRevTDC >= 0 ? pDate(row[iRevTDC]) : null;
    if (revDate && revDate < today && pEmpty(doneStr))
      ex4.push({ label:label, tdc:pFmt(revDate), days:pDays(revDate,today) });
  }

  s3WriteSection(TRC_SECTION_ID, _buildTRCOutput(scanned, skipped, ex1, ex2, ex3, ex4));

  try {
    var summary = s3_buildSummary({
      "Speed restriction pending": ex1.map(function(e){ return e.label; }),
      "Photo missing": ex2,
      "TDC lapsed": ex3.map(function(e){ return e.label; }),
      "Revised TDC lapsed": ex4.map(function(e){ return e.label; })
    });
    s3_writeSummaryToSheet1Cache("generateTRCReport", summary);
  } catch(e) { console.log("TRC cache error: " + e.message); }

  var tot = ex1.length+ex2.length+ex3.length+ex4.length;
  try { SpreadsheetApp.getUi().alert("TRC UML Peaks Report updated.\n\nScanned: "+scanned+
    "  Completed: "+skipped+"\nExceptions: "+tot); }
  catch(_) { console.log("TRC UML: "+tot+" exceptions."); }
}

function s3_TRC()         { s3RunOne(TRC_SECTION_ID,"TRC_UML_Peaks","generateTRCReport",S3_RECIPIENTS); }
function s3_TRC_trigger() { s3_TRC(); }

function _buildTRCOutput(scanned, skipped, ex1, ex2, ex3, ex4) {
  var tot=ex1.length+ex2.length+ex3.length+ex4.length, out=[];
  var today=pToday();
  function row(t,b,c){ out.push([t||"",b||false,c||null]); }
  row("TRC UML PEAKS — EXCEPTION REPORT", true, S3_C.TITLE_BG);
  row("Date: "+today+"   |   Howrah Division / Eastern Railway", false, S3_C.TITLE_BG);
  row("Scanned: "+scanned+"   |   Work complete (excluded): "+skipped+"   |   Total exceptions: "+tot, false, S3_C.TITLE_BG);
  row("");
  row("EXCEPTION 1 — SPEED RESTRICTION PENDING  ("+ex1.length+" entries)", true, S3_C.HEADER_BG);
  if (!ex1.length) row("  No exceptions found", false, S3_C.OK_BG);
  else ex1.forEach(function(e){
    row("  * "+e.label, false, S3_C.EXCEPT_BG);
    row("      Speed Restriction: "+e.speed);
  }); row("");
  function sect(title, items, fn) {
    row(title+"  ("+items.length+" entries)", true, S3_C.HEADER_BG);
    if (!items.length) row("  No exceptions found", false, S3_C.OK_BG);
    else items.forEach(fn); row("");
  }
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
  row("END OF TRC UML PEAKS REPORT — "+today, true, S3_C.TITLE_BG);
  return out;
}
