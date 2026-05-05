// ============================================================
//  TRCUMLReport.gs — Sheet 3
//  FIXED: Forward-fills merged cells and corrects photo count.
// ============================================================

var TRC_SECTION_ID = "TRC_UML_PEAKS";

function _trcCellImage(v) {
  return v !== null && v !== undefined && typeof v === 'object' && !Array.isArray(v);
}

function _trcStr(row, idx) {
  if (idx < 0 || idx >= row.length) return "";
  if (_trcCellImage(row[idx])) return "";
  return pStr(row, idx);
}

function _trcPhotoFull(val, formula) {
  if (_trcCellImage(val)) return true;
  if (formula) {
    var f = String(formula).toLowerCase();
    if (f.indexOf("hyperlink(") > -1 || f.indexOf("image(") > -1 || f.indexOf("http") > -1) return true;
  }
  return pPhoto(val);
}

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
  if (!dataSheet) return;

  var dataRange  = dataSheet.getDataRange();
  var allData    = dataRange.getValues();
  var allFormulas = dataRange.getFormulas();

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
  if (hdrIdx < 0) return;

  var hdr    = allData[hdrIdx];
  var subHdr = hdrIdx+1 < allData.length ? allData[hdrIdx+1] : [];

  var iAEN     = _trcColStrict(hdr, "aen");
  var iPWI     = pCol(hdr, "pwi");
  var iSection = pCol(hdr, "section");
  var iLine    = pCol(hdr, "line");
  var iRail    = pCol(hdr, "rail");
  var iDefect  = pCol(hdr, "defect");
  var iSpeed   = pCol(hdr, "speed");

  var iKM = pCol(hdr, "km");
  if (iKM < 0) iKM = pCol(subHdr, "from");
  if (iKM < 0) iKM = pCol(subHdr, "km");

  var iPhoto1 = pCol(hdr, "photo");
  if (iPhoto1 < 0) iPhoto1 = pCol(subHdr, "photo of inspection");
  var iPhoto2 = pCol(subHdr, "photo of attention");

  var iTDC    = pCol(hdr, "tdc");
  var iRevTDC = pCol(hdr, "revised");
  var iDone   = pCol(hdr, "completion");
  if (iDone < 0) iDone = pCol(hdr, "date of work");

  var today  = new Date(); today.setHours(0,0,0,0);
  var ex1=[],ex2=[],ex3=[],ex4=[];
  var scanned=0, skipped=0;
  
  // Persistence states for forward-filling merged cells
  var curAEN="", curPWI="", curSection="", curLine="", curKM="";
  var carryPhotoKey = "", carryPhotoActive = false;

  var dataStart = hdrIdx + 2;

  for (var r = dataStart; r < allData.length; r++) {
    var row  = allData[r];
    var rowF = allFormulas[r] || [];

    // Row validation: Sl.No must be present
    if (pEmpty(row[0])) continue;

    // Forward-fill location metadata (Handles merged cells)
    var ra = _trcStr(row, iAEN); if (!pEmpty(ra)) curAEN = ra;
    var rp = _trcStr(row, iPWI); if (!pEmpty(rp)) curPWI = rp;
    var rs = _trcStr(row, iSection); if (!pEmpty(rs)) curSection = rs;
    var rl = _trcStr(row, iLine); if (!pEmpty(rl)) curLine = rl;
    var rk = _trcStr(row, iKM); if (!pEmpty(rk)) curKM = rk;

    if (pEmpty(curKM) && pEmpty(curSection)) continue;

    var label = (curAEN ? "AEN: "+curAEN+" | " : "") + "PWI: "+curPWI+" | Sec: "+curSection+" | KM: "+curKM+" | Line: "+curLine;
    var doneStr = iDone >= 0 ? _trcStr(row, iDone) : "";

    // Comprehensive Photo Check
    var p1Val  = iPhoto1 >= 0 ? row[iPhoto1]  : "";
    var p1Form = iPhoto1 >= 0 ? (rowF[iPhoto1] || "") : "";
    var p2Val  = iPhoto2 >= 0 ? row[iPhoto2]  : "";
    var p2Form = iPhoto2 >= 0 ? (rowF[iPhoto2] || "") : "";
    var rawHasPhoto = _trcPhotoFull(p1Val, p1Form) || _trcPhotoFull(p2Val, p2Form);

    // Carry-Forward Logic
    var locationKey = curPWI + "|" + curSection + "|" + curLine + "|" + curKM;
    var hasPhoto = rawHasPhoto;
    if (rawHasPhoto) {
      carryPhotoActive = true;
      carryPhotoKey    = locationKey;
    } else if (carryPhotoActive && locationKey === carryPhotoKey) {
      hasPhoto = true; 
    } else {
      carryPhotoActive = false;
      carryPhotoKey    = "";
    }

    if (!pEmpty(doneStr) && hasPhoto) { skipped++; continue; }
    scanned++;

    if (iSpeed >= 0 && !pEmpty(_trcStr(row,iSpeed)) && pEmpty(doneStr))
      ex1.push({ label:label, speed:_trcStr(row,iSpeed) });

    if (!hasPhoto) ex2.push(label);

    var tdcDate = iTDC >= 0 ? pDate(row[iTDC]) : null;
    var revStr  = iRevTDC >= 0 ? _trcStr(row, iRevTDC) : "";
    if (tdcDate && tdcDate < today && pEmpty(revStr) && pEmpty(doneStr))
      ex3.push({ label:label, tdc:pFmt(tdcDate), days:pDays(tdcDate,today) });

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
  } catch(e) {}

  var tot = ex1.length+ex2.length+ex3.length+ex4.length;
  try { SpreadsheetApp.getUi().alert("TRC UML Report updated.\n\nScanned: "+scanned+" | Exceptions: "+tot); } catch(_) {}
}