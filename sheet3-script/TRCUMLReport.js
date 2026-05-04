// ============================================================
//  TRCUMLReport.gs — Sheet 3
//  TRC UML Peaks exception report.
//  Updated with new color scheme and cache writing.
// ============================================================

var TRC_SECTION_ID = "TRC_UML_PEAKS";

function generateTRCReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = pFindSheet(ss, ["trc","trc uml","trc peak","uml"]);
  if (!dataSheet) {
    try { SpreadsheetApp.getUi().alert('TRC UML sheet not found.'); }
    catch(_) { console.log('TRC UML sheet not found.'); }
    return;
  }
  var allData = dataSheet.getDataRange().getValues();
  var hdrIdx  = -1;
  for (var i = 0; i < allData.length; i++) {
    for (var j = 0; j < allData[i].length; j++) {
      var v = String(allData[i][j]).toLowerCase();
      if (v.indexOf("km") > -1 || v.indexOf("chainage") > -1 || v.indexOf("section") > -1) {
        hdrIdx = i; break;
      }
    }
    if (hdrIdx > -1) break;
  }
  if (hdrIdx < 0) {
    try { SpreadsheetApp.getUi().alert("Header not found in TRC UML sheet."); }
    catch(_) {} return;
  }
  var hdr     = allData[hdrIdx];
  var iAEN    = pCol(hdr,"aen"),      iPWI     = pCol(hdr,"pwi");
  var iKM     = pCol(hdr,"km"),       iSection = pCol(hdr,"section");
  var iLine   = pCol(hdr,"line"),     iRail    = pCol(hdr,"rail");
  var iDefect = pCol(hdr,"defect"),   iSpeed   = pCol(hdr,"speed");
  var iPhoto  = pCol(hdr,"photo");
  var iTDC    = pCol(hdr,"tdc"),      iRevTDC  = pCol(hdr,"revised");
  var iDone   = pCol(hdr,"completion");
  if (iDone < 0) iDone = pCol(hdr,"date of work");

  var today   = new Date(); today.setHours(0,0,0,0);
  var ex1=[],ex2=[],ex3=[],ex4=[];
  var scanned=0, skipped=0, curAEN="", curPWI="";

  for (var r = hdrIdx+1; r < allData.length; r++) {
    var row = allData[r];
    var ra = pStr(row,iAEN), rp = pStr(row,iPWI);
    if (!pEmpty(ra)) curAEN = ra;
    if (!pEmpty(rp)) curPWI = rp;
    var km      = iKM >= 0 ? pStr(row,iKM) : "";
    var section = iSection >= 0 ? pStr(row,iSection) : "";
    var line    = pStr(row,iLine);
    if (pEmpty(km) && pEmpty(section)) continue;
    var label = "AEN: " + curAEN;
    if (!pEmpty(curPWI))   label += " | PWI: " + curPWI;
    if (!pEmpty(section))  label += " | Sec: " + section;
    if (!pEmpty(km))       label += " | KM: " + km;
    if (!pEmpty(line))     label += " | Line: " + line;
    if (iRail>=0 && !pEmpty(pStr(row,iRail))) label += " | Rail: " + pStr(row,iRail);
    if (iDefect>=0 && !pEmpty(pStr(row,iDefect))) label += " | Defect: " + pStr(row,iDefect);
    var doneStr  = iDone >= 0 ? pStr(row,iDone) : "";
    var photoVal = iPhoto >= 0 ? row[iPhoto] : "";
    if (!pEmpty(doneStr) && pPhoto(photoVal)) { skipped++; continue; }
    scanned++;
    // Exception 1: Speed restriction not lifted (if speed column has value)
    if (iSpeed >= 0 && !pEmpty(pStr(row,iSpeed)) && pEmpty(doneStr))
      ex1.push({ label:label, speed:pStr(row,iSpeed) });
    // Exception 2: Photo missing
    if (!pPhoto(photoVal)) ex2.push(label);
    // Exception 3: TDC lapsed
    var tdcDate = iTDC >= 0 ? pDate(row[iTDC]) : null;
    var revStr  = iRevTDC >= 0 ? pStr(row,iRevTDC) : "";
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