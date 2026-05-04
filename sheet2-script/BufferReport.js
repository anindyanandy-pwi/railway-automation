// ============================================================
//  BufferReport.gs — Sheet 2
//  Buffer exception report with cache writing.
//
//  VERIFIED LOGIC (tested against CSV 27.04.2026):
//  Completed=17 | Ex1 Photo=215 | Ex2 TDC=2 | Ex3 RevTDC=0
//
//  Sheet structure (header at index 2):
//  Col 0=SL | 1=DEN | 2=AEN | 3=PWI | 4=LINE | 5=PRE PICTURE
//      6=REMARKS | 7=TDC | 8=REVISED TDC | 9=POST PICTURE | 10=ACTION PLAN
// ============================================================

var BUF_SECTION_ID = "BUFFER";

var BUF_COMPLETION_WORDS = [
  "complet","already constructed","line not exist",
  "already done","construction not required"
];

function generateBufferReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = gFindSheet(ss, ["buffer"]);
  if (!dataSheet) {
    try { SpreadsheetApp.getUi().alert('Buffer sheet not found.'); }
    catch(_) { console.log('Buffer sheet not found.'); }
    return;
  }

  var allData     = dataSheet.getDataRange().getValues();
  var allFormulas = dataSheet.getDataRange().getFormulas();

  // Find header: row with col2="AEN" and col3="PWI" and col4="LINE"
  var hdrIdx = -1;
  for (var i = 0; i < allData.length; i++) {
    var c2 = String(allData[i][2]||"").trim().toLowerCase();
    var c3 = String(allData[i][3]||"").trim().toLowerCase();
    var c4 = String(allData[i][4]||"").trim().toLowerCase();
    if (c2 === "aen" && c3 === "pwi" && c4 === "line") { hdrIdx = i; break; }
  }
  if (hdrIdx < 0) {
    try { SpreadsheetApp.getUi().alert("Header row not found in Buffer sheet."); }
    catch(_) {} return;
  }

  // Fixed column positions
  var iAEN=2, iPWI=3, iLINE=4, iPRE=5, iREM=6, iTDC=7, iREV=8, iPOST=9, iACTION=10;

  var today = new Date(); today.setHours(0,0,0,0);
  var ex1=[], ex2=[], ex3=[];
  var scanned=0, completed=0;
  var curAEN="", curPWI="";

  for (var r = hdrIdx + 1; r < allData.length; r++) {
    var row = allData[r];
    var rowF = allFormulas[r] || [];
    if (row[iAEN] && String(row[iAEN]).trim()) curAEN = String(row[iAEN]).trim();
    if (row[iPWI] && String(row[iPWI]).trim()) curPWI = String(row[iPWI]).trim();
    var line = String(row[iLINE]||"").trim();
    if (!line || gEmpty(line) || line.toUpperCase()==="NIL" || line.toUpperCase()==="NA") continue;

    var pre    = String(row[iPRE]   ||"").trim();
    var rem    = String(row[iREM]   ||"").trim();
    var tdc_r  = String(row[iTDC]   ||"").trim();
    var rev_r  = String(row[iREV]   ||"").trim();
    var post   = String(row[iPOST]  ||"").trim();
    var action = String(row[iACTION]||"").trim();
    var preF   = String(rowF[iPRE]  ||"").trim();
    var postF  = String(rowF[iPOST] ||"").trim();
    var lbl    = "AEN: "+curAEN+" | PWI: "+curPWI+" | Line: "+line;

    // Completion check
    var postDone = _bufHasPhoto(post, postF) ||
      _bufAnyWord(action.toLowerCase(), BUF_COMPLETION_WORDS) ||
      _bufAnyWord(rem.toLowerCase(),   ["already constructed","line not exist"]);
    if (postDone) { completed++; continue; }
    scanned++;

    var preOk = _bufHasPhoto(pre, preF);
    var tdcDate = gDate(row[iTDC]);
    var revDate = _bufLatestDate(rev_r);

    if (!preOk)   ex1.push(lbl);
    if (tdcDate && tdcDate < today && !revDate)
      ex2.push({ label: lbl, tdc: gFmt(tdcDate), days: gDays(tdcDate,today) });
    if (revDate && revDate < today)
      ex3.push({ label: lbl, tdc: gFmt(revDate), days: gDays(revDate,today) });
  }

  s2WriteSection(BUF_SECTION_ID, _buildBufOutput(scanned, completed, ex1, ex2, ex3));

  try {
    var summary = s2_buildSummary({
      "Photo missing":       ex1,
      "TDC lapsed":         ex2.map(function(e){ return e.label; }),
      "Revised TDC lapsed": ex3.map(function(e){ return e.label; })
    });
    s2_writeSummaryToSheet1Cache("generateBufferReport", summary);
  } catch(e) { console.log("Buffer cache error: "+e.message); }

  var tot = ex1.length+ex2.length+ex3.length;
  try { SpreadsheetApp.getUi().alert("Buffer Report updated.\n\nScanned: "+scanned+
    "  Completed: "+completed+"\nExceptions: "+tot); }
  catch(_) { console.log("Buffer: "+tot+" exceptions."); }
}

function _bufHasPhoto(val, formula) {
  if (val !== null && val !== undefined && typeof val==="object" && !Array.isArray(val)) return true;
  var v = String(val||"").trim();
  if (/\.(jpg|jpeg|png)/i.test(v)) return true;
  if (v && v!=="--" && v!=="---" && v!=="----" && !gEmpty(v)) {
    var vl = v.toLowerCase();
    if (vl.indexOf("http")>-1 || vl.indexOf("drive.google")>-1) return true;
  }
  if (formula) {
    var fl = String(formula).toLowerCase();
    if (fl.indexOf("image(")>-1 || fl.indexOf("hyperlink(")>-1) return true;
  }
  return false;
}

function _bufAnyWord(text, words) {
  for (var i=0; i<words.length; i++) if (text.indexOf(words[i])>-1) return true;
  return false;
}

function _bufLatestDate(raw) {
  var parts = raw.split("\n"), latest = null;
  for (var i=0; i<parts.length; i++) {
    var s = parts[i].trim().replace(/(?:RTDC|revised\s*tdc)[-:\s]*/i,"").trim();
    var d = gDate(s);
    if (d && (!latest || d > latest)) latest = d;
  }
  return latest;
}

function s2_Buffer() {
  s2RunOne("BUFFER","Buffer","generateBufferReport",S2_RECIPIENTS);
}

function _buildBufOutput(scanned, completed, ex1, ex2, ex3) {
  var tot=ex1.length+ex2.length+ex3.length, out=[];
  var today=gToday();
  function row(t,b,c){ out.push([t||"",b||false,c||null]); }
  row("BUFFER — EXCEPTION REPORT",true,S2_C.TITLE_BG);
  row("Date: "+today+"   |   Howrah Division / Eastern Railway",false,S2_C.TITLE_BG);
  row("Scanned: "+scanned+"   |   Completed (excluded): "+completed+
      "   |   Total exceptions: "+tot,false,S2_C.TITLE_BG);
  row("");
  function sect(title,items,fn){
    row(title+"  ("+items.length+" entries)",true,S2_C.HEADER_BG);
    if (!items.length) row("  No exceptions found",false,S2_C.OK_BG);
    else items.forEach(fn); row("");
  }
  sect("EXCEPTION 1 — PRE-PICTURE (PHOTO) MISSING",ex1,
    function(e){ row("  * "+e,false,S2_C.EXCEPT_BG); });
  row("EXCEPTION 2 — TDC LAPSED  ("+ex2.length+" entries)",true,S2_C.HEADER_BG);
  if (!ex2.length) row("  No exceptions found",false,S2_C.OK_BG);
  else ex2.forEach(function(e){
    row("  * "+e.label,false,S2_C.EXCEPT_BG);
    row("      TDC: "+e.tdc+"   |   Lapsed by: "+e.days+" days");
  }); row("");
  row("EXCEPTION 3 — REVISED TDC LAPSED  ("+ex3.length+" entries)",true,S2_C.HEADER_BG);
  if (!ex3.length) row("  No exceptions found",false,S2_C.OK_BG);
  else ex3.forEach(function(e){
    row("  * "+e.label,false,S2_C.EXCEPT_BG);
    row("      Revised TDC: "+e.tdc+"   |   Lapsed by: "+e.days+" days");
  }); row("");
  row("END OF BUFFER REPORT — "+today,true,S2_C.TITLE_BG);
  return out;
}