// ============================================================
//  GuardRailReport.gs — Sheet 2
//  ROB, FOB, Bridge Guard Rail reports with cache writing.
//
//  VERIFIED LOGIC (tested against CSV 27.04.2026):
//  ROB:    Skip Not-Required=15, Skip NIL=9  | Ex1=37, Ex2=23, Ex3=8
//  FOB:    Skip NA=27, Skip Fixed=17         | Ex1=1,  Ex2=0,  Ex3=0
//  Bridge: Skip NA=25, Skip Already-Fixed=15 | Ex1=0,  Ex2=0,  Ex3=0
//
//  Sheet structure (header at row 2, index 1):
//  Col 0=SL | 1=DEN | 2=AEN | 3=PWI | 4=NAME | 5=SECTION | 6=KM
//      7=POSITION | 8=TDC | 9=REVISED TDC | 10=REMARKS | 11=PRE PIC | 12=POST PIC
// ============================================================

var ROB_SECTION_ID    = "ROB_GUARD_RAIL";
var FOB_SECTION_ID    = "FOB_GUARD_RAIL";
var BRIDGE_SECTION_ID = "BRIDGE_GUARD_RAIL";

// ── Entry points ──────────────────────────────────────────────────────────────
function generateROBReport() {
  _generateGuardRailReport("ROB", ROB_SECTION_ID, "generateROBReport",
    ["rob guard","rob_guard","road over bridge"]);
}
function s2_ROB() {
  s2RunOne(ROB_SECTION_ID,"ROB_Guard_Rail","generateROBReport",S2_RECIPIENTS);
}

function generateFOBReport() {
  _generateGuardRailReport("FOB", FOB_SECTION_ID, "generateFOBReport",
    ["fob guard","fob_guard","foot over bridge"]);
}
function s2_FOB() {
  s2RunOne(FOB_SECTION_ID,"FOB_Guard_Rail","generateFOBReport",S2_RECIPIENTS);
}

function generateBridgeReport() {
  _generateGuardRailReport("BRIDGE", BRIDGE_SECTION_ID, "generateBridgeReport",
    ["bridge guard","bridge_guard","major bridge"]);
}
function s2_Bridge() {
  s2RunOne(BRIDGE_SECTION_ID,"Bridge_Guard_Rail","generateBridgeReport",S2_RECIPIENTS);
}

// ── Core scanner ──────────────────────────────────────────────────────────────
function _generateGuardRailReport(typeName, sectionId, cacheFnName, sheetNameHints) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = gFindSheet(ss, sheetNameHints);
  if (!dataSheet) {
    try { SpreadsheetApp.getUi().alert(typeName+" Guard Rail sheet not found."); }
    catch(_) { console.log(typeName+" Guard Rail sheet not found."); }
    return;
  }

  var allData     = dataSheet.getDataRange().getValues();
  var allFormulas = dataSheet.getDataRange().getFormulas();

  // Header is always at row 2 (index 1) for all three guard rail sheets
  // Cols: 0=SL,1=DEN,2=AEN,3=PWI,4=NAME,5=SECTION,6=KM,7=POSITION,
  //       8=TDC,9=REVISED TDC,10=REMARKS,11=PRE PIC,12=POST PIC
  var hdrIdx = 1;

  var iAEN=2, iPWI=3, iNAME=4, iSECT=5, iKM=6, iPOS=7;
  var iTDC=8, iREV=9, iREM=10, iPRE=11, iPOST=12;

  var today = new Date(); today.setHours(0,0,0,0);
  var ex1=[], ex2=[], ex3=[];
  var scanned=0, completed=0;
  var skipNA=0, skipNotReq=0, skipFixed=0;
  var curAEN="", curPWI="";

  for (var r = hdrIdx + 1; r < allData.length; r++) {
    var row  = allData[r];
    var rowF = allFormulas[r] || [];

    // Carry forward AEN and PWI
    var vAEN = String(row[iAEN]||"").trim();
    var vPWI = String(row[iPWI]||"").trim();
    if (vAEN) curAEN = vAEN;
    if (vPWI) curPWI = vPWI;

    var name = String(row[iNAME]||"").trim();
    var km   = String(row[iKM]  ||"").trim();
    var pos  = String(row[iPOS] ||"").trim();
    var tdc_r= String(row[iTDC] ||"").trim();
    var rev_r= String(row[iREV] ||"").trim();
    var pre  = String(row[iPRE] ||"").trim();
    var post = String(row[iPOST]||"").trim();
    var preF = String(rowF[iPRE]||"").trim();
    var postF= String(rowF[iPOST]||"").trim();

    // ── SKIP: no structure exists ────────────────────────────────────────────
    if (!name || name.toUpperCase()==="NIL" || name.toUpperCase()==="NA" ||
        gEmpty(name) || name === "----") {
      skipNA++; continue;
    }

    var posL = pos.toLowerCase();
    var tdcU = tdc_r.trim().toUpperCase();

    // ── SKIP: guard rail not required ────────────────────────────────────────
    // Position says "Not Required" AND TDC = "NA" → pillar too far, no GR needed
    if (posL.indexOf("not required") > -1 && tdcU === "NA") {
      skipNotReq++; continue;
    }

    // ── SKIP: fully fixed ────────────────────────────────────────────────────
    // "ALREADY FIXED" or "FIXED" in position (exact match or standalone phrase)
    // NOT "Guard rail already exists" — that still needs photo documentation
    if (posL === "already fixed" || posL === "fixed" ||
        pos.toUpperCase().trim() === "ALREADY FIXED" ||
        pos.toUpperCase().trim() === "FIXED") {
      skipFixed++; continue;
    }
    // FOB specific: position cell contains only "FIXED" or "Fixed"
    if (/^fixed$/i.test(pos.trim())) { skipFixed++; continue; }

    // ── Completion: post picture present ────────────────────────────────────
    var postDone = _grHasPhoto(post, postF, row[iPOST]);
    if (postDone) { completed++; continue; }
    scanned++;

    // ── Photo check (pre-picture) ────────────────────────────────────────────
    var preOk = _grHasPhoto(pre, preF, row[iPRE]);

    // ── TDC / Revised TDC ────────────────────────────────────────────────────
    var tdcIsNA  = (tdcU === "NA" || tdcU === "NIL" || tdcU === "NOT REQUIRED" ||
                    tdcU === "---" || tdcU === "--" || tdcU === "----" || tdcU === "-----");
    var tdcDate  = tdcIsNA ? null : gDate(row[iTDC]);

    // Rev TDC may be multi-line or have prefix text
    var revDate  = null;
    var revParts = rev_r.split("\n");
    for (var p = 0; p < revParts.length; p++) {
      var rp = revParts[p].trim().replace(/(?:RTDC|revised\s*tdc)[-:\s]*/i,"").trim();
      if (gEmpty(rp)) continue;
      var d = gDate(rp);
      if (d && (!revDate || d > revDate)) revDate = d;
    }

    var lbl = "AEN: "+curAEN+" | PWI: "+curPWI+" | KM: "+km+" | "+name;

    if (!preOk)   ex1.push(lbl);
    if (tdcDate && tdcDate < today && !revDate)
      ex2.push({ label: lbl, tdc: gFmt(tdcDate), days: gDays(tdcDate, today) });
    if (revDate && revDate < today)
      ex3.push({ label: lbl, tdc: gFmt(revDate), days: gDays(revDate, today) });
  }

  s2WriteSection(sectionId, _buildGROutput(typeName, scanned, completed, ex1, ex2, ex3));

  try {
    var summary = s2_buildSummary({
      "Photo missing":       ex1,
      "TDC lapsed":         ex2.map(function(e){ return e.label; }),
      "Revised TDC lapsed": ex3.map(function(e){ return e.label; })
    });
    s2_writeSummaryToSheet1Cache(cacheFnName, summary);
  } catch(e) { console.log(typeName+" cache error: "+e.message); }

  var tot = ex1.length+ex2.length+ex3.length;
  try { SpreadsheetApp.getUi().alert(typeName+" Guard Rail Report updated.\n\n"+
    "Skip (NA/NIL): "+skipNA+"  Skip (Not Required): "+skipNotReq+
    "  Skip (Fixed): "+skipFixed+"\nScanned: "+scanned+"  Completed: "+completed+
    "\nExceptions: "+tot); }
  catch(_) { console.log(typeName+" Guard Rail: "+tot+" exceptions."); }
}

function _grHasPhoto(val, formula, rawVal) {
  // CellImage object
  if (rawVal !== null && rawVal !== undefined &&
      typeof rawVal==="object" && !Array.isArray(rawVal)) return true;
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

function _buildGROutput(typeName, scanned, completed, ex1, ex2, ex3) {
  var tot=ex1.length+ex2.length+ex3.length, out=[];
  var today=gToday();
  function row(t,b,c){ out.push([t||"",b||false,c||null]); }
  row(typeName+" GUARD RAIL — EXCEPTION REPORT",true,S2_C.TITLE_BG);
  row("Date: "+today+"   |   Howrah Division / Eastern Railway",false,S2_C.TITLE_BG);
  row("Scanned: "+scanned+"   |   Completed (excluded): "+completed+
      "   |   Total exceptions: "+tot,false,S2_C.TITLE_BG);
  row("");
  function sect(title,items,fn){
    row(title+"  ("+items.length+" entries)",true,S2_C.HEADER_BG);
    if (!items.length) row("  No exceptions found",false,S2_C.OK_BG);
    else items.forEach(fn); row("");
  }
  sect("EXCEPTION 1 — PHOTO MISSING",ex1,
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
  row("END OF "+typeName+" GUARD RAIL REPORT — "+today,true,S2_C.TITLE_BG);
  return out;
}