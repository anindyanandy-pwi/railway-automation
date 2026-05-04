// ============================================================
//  SandHumpReport.gs — Sheet 2
//  Sand Hump exception report with cache writing.
//
//  VERIFIED LOGIC (tested against CSV 27.04.2026):
//  Skip NIL=10 | Completed=4 | Ex1 Photo=37 | Ex2 TDC=9 | Ex3 RevTDC=1
//
//  Sheet structure:
//  Row 3 (index 2): TRUE header — Sr DEN | AEN | PWI | YARD | LINE ATTACHED | ...
//  Rows 4,5 (index 3,4): Sub-headers (measurement labels) — SKIP
//  Row 6+: Data rows
//  Col 0=Sr DEN | 1=AEN | 2=PWI | 3=YARD | 4=LINE | 8=Current PHOTO | 9=TDC
// ============================================================

var SH_SECTION_ID = "SAND_HUMP";

var SH_SUB_HEADER_WORDS = [
  "standard measurement","length and gradient","length before",
  "length of ramp","ramp (slope"
];

function generateSandHumpReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = gFindSheet(ss, ["sand hump","sandhump","sand-hump"]);
  if (!dataSheet) {
    try { SpreadsheetApp.getUi().alert('Sand Hump sheet not found.'); }
    catch(_) { console.log('Sand Hump sheet not found.'); }
    return;
  }

  var allData     = dataSheet.getDataRange().getValues();
  var allFormulas = dataSheet.getDataRange().getFormulas();

  // ── Header detection ────────────────────────────────────────────────────────
  // Find the TRUE header: row containing BOTH "aen" in col1 AND "yard" in col3
  // This avoids picking up the "Sr DEN/HQ", "Sr DEN/1" group-label rows
  var hdrIdx = -1;
  for (var i = 0; i < allData.length; i++) {
    var c1 = String(allData[i][1]||"").trim().toLowerCase();
    var c3 = String(allData[i][3]||"").trim().toLowerCase();
    if (c1 === "aen" && c3 === "yard") { hdrIdx = i; break; }
  }
  if (hdrIdx < 0) {
    try { SpreadsheetApp.getUi().alert("Header row not found in Sand Hump sheet."); }
    catch(_) {} return;
  }

  // Fixed column positions (verified from actual sheet)
  var iAEN  = 1, iPWI = 2, iYARD = 3, iLINE = 4;
  var iPHOTO = 8, iTDC = 9;

  // Data starts 3 rows after the true header (2 sub-header rows between)
  var dataStart = hdrIdx + 3;
  // Find first real data row (has "Sr DEN/" in col0)
  for (var i = hdrIdx + 1; i < Math.min(hdrIdx + 10, allData.length); i++) {
    var v0 = String(allData[i][0]||"").trim().toLowerCase();
    if (v0.indexOf("sr den/") > -1) { dataStart = i; break; }
  }

  var today = new Date(); today.setHours(0,0,0,0);
  var ex1=[], ex2=[], ex3=[];
  var scanned=0, completed=0, skipNil=0;
  var curAEN="", curPWI="";

  for (var r = dataStart; r < allData.length; r++) {
    var row  = allData[r];
    var rowF = allFormulas[r] || [];

    // Skip sub-header rows that appear between data groups
    var anySubHeader = false;
    for (var s=0; s<SH_SUB_HEADER_WORDS.length; s++) {
      for (var c=0; c<6; c++) {
        if (String(row[c]||"").toLowerCase().indexOf(SH_SUB_HEADER_WORDS[s]) > -1) {
          anySubHeader = true; break;
        }
      }
      if (anySubHeader) break;
    }
    if (anySubHeader) continue;

    // Carry forward AEN and PWI — skip "Sr DEN" labels
    var v1 = String(row[iAEN]||"").trim();
    var v2 = String(row[iPWI]||"").trim();
    if (v1 && v1.toLowerCase().indexOf("sr den") < 0) curAEN = v1;
    if (v2) curPWI = v2;

    var yard  = String(row[iYARD]||"").trim();
    var line  = String(row[iLINE]||"").trim();

    // Skip NIL rows (no sand hump exists)
    var yardU = yard.toUpperCase().trim();
    var lineU = line.toUpperCase().trim();
    if ((yardU === "NIL" || yardU === "NIL ") && (lineU === "NIL" || lineU === "NIL ")) {
      skipNil++; continue;
    }
    if (!yard && !line) continue;
    if (!curAEN) continue;

    var lbl = "AEN: "+curAEN+" | PWI: "+curPWI+" | Yard: "+yard;
    if (line && line !== "----" && line.toUpperCase() !== "NIL") lbl += " | Line: "+line;

    // Photo check (col 8 = current photo of sand hump)
    var photoVal = row[iPHOTO];
    var photoF   = String(rowF[iPHOTO]||"").trim();
    var photoOk  = _shHasPhoto(photoVal, photoF);

    if (photoOk) { completed++; continue; }
    scanned++;

    // Ex1: Photo missing
    ex1.push(lbl);

    // Parse TDC — cell may contain "25/03/2026\nRTDC-10.04.2026"
    var tdc_r   = String(row[iTDC]||"").trim();
    var tdcDate = null, revDate = null;
    var parts   = tdc_r.split("\n");
    for (var p = 0; p < parts.length; p++) {
      var ps = parts[p].trim();
      if (/(?:rtdc|revised)/i.test(ps)) {
        var d = gDate(ps.replace(/(?:RTDC|revised\s*tdc)[-:\s]*/i,"").trim());
        if (d && (!revDate || d > revDate)) revDate = d;
      } else {
        var d2 = gDate(ps);
        if (d2 && !tdcDate) tdcDate = d2;
      }
    }

    // Ex2: TDC lapsed (no revised TDC, photo still missing)
    if (tdcDate && tdcDate < today && !revDate)
      ex2.push({ label: lbl, tdc: gFmt(tdcDate), days: gDays(tdcDate, today) });

    // Ex3: Revised TDC lapsed
    if (revDate && revDate < today)
      ex3.push({ label: lbl, tdc: gFmt(revDate), days: gDays(revDate, today) });
  }

  s2WriteSection(SH_SECTION_ID, _buildSHOutput(scanned, completed, ex1, ex2, ex3));

  try {
    var summary = s2_buildSummary({
      "Current sand hump photo missing": ex1,
      "TDC lapsed (correction pending)": ex2.map(function(e){ return e.label; }),
      "Revised TDC lapsed":              ex3.map(function(e){ return e.label; })
    });
    s2_writeSummaryToSheet1Cache("generateSandHumpReport", summary);
  } catch(e) { console.log("SandHump cache error: "+e.message); }

  var tot = ex1.length+ex2.length+ex3.length;
  try { SpreadsheetApp.getUi().alert("Sand Hump Report updated.\n\nScanned: "+scanned+
    "  Completed (photo present): "+completed+"\nExceptions: "+tot); }
  catch(_) { console.log("Sand Hump: "+tot+" exceptions."); }
}

function _shHasPhoto(val, formula) {
  if (val !== null && val !== undefined && typeof val==="object" && !Array.isArray(val)) return true;
  var v = String(val||"").trim();
  if (/\.(jpg|jpeg|png)/i.test(v)) return true;
  if (v && !gEmpty(v)) {
    var vl = v.toLowerCase();
    if (vl.indexOf("http")>-1 || vl.indexOf("drive.google")>-1) return true;
  }
  if (formula) {
    var fl = String(formula).toLowerCase();
    if (fl.indexOf("image(")>-1 || fl.indexOf("hyperlink(")>-1) return true;
  }
  return false;
}

function s2_SandHump() {
  s2RunOne("SAND_HUMP","Sand_Hump","generateSandHumpReport",S2_RECIPIENTS);
}

function _buildSHOutput(scanned, completed, ex1, ex2, ex3) {
  var tot=ex1.length+ex2.length+ex3.length, out=[];
  var today=gToday();
  function row(t,b,c){ out.push([t||"",b||false,c||null]); }
  row("SAND HUMP — EXCEPTION REPORT",true,S2_C.TITLE_BG);
  row("Date: "+today+"   |   Howrah Division / Eastern Railway",false,S2_C.TITLE_BG);
  row("Scanned: "+scanned+"   |   Completed (photo present, excluded): "+completed+
      "   |   Total exceptions: "+tot,false,S2_C.TITLE_BG);
  row("");
  function sect(title,items,fn){
    row(title+"  ("+items.length+" entries)",true,S2_C.HEADER_BG);
    if (!items.length) row("  No exceptions found",false,S2_C.OK_BG);
    else items.forEach(fn); row("");
  }
  sect("EXCEPTION 1 — CURRENT SAND HUMP PHOTO MISSING",ex1,
    function(e){ row("  * "+e,false,S2_C.EXCEPT_BG); });
  row("EXCEPTION 2 — TDC LAPSED, CORRECTION PENDING  ("+ex2.length+" entries)",true,S2_C.HEADER_BG);
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
  row("END OF SAND HUMP REPORT — "+today,true,S2_C.TITLE_BG);
  return out;
}