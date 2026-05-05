// ============================================================
//  GaplessReport.gs — Sheet 1 | Gapless Joint of CMS Crossing
//  Updated with new color scheme and cache writing.
//  FIX: Photo detection now handles CellImage objects + formulas
// ============================================================

var GAP_SECTION_ID   = "GAPLESS_CMS";
var GAP_THRESHOLD_MM = 2;
var COMPLETION_WORDS = ["rectif","correct","compli","done","complet"];

function generateGaplessReport() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet=null;
  var sheets=ss.getSheets();
  for (var s=0;s<sheets.length;s++) {
    var n=sheets[s].getName().toLowerCase();
    if (n.indexOf("gapless")>-1||n.indexOf("cms")>-1) { dataSheet=sheets[s]; break; }
  }
  if (!dataSheet) {
    try { SpreadsheetApp.getUi().alert('Gapless/CMS sheet not found.'); }
    catch(_) { console.log('Gapless/CMS sheet not found.'); }
    return;
  }

  // ── FIX: Read both values AND formulas to detect image/hyperlink cells ──
  var dataRange=dataSheet.getDataRange();
  var allData=dataRange.getValues();
  var allFormulas=dataRange.getFormulas();

  var mainHdrIdx=-1;
  for (var i=0;i<allData.length;i++) {
    for (var j=0;j<allData[i].length;j++) {
      if (String(allData[i][j]).toLowerCase().indexOf("point no")>-1) { mainHdrIdx=i; break; }
    }
    if (mainHdrIdx>-1) break;
  }
  if (mainHdrIdx<0) {
    try { SpreadsheetApp.getUi().alert("Header row not found in Gapless CMS sheet."); }
    catch(_) {}
    return;
  }
  var mainHdr=allData[mainHdrIdx];
  var iAEN=_gapColIdx(mainHdr,"aen"), iPWI=_gapColIdx(mainHdr,"pwi");
  var iLine=_gapColIdx(mainHdr,"line"), iPoint=_gapColIdx(mainHdr,"point no");
  var iForeStart=_gapColIdx(mainHdr,"fore"), iBackStart=_gapColIdx(mainHdr,"back");
  var iForeML=iForeStart, iForeTo=iForeStart+1;
  var iBackML=iBackStart, iBackTo=iBackStart+1;
  var iForePhoto=_gapColIdx(mainHdr,"foreleg");
  var iBackPhoto=_gapColIdx(mainHdr,"backleg");
  if (iForePhoto<0) iForePhoto=_gapColIdx(mainHdr,"fore leg");
  if (iBackPhoto<0) iBackPhoto=_gapColIdx(mainHdr,"back leg");
  var subHdr=mainHdrIdx+1<allData.length?allData[mainHdrIdx+1]:[];
  if (iForePhoto<0||iBackPhoto<0) {
    for (var c=0;c<mainHdr.length;c++) {
      var mh=String(mainHdr[c]).toLowerCase(), sh2=String(subHdr[c]||"").toLowerCase();
      if (mh.indexOf("photo")>-1||sh2.indexOf("photo")>-1) {
        if (mh.indexOf("fore")>-1||sh2.indexOf("fore")>-1) iForePhoto=c;
        if (mh.indexOf("back")>-1||sh2.indexOf("back")>-1) iBackPhoto=c;
      }
    }
  }
  var iTDC=_gapColIdx(mainHdr,"tdc");
  var today=new Date(); today.setHours(0,0,0,0);
  var ex1=[],ex2=[],ex3=[], scanned=0,skipped=0;
  var curAEN="",curPWI="";
  var dataStart=mainHdrIdx+2;
  for (var r=dataStart;r<allData.length;r++) {
    var row=allData[r];
    var rowF=allFormulas[r]; // ── FIX: get formula row ──

    var ra=String(row[iAEN]||"").trim(), rp=String(row[iPWI]||"").trim();
    if (!_gapEmpty(ra)) curAEN=ra;
    if (!_gapEmpty(rp)) curPWI=rp;
    var point=String(row[iPoint]||"").trim();
    if (_gapEmpty(point)||point==="--"||point==="---") continue;
    var line=String(row[iLine]||"").trim();
    var label="AEN: "+curAEN+" | PWI: "+curPWI;
    if (!_gapEmpty(line)) label+=" | Line: "+line;
    label+=" | Point: "+point;
    var gForeML=_parseGap(row[iForeML]), gForeTo=_parseGap(row[iForeTo]);
    var gBackML=_parseGap(row[iBackML]), gBackTo=_parseGap(row[iBackTo]);
    var maxGap=Math.max(gForeML!==null?gForeML:0,gForeTo!==null?gForeTo:0,
                        gBackML!==null?gBackML:0,gBackTo!==null?gBackTo:0);
    var gapExceeds=maxGap>GAP_THRESHOLD_MM;
    var tdcRaw=String(row[iTDC]||"").trim(), tdcLow=tdcRaw.toLowerCase();
    var workDone=false;
    if (!gapExceeds) workDone=true;
    else { for (var w=0;w<COMPLETION_WORDS.length;w++) { if (tdcLow.indexOf(COMPLETION_WORDS[w])>-1) { workDone=true; break; } } }
    if (workDone&&gapExceeds) { skipped++; continue; }
    scanned++;
    if (gapExceeds) {
      var badCols=[];
      if (gForeML!==null&&gForeML>GAP_THRESHOLD_MM) badCols.push("Fore-ML: "+gForeML+"mm");
      if (gForeTo!==null&&gForeTo>GAP_THRESHOLD_MM) badCols.push("Fore-TO: "+gForeTo+"mm");
      if (gBackML!==null&&gBackML>GAP_THRESHOLD_MM) badCols.push("Back-ML: "+gBackML+"mm");
      if (gBackTo!==null&&gBackTo>GAP_THRESHOLD_MM) badCols.push("Back-To: "+gBackTo+"mm");
      ex1.push({label:label, cols:badCols.join(", ")});
    }

    // ── FIX: Pass formula to _gapPhoto for each photo cell ──
    var missingPhotos=[];
    var foreFormula = rowF&&iForePhoto>=0 ? rowF[iForePhoto] : "";
    var backFormula = rowF&&iBackPhoto>=0 ? rowF[iBackPhoto] : "";
    if (!_gapPhoto(row[iForePhoto], foreFormula)) missingPhotos.push("Fore");
    if (!_gapPhoto(row[iBackPhoto], backFormula)) missingPhotos.push("Back");
    if (missingPhotos.length>0) ex2.push({label:label, missing:missingPhotos.join(", ")+" photo missing"});

    if (gapExceeds) {
      if (_gapEmpty(tdcRaw)||tdcRaw==="---") {
        ex3.push({label:label, status:"TDC not populated"});
      } else {
        var effDate=_extractLatestGapDate(tdcRaw);
        if (effDate&&effDate<today) {
          var days=Math.floor((today-effDate)/864e5);
          ex3.push({label:label, status:"TDC lapsed by "+days+" days (Effective TDC: "+_gapFmt(effDate)+")"});
        }
      }
    }
  }
  writeReportSection(GAP_SECTION_ID, _buildGapOutput(scanned,skipped,ex1,ex2,ex3));
  try {
    var summary=cm_buildSummary({
      "Gap exceeds 2mm": ex1.map(function(e){return e.label;}),
      "Photo not uploaded": ex2.map(function(e){return e.label;}),
      "TDC missing or lapsed": ex3.map(function(e){return e.label;})
    });
    cm_writeSummaryToCache("generateGaplessReport", summary);
  } catch(e) { console.log("Gap cache error: "+e.message); }
  var tot=ex1.length+ex2.length+ex3.length;
  try { SpreadsheetApp.getUi().alert("Gapless CMS Report updated.\n\nScanned: "+scanned+"  Work complete: "+skipped+"\nTotal exceptions: "+tot); }
  catch(_) { console.log("Gapless CMS: "+tot+" exceptions."); }
}

function _gapColIdx(hdr,kw) {
  var k=kw.toLowerCase();
  for (var i=0;i<hdr.length;i++) if (String(hdr[i]).toLowerCase().indexOf(k)>-1) return i;
  return -1;
}
function _gapEmpty(v) {
  if (v===null||v===undefined) return true;
  return ["","---","--","----","nan","none","nil","na","n/a","-"].indexOf(String(v).trim().toLowerCase())>-1;
}

// ── FIX: Now handles CellImage objects AND formula-based images ──
function _gapPhoto(v, formula) {
  // CellImage: embedded image inserted via Insert > Image in cell
  if (v !== null && v !== undefined && typeof v === 'object' && !Array.isArray(v)) return true;
  if (!v) return false;
  var s=String(v).trim();
  if (_gapEmpty(s)) return false;
  var l=s.toLowerCase();
  // Value-based photo indicators
  if (l.indexOf("hyperlink")>-1||l.indexOf("image(")>-1||l.indexOf("http")>-1||
      l.indexOf("drive.google")>-1||/\.(jpg|jpeg|png|gif)/i.test(s)) return true;
  // Formula-based: =IMAGE() or =HYPERLINK()
  if (formula) {
    var f=String(formula).toLowerCase();
    return f.indexOf("image(")>-1||f.indexOf("hyperlink(")>-1||f.indexOf("http")>-1;
  }
  return false;
}

function _parseGap(val) {
  if (val===null||val===undefined) return null;
  var s=String(val).trim(); if (_gapEmpty(s)) return null;
  s=s.split('\n')[0]; s=s.replace(/^[A-Za-z0-9\/]+\s*[-–]\s*/g,"");
  var m=s.match(/(\d+(?:\.\d+)?)/); return m?parseFloat(m[1]):null;
}
function _extractLatestGapDate(text) {
  var pattern=/(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})/g, match, latest=null;
  while ((match=pattern.exec(text))!==null) {
    var d=parseInt(match[1]),m=parseInt(match[2])-1,y=parseInt(match[3]);
    if (y<100) y+=2000;
    var dt=new Date(y,m,d);
    if (!isNaN(dt.getTime())&&(!latest||dt>latest)) latest=dt;
  }
  return latest;
}
function _gapFmt(d) {
  if (!d) return "";
  return ('0'+d.getDate()).slice(-2)+'.'+('0'+(d.getMonth()+1)).slice(-2)+'.'+d.getFullYear();
}
function _buildGapOutput(scanned,skipped,ex1,ex2,ex3) {
  var tot=ex1.length+ex2.length+ex3.length, out=[];
  var today=_gapFmt(new Date());
  function row(t,b,c){ out.push([t||"",b||false,c||null]); }
  row("GAPLESS JOINT OF CMS CROSSING — EXCEPTION REPORT",true,CM_C.TITLE_BG);
  row("Date: "+today+"   |   Howrah Division / Eastern Railway",false,CM_C.TITLE_BG);
  row("Entries scanned: "+scanned+"   |   Work complete: "+skipped+"   |   Total exceptions: "+tot,false,CM_C.TITLE_BG);
  row("");
  row("EXCEPTION 1 — GAP VALUE EXCEEDS 2mm  ("+ex1.length+" entries)",true,CM_C.HEADER_BG);
  if (!ex1.length) row("  No exceptions found",false,CM_C.OK_BG);
  else ex1.forEach(function(e){row("  * "+e.label,false,CM_C.EXCEPT_BG);row("      Exceeding: "+e.cols);});
  row("");
  row("EXCEPTION 2 — PHOTO NOT UPLOADED  ("+ex2.length+" entries)",true,CM_C.HEADER_BG);
  row("  (Applies to all entries — every crossing must have photos)",false,CM_C.HEADER_BG);
  if (!ex2.length) row("  No exceptions found",false,CM_C.OK_BG);
  else ex2.forEach(function(e){row("  * "+e.label,false,CM_C.EXCEPT_BG);row("      "+e.missing);});
  row("");
  row("EXCEPTION 3 — GAP > 2mm AND TDC MISSING OR LAPSED  ("+ex3.length+" entries)",true,CM_C.HEADER_BG);
  if (!ex3.length) row("  No exceptions found",false,CM_C.OK_BG);
  else ex3.forEach(function(e){row("  * "+e.label,false,CM_C.EXCEPT_BG);row("      Status: "+e.status);});
  row("");
  row("END OF GAPLESS CMS REPORT — "+today,true,CM_C.TITLE_BG);
  return out;
}
