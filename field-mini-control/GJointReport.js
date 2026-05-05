// ============================================================
//  GJointReport.gs — Sheet 1 | Physically Damaged G/Joint
//  Updated with new color scheme and cache writing.
// ============================================================

var GJ_SECTION_ID = "GJOINT";
var GJ_MAX_COL    = 14;

function generateGJointReport() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet=ss.getSheetByName("Physically Damaged G/Joint");
  if (!dataSheet) {
    var sheets=ss.getSheets();
    for (var s=0;s<sheets.length;s++) {
      var n=sheets[s].getName().toLowerCase();
      if (n.indexOf("glued")>-1||n.indexOf("g/joint")>-1||n.indexOf("damaged")>-1) { dataSheet=sheets[s]; break; }
    }
  }
  if (!dataSheet) {
    try { SpreadsheetApp.getUi().alert('G/Joint sheet not found.'); }
    catch(_) { console.log('G/Joint sheet not found.'); }
    return;
  }
  var allData=dataSheet.getDataRange().getValues();
  var hdrIdx=-1;
  for (var i=0;i<allData.length;i++) {
    for (var j=0;j<GJ_MAX_COL;j++) {
      if (String(allData[i][j]).toLowerCase().indexOf("chainage")>-1) { hdrIdx=i; break; }
    }
    if (hdrIdx>-1) break;
  }
  if (hdrIdx<0) {
    try { SpreadsheetApp.getUi().alert("Header row not found in G/Joint sheet."); }
    catch(_) {}
    return;
  }
  var hdr=allData[hdrIdx];
  var iAEN=_gjC(hdr,"aen"), iPWI=_gjC(hdr,"pwi"), iLine=_gjC(hdr,"line");
  var iLoc=_gjC(hdr,"location"), iCh=_gjC(hdr,"chainage"), iLHRH=_gjC(hdr,"lh");
  var iGlued=_gjC(hdr,"glued"), iPhoto=_gjC(hdr,"photo of glued");
  var iRevTDC=-1;
  for (var c=0;c<GJ_MAX_COL;c++) { if (String(hdr[c]).toLowerCase().indexOf("revised")>-1) { iRevTDC=c; break; } }
  var iTDC=-1;
  for (var c2=0;c2<GJ_MAX_COL;c2++) {
    var hv=String(hdr[c2]).toLowerCase();
    if (hv==="tdc"||(hv.indexOf("tdc")>-1&&hv.indexOf("revised")<0)) { iTDC=c2; break; }
  }
  var iRectPh=_gjC(hdr,"rectified photo"), iRectDt=_gjC(hdr,"rectification date");
  var today=new Date(); today.setHours(0,0,0,0);
  var ex1=[],ex2=[],ex3=[],ex4=[],ex5=[], scanned=0,skipped=0;
  var curAEN="",curPWI="";
  for (var r=hdrIdx+1;r<allData.length;r++) {
    var row=allData[r];
    var ra=_gjStr(row,iAEN), rp=_gjStr(row,iPWI);
    if (!_gjE(ra)) curAEN=ra;
    if (!_gjE(rp)) curPWI=rp;
    var line=_gjStr(row,iLine), loc=_gjStr(row,iLoc), ch=_gjStr(row,iCh);
    var lhrh=_gjStr(row,iLHRH), glued=_gjStr(row,iGlued);
    var anyFilled=!_gjE(line)||!_gjE(loc)||!_gjE(ch)||!_gjE(lhrh)||!_gjE(glued);
    if (!anyFilled) continue;
    var allDash=_gjAllDash(line)&&_gjAllDash(loc)&&_gjAllDash(ch)&&_gjAllDash(lhrh)&&_gjAllDash(glued);
    if (allDash) continue;
    scanned++;
    var label="AEN: "+curAEN+" | PWI: "+curPWI;
    if (!_gjE(line))  label+=" | Line: "+line;
    if (!_gjE(ch))    label+=" | Chainage: "+ch;
    if (!_gjE(lhrh))  label+=" | "+lhrh;
    var photoVal=row[iPhoto], tdcVal=row[iTDC], revTDCVal=row[iRevTDC];
    var rectPhVal=row[iRectPh], rectDtVal=row[iRectDt];
    var tdcStr=_gjStr(row,iTDC), revTDCStr=_gjStr(row,iRevTDC), rectDtStr=_gjStr(row,iRectDt);
    var tdcDate=_gjDate(tdcVal), revDate=_gjDate(revTDCVal);
    var workDone=_gjPhoto(rectPhVal)&&!_gjE(rectDtStr);
    if (workDone) { skipped++; scanned--; continue; }
    var missing1=[];
    if (_gjE(line))  missing1.push("LINE");
    if (_gjE(loc))   missing1.push("LOCATION");
    if (_gjE(lhrh))  missing1.push("LH/RH");
    if (_gjE(glued)) missing1.push("GLUED NO (TMS)");
    if (!_gjPhoto(photoVal)) missing1.push("PHOTO OF GLUED JOINT");
    if (_gjE(tdcStr)) missing1.push("TDC");
    if (missing1.length>0) ex1.push({label:label, missing:missing1.join(", ")});
    if (!_gjPhoto(photoVal)) ex2.push({label:label, type:"Photo of Glued Joint missing"});
    if (!_gjE(revTDCStr)&&!_gjPhoto(rectPhVal)) ex2.push({label:label, type:"Rectified Photo missing (Revised TDC filled)"});
    if (tdcDate&&tdcDate<today) {
      var daysLapsed=Math.floor((today-tdcDate)/864e5);
      var rectMissing=!_gjPhoto(rectPhVal)||_gjE(rectDtStr);
      if (rectMissing) {
        var what=_gjMissingRectDesc(rectPhVal,rectDtStr);
        if (!_gjE(revTDCStr)) ex3.push({label:label, tdc:_gjFmt(tdcDate), days:daysLapsed, what:what});
        else ex4.push({label:label, tdc:_gjFmt(tdcDate), days:daysLapsed, what:what});
      }
    }
    if (revDate&&revDate<today) ex5.push({label:label, tdc:_gjFmt(revDate), days:Math.floor((today-revDate)/864e5)});
  }
  writeReportSection(GJ_SECTION_ID, _buildGJOutput(scanned,skipped,ex1,ex2,ex3,ex4,ex5));
  // Cache — break mandatory missing into per-column counts for concise summary
  try {
    // Count which specific columns are missing per entry
    var colExTypes = {};
    ex1.forEach(function(e) {
      // e.missing = "LINE, CHAINAGE, PHOTO OF GLUED JOINT" etc
      var cols = e.missing.split(",").map(function(c){ return c.trim().replace(/ missing$/i,""); });
      cols.forEach(function(col) {
        var key = col + " missing";
        if (!colExTypes[key]) colExTypes[key] = [];
        colExTypes[key].push(e.label);
      });
    });
    var exTypesObj = colExTypes;
    exTypesObj["Photo not uploaded"]          = ex2.map(function(e){return e.label;});
    exTypesObj["TDC lapsed (Revised TDC set"] = ex3.map(function(e){return e.label;});
    exTypesObj["TDC lapsed (no Revised TDC)"] = ex4.map(function(e){return e.label;});
    exTypesObj["Revised TDC lapsed"]          = ex5.map(function(e){return e.label;});
    var summary = cm_buildSummary(exTypesObj);
    cm_writeSummaryToCache("generateGJointReport", summary);
  } catch(e) { console.log("GJoint cache error: "+e.message); }
  var tot=ex1.length+ex2.length+ex3.length+ex4.length+ex5.length;
  try { SpreadsheetApp.getUi().alert("G/Joint Report updated.\n\nScanned: "+scanned+"  Work complete: "+skipped+"\nTotal: "+tot); }
  catch(_) { console.log("GJoint: "+tot+" exceptions."); }
}

function _gjMissingRectDesc(p,d) {
  var m=[];
  if (!_gjPhoto(p)) m.push("Rectified Photo");
  if (_gjE(d)) m.push("Rectification Date");
  return m.length ? m.join(" and ")+" missing" : "";
}
var _GJ_JUNK=["","---","--","----","-----","nan","none","nil","na","n/a","-","–"];
function _gjE(v) {
  if (v===null||v===undefined) return true;
  return _GJ_JUNK.indexOf(String(v).trim().toLowerCase())>-1;
}
function _gjAllDash(v) {
  if (v===null||v===undefined) return true;
  var s=String(v).trim();
  return s===""||/^[-–—]+$/.test(s)||s.toLowerCase()==="nil"||s.toLowerCase()==="na"||s.toLowerCase()==="n/a";
}
function _gjPhoto(v) {
  if (!v) return false; var s=String(v).trim(); if (_gjE(s)) return false;
  var l=s.toLowerCase();
  return l.indexOf("hyperlink")>-1||l.indexOf("image(")>-1||l.indexOf("http")>-1||l.indexOf("drive.google")>-1||
         /\.(jpg|jpeg|png|gif)/i.test(s);
}
function _gjDate(v) {
  if (!v) return null; if (v instanceof Date&&!isNaN(v.getTime())) return v;
  var s=String(v).split('\n')[0].split('\r')[0].trim(); if (_gjE(s)) return null;
  var m=s.match(/(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})/); if (!m) return null;
  var d=parseInt(m[1]),mo=parseInt(m[2])-1,y=parseInt(m[3]); if (y<100) y+=2000;
  var dt=new Date(y,mo,d); return isNaN(dt.getTime())?null:dt;
}
function _gjFmt(d) {
  if (!d) return "";
  return ('0'+d.getDate()).slice(-2)+'.'+('0'+(d.getMonth()+1)).slice(-2)+'.'+d.getFullYear();
}
function _gjC(hdr,kw) {
  var k=kw.toLowerCase();
  for (var i=0;i<GJ_MAX_COL;i++) if (String(hdr[i]).toLowerCase().indexOf(k)>-1) return i;
  return -1;
}
function _gjStr(row,i) {
  if (i<0||i>=row.length) return "";
  var v = row[i];
  // Handle Date objects bleeding into wrong columns (GAS getValues() returns Date objects)
  if (v instanceof Date && !isNaN(v.getTime())) return sharedFmtDate(v);
  return String(v||"").trim();
}
function _buildGJOutput(scanned,skipped,ex1,ex2,ex3,ex4,ex5) {
  var tot=ex1.length+ex2.length+ex3.length+ex4.length+ex5.length, out=[];
  var today=_gjFmt(new Date());
  function row(t,b,c){ out.push([t||"",b||false,c||null]); }
  row("PHYSICALLY DAMAGED G/JOINT — EXCEPTION REPORT",true,CM_C.TITLE_BG);
  row("Date: "+today+"   |   Howrah Division / Eastern Railway",false,CM_C.TITLE_BG);
  row("Entries scanned: "+scanned+"   |   Work complete: "+skipped+"   |   Total exceptions: "+tot,false,CM_C.TITLE_BG);
  row("");
  function sect2(title,items,fn){
    row(title+"  ("+items.length+" entries)",true,CM_C.HEADER_BG);
    if (!items.length) row("  No exceptions found",false,CM_C.OK_BG);
    else items.forEach(fn); row("");
  }
  row("EXCEPTION 1 — MANDATORY COLUMNS MISSING  ("+ex1.length+" entries)",true,CM_C.HEADER_BG);
  if (!ex1.length) row("  No exceptions found",false,CM_C.OK_BG);
  else ex1.forEach(function(e){row("  * "+e.label,false,CM_C.EXCEPT_BG);row("      Missing: "+e.missing);});
  row("");
  row("EXCEPTION 2 — PHOTO NOT UPLOADED  ("+ex2.length+" entries)",true,CM_C.HEADER_BG);
  if (!ex2.length) row("  No exceptions found",false,CM_C.OK_BG);
  else ex2.forEach(function(e){row("  * "+e.label,false,CM_C.EXCEPT_BG);row("      "+e.type);});
  row("");
  row("EXCEPTION 3 — TDC LAPSED, REVISED TDC EXISTS  ("+ex3.length+" entries)",true,CM_C.HEADER_BG);
  if (!ex3.length) row("  No exceptions found",false,CM_C.OK_BG);
  else ex3.forEach(function(e){row("  * "+e.label,false,CM_C.EXCEPT_BG);row("      TDC: "+e.tdc+"   |   Lapsed: "+e.days+" days   |   "+e.what);});
  row("");
  row("EXCEPTION 4 — TDC LAPSED, NO REVISED TDC  ("+ex4.length+" entries)",true,CM_C.HEADER_BG);
  if (!ex4.length) row("  No exceptions found",false,CM_C.OK_BG);
  else ex4.forEach(function(e){row("  * "+e.label,false,CM_C.EXCEPT_BG);row("      TDC: "+e.tdc+"   |   Lapsed: "+e.days+" days   |   "+e.what);});
  row("");
  sect2("EXCEPTION 5 — REVISED TDC LAPSED",ex5,function(e){row("  * "+e.label,false,CM_C.EXCEPT_BG);row("      Revised TDC: "+e.tdc+"   |   Lapsed: "+e.days+" days");});
  row("END OF G/JOINT REPORT — "+today,true,CM_C.TITLE_BG);
  return out;
}