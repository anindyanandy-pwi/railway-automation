// ============================================================
//  TWSTieBarReport.gs — Sheet 1
//  Updated with new color scheme and cache writing.
// ============================================================

var TWS_DATA_SHEET = "TWS TIE Bar";
var TWS_SECTION_ID = "TWS_TIE_BAR";
var TWS_TDC_DAYS   = 7;

function generateTWSTieBarReport() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet=ss.getSheetByName(TWS_DATA_SHEET);
  if (!dataSheet) {
    var sheets=ss.getSheets();
    for (var s=0;s<sheets.length;s++) {
      var n=sheets[s].getName().toLowerCase();
      if (n.indexOf("tws")>-1&&n.indexOf("tie")>-1) { dataSheet=sheets[s]; break; }
    }
  }
  if (!dataSheet) {
    try { SpreadsheetApp.getUi().alert('TWS Tie Bar sheet not found.'); }
    catch(_) { console.log('TWS Tie Bar sheet not found.'); }
    return;
  }
  var allData=dataSheet.getDataRange().getValues();
  var headerRowIdx=-1;
  for (var i=0;i<allData.length;i++) {
    for (var j=0;j<allData[i].length;j++) {
      var v=String(allData[i][j]).toLowerCase();
      if (v.indexOf("point no")>-1) { headerRowIdx=i; break; }
    }
    if (headerRowIdx>-1) break;
  }
  if (headerRowIdx===-1) {
    try { SpreadsheetApp.getUi().alert("Header row not found in TWS Tie Bar sheet."); }
    catch(_) {}
    return;
  }
  var headers=allData[headerRowIdx];
  var iAEN=sharedFindColIdx(headers,"aen"), iStation=sharedFindColIdx(headers,"station");
  var iLine=sharedFindColIdx(headers,"line"), iPointNo=sharedFindColIdx(headers,"point no");
  var iDateConv=sharedFindColIdx(headers,"date of tws"), iPhoto=sharedFindColIdx(headers,"tie bar fixing image");
  var iDateFixed=sharedFindColIdx(headers,"date of tie bar fixing"), iTDC=sharedFindColIdx(headers,"tdc");
  var iRemark=headers.length-1;
  var today=new Date(); today.setHours(0,0,0,0);
  var ex1=[],ex2=[],ex3=[],ex4=[],ex5=[];
  var scanned=0, skipped=0, currentAEN="";

  for (var r=headerRowIdx+1;r<allData.length;r++) {
    var row=allData[r];
    var rowAEN=String(row[iAEN]||"").trim();
    if (!sharedIsEmpty(rowAEN)) currentAEN=rowAEN;
    var pointNo=String(row[iPointNo]||"").trim();
    var station=String(row[iStation]||"").trim();
    var line=String(row[iLine]||"").trim();
    if (sharedIsEmpty(pointNo)||pointNo.toLowerCase().indexOf("point")>-1) continue;
    var label="Point: "+pointNo;
    if (!sharedIsEmpty(station)) label+=" | Stn: "+station;
    if (!sharedIsEmpty(line))    label+=" | Line: "+line;
    if (!sharedIsEmpty(currentAEN)) label+=" (AEN: "+currentAEN+")";
    var remark=String(row[iRemark]||"").trim();
    var remarkStr=sharedIsEmpty(remark)?"":"  [Remark: "+remark+"]";
    var photoVal=row[iPhoto], dateFixedVal=row[iDateFixed];
    var dateFixedStr=String(dateFixedVal||"").trim();
    if (!sharedIsEmpty(dateFixedStr)&&sharedHasPhoto(photoVal)) { skipped++; continue; }
    scanned++;
    var tdcDate=sharedParseDate(row[iTDC]);
    if (!tdcDate) {
      var convDate=sharedParseDate(row[iDateConv]);
      if (convDate) tdcDate=new Date(convDate.getTime()+TWS_TDC_DAYS*864e5);
    }
    if (!sharedHasPhoto(photoVal)) ex1.push(label+remarkStr);
    if (tdcDate&&tdcDate<today&&sharedIsEmpty(dateFixedStr)) {
      var days=Math.floor((today-tdcDate)/864e5);
      ex2.push({label:label, tdc:sharedFmtDate(tdcDate), days:days, remark:remarkStr});
    }
    if (tdcDate&&tdcDate<today&&sharedIsEmpty(dateFixedStr))
      ex3.push({label:label, tdc:sharedFmtDate(tdcDate), remark:remarkStr});
    if (sharedHasPhoto(photoVal)&&sharedIsEmpty(dateFixedStr)) ex4.push(label+remarkStr);
    if (!sharedIsEmpty(dateFixedStr)&&!sharedHasPhoto(photoVal)) ex5.push(label+remarkStr);
  }

  writeReportSection(TWS_SECTION_ID, _buildTWSOutput(scanned,skipped,ex1,ex2,ex3,ex4,ex5));

  // Write cache
  try {
    var summary=cm_buildSummary({
      "Tie bar fixing image missing": ex1,
      "TDC lapsed": ex2.map(function(e){return e.label;}),
      "TDC passed but date not entered": ex3.map(function(e){return e.label;}),
      "Image present but date missing": ex4,
      "Date present but image missing": ex5
    });
    cm_writeSummaryToCache("generateTWSTieBarReport", summary);
  } catch(e) { console.log("TWS cache write error: "+e.message); }

  var total=ex1.length+ex2.length+ex3.length+ex4.length+ex5.length;
  try { SpreadsheetApp.getUi().alert(
    "TWS Tie Bar Report updated.\n\nScanned: "+scanned+
    "  Completed (excluded): "+skipped+"\nTotal exceptions: "+total);
  } catch(_) { console.log("TWS Tie Bar: "+total+" exceptions."); }
}

function _buildTWSOutput(scanned,skipped,ex1,ex2,ex3,ex4,ex5) {
  var tot=ex1.length+ex2.length+ex3.length+ex4.length+ex5.length;
  var today=sharedFmtDate(new Date()), out=[];
  function row(t,b,c){ out.push([t||"",b||false,c||null]); }
  row("TWS TIE BAR — EXCEPTION REPORT",true,CM_C.TITLE_BG);
  row("Date: "+today+"   |   Howrah Division / Eastern Railway",false,CM_C.TITLE_BG);
  row("Scanned: "+scanned+"   |   Completed (excluded): "+skipped+"   |   Total exceptions: "+tot,false,CM_C.TITLE_BG);
  row("");
  function sect(title,items,fn){
    row(title+"  ("+items.length+" entries)",true,CM_C.HEADER_BG);
    if (!items.length) row("  No exceptions found",false,CM_C.OK_BG);
    else items.forEach(fn); row("");
  }
  sect("EXCEPTION 1 — TIE BAR FIXING IMAGE MISSING",ex1,function(e){row("  * "+e,false,CM_C.EXCEPT_BG);});
  row("EXCEPTION 2 — TDC LAPSED  ("+ex2.length+" entries)",true,CM_C.HEADER_BG);
  row("  (No Revised TDC accepted — flagged until work complete)",false,CM_C.HEADER_BG);
  if (!ex2.length) row("  No exceptions found",false,CM_C.OK_BG);
  else ex2.forEach(function(e){row("  * "+e.label+e.remark,false,CM_C.EXCEPT_BG);row("      TDC: "+e.tdc+"   |   Lapsed by: "+e.days+" days");});
  row("");
  sect("EXCEPTION 3 — TDC PASSED BUT DATE OF FIXING NOT ENTERED",ex3,
    function(e){row("  * "+e.label+e.remark,false,CM_C.EXCEPT_BG);row("      TDC was: "+e.tdc);});
  sect("EXCEPTION 4 — FIXING IMAGE PRESENT BUT DATE MISSING",ex4,function(e){row("  * "+e,false,CM_C.EXCEPT_BG);});
  sect("EXCEPTION 5 — DATE PRESENT BUT FIXING IMAGE MISSING",ex5,function(e){row("  * "+e,false,CM_C.EXCEPT_BG);});
  row("END OF REPORT — "+today,true,CM_C.TITLE_BG);
  return out;
}