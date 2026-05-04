// ============================================================
//  bad_road_surface_apps_script.gs — Sheet 1
//  Bad Road Surface + TWS Tie Bar exception reports.
//  Includes cache writing, updated color scheme.
// ============================================================

var SHARED_REPORT_SHEET = "Exception Report";
var BRS_DATA_SHEET      = "Bad Road Surface";
var BRS_SECTION_ID      = "BAD_ROAD_SURFACE";

// ── Shared helpers ────────────────────────────────────────────────────────────
function sharedIsEmpty(val) {
  if (val===null||val===undefined) return true;
  return ["","---","--","----","nan","none","#value!","#ref!","error","false","no"]
    .indexOf(String(val).trim().toLowerCase()) > -1;
}
function sharedHasPhoto(val) {
  if (!val) return false;
  if (val !== null && typeof val === "object" && !Array.isArray(val)) return true; // CellImage
  var s=String(val).trim();
  if (!s || sharedIsEmpty(s)) return false;
  if (s === "CellImage") return true;
  var l=s.toLowerCase();
  // Only return true for actual photo indicators — NOT for arbitrary text
  return l.indexOf("hyperlink")>-1 || l.indexOf("image(")>-1 ||
         l.indexOf("http")>-1 || l.indexOf("drive.google")>-1 ||
         /\.(jpg|jpeg|png|gif)/i.test(s);
}
function sharedParseDate(val) {
  if (sharedIsEmpty(val)) return null;
  if (val instanceof Date && !isNaN(val)) return val;
  var s=String(val).trim(), sep=s.indexOf(".")>-1?".":"/";
  var p=s.split(sep); if (p.length!==3) return null;
  var d=parseInt(p[0]),m=parseInt(p[1])-1,y=parseInt(p[2]);
  if (y<100) y+=2000;
  var dt=new Date(y,m,d); return isNaN(dt.getTime())?null:dt;
}
function sharedFmtDate(d) {
  if (!d) return "";
  return ('0'+d.getDate()).slice(-2)+'.'+('0'+(d.getMonth()+1)).slice(-2)+'.'+d.getFullYear();
}
function sharedFindColIdx(headers, keyword) {
  var kw=keyword.toLowerCase();
  for (var i=0;i<headers.length;i++)
    if (String(headers[i]).toLowerCase().indexOf(kw)>-1) return i;
  return -1;
}

// ── Section Writer ────────────────────────────────────────────────────────────
function writeReportSection(sectionId, outputRows) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getSheetByName(SHARED_REPORT_SHEET);
  if (!sheet) sheet=ss.insertSheet(SHARED_REPORT_SHEET);
  var SM="##START:"+sectionId+"##", EM="##END:"+sectionId+"##";
  var lastRow=sheet.getLastRow(), startRow=-1, endRow=-1;
  if (lastRow>0) {
    var colA=sheet.getRange(1,1,lastRow,1).getValues();
    for (var i=0;i<colA.length;i++) {
      if (colA[i][0]===SM) startRow=i+1;
      if (colA[i][0]===EM) endRow=i+1;
    }
  }
  var newRows=outputRows.length+3;
  if (startRow>0&&endRow>0) {
    var oldCount=endRow-startRow+1;
    if (newRows>oldCount) sheet.insertRowsBefore(startRow,newRows-oldCount);
    else if (newRows<oldCount) {
      var ds=startRow+newRows, dc=oldCount-newRows, mx=sheet.getLastRow();
      if (ds<=mx) sheet.deleteRows(ds,Math.min(dc,mx-ds+1));
    }
  } else {
    startRow=(lastRow===0)?1:lastRow+2;
    var need=startRow+newRows-1, maxR=sheet.getMaxRows();
    if (need>maxR) sheet.insertRowsAfter(maxR,need-maxR+20);
  }
  _brsWriteRows(sheet,startRow,SM,EM,outputRows);
  sheet.autoResizeColumn(1);
  ss.setActiveSheet(sheet);
}

function _brsWriteRows(sheet,startRow,SM,EM,outputRows) {
  var V=[],W=[],B=[],F=[];
  V.push([SM]); W.push(["normal"]); B.push([CM_C.WHITE]); F.push([CM_C.WHITE]);
  V.push([""]); W.push(["normal"]); B.push([CM_C.WHITE]); F.push([CM_C.DATA_FG]);
  for (var r=0;r<outputRows.length;r++) {
    var row=outputRows[r], bg=row[2]||CM_C.WHITE;
    var fg=row[2]?CM_C.BLACK:CM_C.DATA_FG;
    V.push([row[0]||""]); W.push([row[1]?"bold":"normal"]); B.push([bg]); F.push([fg]);
  }
  V.push([EM]); W.push(["normal"]); B.push([CM_C.WHITE]); F.push([CM_C.WHITE]);
  var rng=sheet.getRange(startRow,1,V.length,1);
  rng.setValues(V); rng.setFontWeights(W); rng.setBackgrounds(B); rng.setFontColors(F);
}

// ── Bad Road Surface ──────────────────────────────────────────────────────────
// Photo detection helper — checks value AND formula (catches =IMAGE() / =HYPERLINK() cells)
function _brsHasPhoto(val, formula) {
  // CellImage object (in-cell embedded image via newer API)
  if (val !== null && val !== undefined && typeof val === 'object' && !Array.isArray(val)) return true;
  // URL / HYPERLINK() / IMAGE() stored as plain value
  if (sharedHasPhoto(val)) return true;
  // Formula-based images: =IMAGE(...) or =HYPERLINK(...)
  if (formula) {
    var f = String(formula).toLowerCase();
    return f.indexOf("image(") > -1 || f.indexOf("hyperlink(") > -1 || f.indexOf("http") > -1;
  }
  return false;
}

function generateBadRoadReport() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet=ss.getSheetByName(BRS_DATA_SHEET);
  if (!dataSheet) {
    try { SpreadsheetApp.getUi().alert('Sheet "'+BRS_DATA_SHEET+'" not found!'); }
    catch(_) { console.log('Sheet "'+BRS_DATA_SHEET+'" not found!'); }
    return;
  }
  var dataRange = dataSheet.getDataRange();
  var allData     = dataRange.getValues();
  var allFormulas = dataRange.getFormulas(); // needed to detect =IMAGE() / =HYPERLINK() photos

  var headerRowIdx=-1;
  for (var i=0;i<allData.length;i++) {
    for (var j=0;j<allData[i].length;j++) {
      if (String(allData[i][j]).toLowerCase().indexOf("name of lc")>-1) { headerRowIdx=i; break; }
    }
    if (headerRowIdx>-1) break;
  }
  if (headerRowIdx===-1) {
    try { SpreadsheetApp.getUi().alert("Could not find header row in Bad Road Surface sheet."); }
    catch(_) {}
    return;
  }
  var headers=allData[headerRowIdx];
  var iLC=sharedFindColIdx(headers,"name of lc"), iAEN=sharedFindColIdx(headers,"aen");
  var iLine=sharedFindColIdx(headers,"line"), iKM=sharedFindColIdx(headers,"km");
  var iRoadPhoto=sharedFindColIdx(headers,"photo of road surface");
  var iTDC=sharedFindColIdx(headers,"tdc for closing");
  var iRevisedTDC=sharedFindColIdx(headers,"revised tdc");
  var iDatePitching=sharedFindColIdx(headers,"date of pitching");
  var iClosingPhoto=sharedFindColIdx(headers,"closing photo");
  var today=new Date(); today.setHours(0,0,0,0);
  var ex1=[],ex2=[],ex3=[],ex4=[],ex5=[];
  var scanned=0, skipped=0, curAEN="";

  for (var r=headerRowIdx+1;r<allData.length;r++) {
    var row=allData[r];
    var rowF=allFormulas[r];

    // Skip if LC cell is a Date object (date value bled into LC column)
    if (row[iLC] instanceof Date) continue;

    var lc=String(row[iLC]||"").trim();

    // Skip empty, junk, single-dash placeholder rows, and header repetitions
    if (!lc || lc==="-" || lc==="–" || lc==="—" ||
        sharedIsEmpty(lc) || lc.toLowerCase().indexOf("name of lc")>-1) continue;

    var aen=String(row[iAEN]||"").trim();
    if (!sharedIsEmpty(aen)) curAEN=aen; else aen=curAEN;
    var line=String(row[iLine]||"").trim();
    var km=String(row[iKM]||"").trim();
    var label="LC: "+lc;
    if (!sharedIsEmpty(line)) label+=" | Line: "+line;
    if (!sharedIsEmpty(km))   label+=" | KM: "+km;
    if (!sharedIsEmpty(aen))  label+=" (AEN: "+aen+")";

    // Skip completely empty placeholder rows:
    // If road photo AND TDC are both truly empty (not even '--'), the row is unregistered
    var rdPhotoRaw = iRoadPhoto>=0 ? String(row[iRoadPhoto]||"").trim() : "";
    var tdcRaw     = iTDC>=0 ? String(row[iTDC]||"").trim() : "";
    if (rdPhotoRaw==="" && tdcRaw==="") continue;

    var datePitchingVal=row[iDatePitching], closingPhotoVal=row[iClosingPhoto];
    var datePitchingStr=String(datePitchingVal||"").trim();
    // Date object in pitching column — convert properly
    if (datePitchingVal instanceof Date && !isNaN(datePitchingVal.getTime()))
      datePitchingStr = sharedFmtDate(datePitchingVal);

    var closingPhotoFormula = iClosingPhoto>=0&&rowF ? rowF[iClosingPhoto] : "";
    var closingPhotoDone = _brsHasPhoto(closingPhotoVal, closingPhotoFormula);

    // Work fully complete: pitching date filled AND closing photo present
    // Work fully complete: pitching date filled AND closing photo present
    // Both required — closing photo alone means work done but date not recorded (→ Ex4)
    // Pitching date alone means date recorded but photo missing (→ Ex5)
    if (!sharedIsEmpty(datePitchingStr) && closingPhotoDone) { skipped++; continue; }
    scanned++;

    // Exception 1: Road surface photo missing
    var roadPhotoFormula = iRoadPhoto>=0&&rowF ? rowF[iRoadPhoto] : "";
    if (!_brsHasPhoto(row[iRoadPhoto], roadPhotoFormula)) ex1.push(label);

    var tdcDate=sharedParseDate(row[iTDC]);
    // Fix: parse LATEST date from possibly multiline Revised TDC cell
    // e.g. "15.04.26\n30.04.2026" → take 30.04.2026 (most recent extension)
    var revisedStr=String(row[iRevisedTDC]||"").trim();
    var revisedDate = null;
    if (!sharedIsEmpty(revisedStr)) {
      var revParts = revisedStr.split("\n");
      var revDates = [];
      for (var rp=0; rp<revParts.length; rp++) {
        var rpd = sharedParseDate(revParts[rp].trim());
        if (rpd) revDates.push(rpd);
      }
      if (revDates.length > 0) {
        revisedDate = revDates.reduce(function(a,b){ return a>b?a:b; });
      }
    }

    // Exception 2: TDC lapsed — ONLY flag if no closing photo either
    if (tdcDate&&tdcDate<today&&!revisedDate&&
        sharedIsEmpty(datePitchingStr)&&!closingPhotoDone)
      ex2.push({label:label, tdc:sharedFmtDate(tdcDate), days:Math.floor((today-tdcDate)/864e5)});

    // Exception 3: Revised TDC lapsed — ONLY flag if no closing photo
    if (revisedDate&&revisedDate<today&&sharedIsEmpty(datePitchingStr)&&!closingPhotoDone)
      ex3.push({label:label, tdc:sharedFmtDate(revisedDate), days:Math.floor((today-revisedDate)/864e5)});

    // Exception 4: Closing photo present but pitching date not filled
    if (closingPhotoDone&&sharedIsEmpty(datePitchingStr)) ex4.push(label);

    // Exception 5: Pitching date filled but closing photo missing
    if (!sharedIsEmpty(datePitchingStr)&&!closingPhotoDone) ex5.push(label);
  }

  var out=_buildBadRoadOutput(scanned,skipped,ex1,ex2,ex3,ex4,ex5);
  writeReportSection(BRS_SECTION_ID, out);

  // Write cache summary
  try {
    var summary = cm_buildSummary({
      "Road surface photo missing": ex1,
      "TDC lapsed": ex2.map(function(e){return e.label;}),
      "Revised TDC lapsed": ex3.map(function(e){return e.label;}),
      "Closing photo present but pitching date missing": ex4,
      "Pitching date present but closing photo missing": ex5
    });
    cm_writeSummaryToCache("generateBadRoadReport", summary);
  } catch(e) { console.log("BRS cache write error: "+e.message); }

  var total=ex1.length+ex2.length+ex3.length+ex4.length+ex5.length;
  try { SpreadsheetApp.getUi().alert(
    "Bad Road Surface Report updated.\n\nScanned: "+scanned+
    "  Completed (excluded): "+skipped+"\nTotal exceptions: "+total);
  } catch(_) { console.log("Bad Road Surface: "+total+" exceptions."); }
}

function _buildBadRoadOutput(scanned,skipped,ex1,ex2,ex3,ex4,ex5) {
  var totalEx=ex1.length+ex2.length+ex3.length+ex4.length+ex5.length;
  var today=sharedFmtDate(new Date()), out=[];
  function row(t,b,c){ out.push([t||"",b||false,c||null]); }
  row("BAD ROAD SURFACE — EXCEPTION REPORT",true,CM_C.TITLE_BG);
  row("Date: "+today+"   |   Howrah Division / Eastern Railway",false,CM_C.TITLE_BG);
  row("Scanned: "+scanned+"   |   Completed (excluded): "+skipped+"   |   Total exceptions: "+totalEx,false,CM_C.TITLE_BG);
  row("");
  function sect(title,items,fn){
    row(title+"  ("+items.length+" entries)",true,CM_C.HEADER_BG);
    if (!items.length) row("  No exceptions found",false,CM_C.OK_BG);
    else items.forEach(fn); row("");
  }
  sect("EXCEPTION 1 — ROAD SURFACE PHOTO MISSING",ex1,function(e){row("  * "+e,false,CM_C.EXCEPT_BG);});
  row("EXCEPTION 2 — TDC FOR CLOSING LAPSED  ("+ex2.length+" entries)",true,CM_C.HEADER_BG);
  row("  (No Revised TDC filled; work not complete)",false,CM_C.HEADER_BG);
  if (!ex2.length) row("  No exceptions found",false,CM_C.OK_BG);
  else ex2.forEach(function(e){row("  * "+e.label,false,CM_C.EXCEPT_BG);row("      TDC: "+e.tdc+"   |   Lapsed by: "+e.days+" days");});
  row("");
  row("EXCEPTION 3 — REVISED TDC LAPSED  ("+ex3.length+" entries)",true,CM_C.HEADER_BG);
  row("  (Work not complete)",false,CM_C.HEADER_BG);
  if (!ex3.length) row("  No exceptions found",false,CM_C.OK_BG);
  else ex3.forEach(function(e){row("  * "+e.label,false,CM_C.EXCEPT_BG);row("      Revised TDC: "+e.tdc+"   |   Lapsed by: "+e.days+" days");});
  row("");
  sect("EXCEPTION 4 — CLOSING PHOTO PRESENT BUT DATE OF PITCHING MISSING",ex4,function(e){row("  * "+e,false,CM_C.EXCEPT_BG);});
  sect("EXCEPTION 5 — DATE OF PITCHING PRESENT BUT CLOSING PHOTO MISSING",ex5,function(e){row("  * "+e,false,CM_C.EXCEPT_BG);});
  row("END OF REPORT — "+today,true,CM_C.TITLE_BG);
  return out;
}