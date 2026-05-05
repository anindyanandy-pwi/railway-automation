// ============================================================
//  _SharedUtils.gs — Sheet 2 ONLY
//  Shared helpers, section writer, manual send-all, menu.
// ============================================================

var S2_REPORT_SHEET = "Exception Report";
var S2_FOLDER_ID    = "1InCD6mlBHmvuvtgp1ycTY-bNkZxwqbld";
var S2_HOUR         = 9;

var S2_RECIPIENTS = [
  "adenbdcer@gmail.com","srdenchwh@gmail.com","dkaeaen@gmail.com",
  "aentrhwh@gmail.com","deepsonamoni@gmail.com","aenndae@gmail.com",
  "adenndaeminicontrol@gmail.com","maityshuvajit22@gmail.com",
  "aenbhper@gmail.com","enggctlhwh@gmail.com","aenazimganj@gmail.com",
  "aenofficial7@gmail.com","aendkaeminicontrol@gmail.com",
  "er.aenrph@gmail.com","aen1llh@gmail.com","adenazminicontrol@gmail.com",
  "aenkatwa.er@gmail.com"
];

// ── Shared helpers ────────────────────────────────────────────────────────────
function gEmpty(v) {
  if (v===null||v===undefined) return true;
  return ["","---","--","----","nan","none","#value!","#ref!","error","false","no","-","–"]
    .indexOf(String(v).trim().toLowerCase())>-1;
}
function gPhoto(v) {
  if (!v) return false; var s=String(v).trim(); if (gEmpty(s)) return false;
  var l=s.toLowerCase();
  return l.indexOf("hyperlink")>-1||l.indexOf("image(")>-1||
         l.indexOf("http")>-1||l.indexOf("drive.google")>-1||
         /\.(jpg|jpeg|png|gif)/i.test(s);
}
function gDate(v) {
  if (!v) return null; if (v instanceof Date&&!isNaN(v.getTime())) return v;
  var s=String(v).trim(); if (gEmpty(s)) return null;
  var p=s.split(/[.\/-]/); if (p.length!==3) return null;
  var d=parseInt(p[0]),m=parseInt(p[1])-1,y=parseInt(p[2]); if (y<100) y+=2000;
  var dt=new Date(y,m,d); return isNaN(dt.getTime())?null:dt;
}
function gFmt(d) {
  if (!d) return "";
  return ('0'+d.getDate()).slice(-2)+'.'+('0'+(d.getMonth()+1)).slice(-2)+'.'+d.getFullYear();
}
function gDays(d1,d2) { return Math.floor((d2-d1)/864e5); }
function gCol(hdr,kw) {
  var k=kw.toLowerCase();
  for (var i=0;i<hdr.length;i++) if (String(hdr[i]).toLowerCase().indexOf(k)>-1) return i;
  return -1;
}
function gV(row,i) { return (i>=0&&i<row.length)?row[i]:""; }
function gStr(row,i) { return String(gV(row,i)||"").trim(); }
function gFindSheet(ss,names) {
  var sheets=ss.getSheets(), nl=names.map(function(n){return n.toLowerCase();});
  for (var s=0;s<sheets.length;s++) {
    if (nl.indexOf(sheets[s].getName().toLowerCase())>-1) return sheets[s];
  }
  for (var t=0;t<sheets.length;t++) {
    var sn=sheets[t].getName().toLowerCase();
    for (var n=0;n<nl.length;n++) { if (sn.indexOf(nl[n])>-1) return sheets[t]; }
  }
  return null;
}
function gToday() { return gFmt(new Date()); }

// ── Section writer (Sheet 2) ──────────────────────────────────────────────────
function s2WriteSection(sectionId, outputRows) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getSheetByName(S2_REPORT_SHEET);
  if (!sheet) sheet=ss.insertSheet(S2_REPORT_SHEET);
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
  } else { startRow=(lastRow===0)?1:lastRow+2; }
  _s2WriteRows(sheet,startRow,SM,EM,outputRows);
  sheet.autoResizeColumn(1);
}

function _s2WriteRows(sheet,startRow,SM,EM,outputRows) {
  var V=[],W=[],B=[],F=[];
  V.push([SM]); W.push(["normal"]); B.push([S2_C.WHITE]); F.push([S2_C.WHITE]);
  V.push([""]); W.push(["normal"]); B.push([S2_C.WHITE]); F.push([S2_C.DATA_FG]);
  for (var r=0;r<outputRows.length;r++) {
    var row=outputRows[r], bg=row[2]||S2_C.WHITE;
    var fg=row[2]?S2_C.BLACK:S2_C.DATA_FG;
    V.push([row[0]||""]); W.push([row[1]?"bold":"normal"]); B.push([bg]); F.push([fg]);
  }
  V.push([EM]); W.push(["normal"]); B.push([S2_C.WHITE]); F.push([S2_C.WHITE]);
  var rng=sheet.getRange(startRow,1,V.length,1);
  rng.setValues(V); rng.setFontWeights(W); rng.setBackgrounds(B); rng.setFontColors(F);
}

// ── PDF export + email (Sheet 2) ─────────────────────────────────────────────
function _sCallFn(fnName) {
  var fn=this[fnName];
  if (typeof fn==="function") fn();
  else throw new Error("Function not found: "+fnName);
}

function s2RunOne(sectionId, reportName, fnName, recipients) {
  var dateStr=gToday();
  try {
    _sCallFn(fnName);
    var tmp=s2Extract(S2_REPORT_SHEET,sectionId);
    if (!tmp) throw new Error("Section '"+sectionId+"' not found.");
    var fname=reportName+"_Exception_"+dateStr+".pdf";
    var pdfBlob=s2ToPDF(tmp,fname);
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(tmp);
    var folder=DriveApp.getFolderById(S2_FOLDER_ID);
    var sub=folder.getFoldersByName(dateStr);
    var dayFolder=sub.hasNext()?sub.next():folder.createFolder(dateStr);
    var driveFile=dayFolder.createFile(pdfBlob);
    s2Email(recipients,reportName.replace(/_/g," "),dateStr,pdfBlob,driveFile.getUrl());
    console.log("SUCCESS: "+reportName+" — "+dateStr);
  } catch(e) {
    try { var t=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("_TMP_EXPORT_");
          if (t) SpreadsheetApp.getActiveSpreadsheet().deleteSheet(t); } catch(_) {}
    try { GmailApp.sendEmail("enggplanningcellhwh@gmail.com",
            "WARNING: Exception Report Failed — "+reportName+" — "+dateStr,
            "Error: "+e.message); } catch(_) {}
    console.log("ERROR in "+reportName+": "+e.message);
  }
}

function s2Extract(reportSheetName, sectionId) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var rsh=ss.getSheetByName(reportSheetName); if (!rsh) return null;
  var lr=rsh.getLastRow(); if (lr<1) return null;
  var col=rsh.getRange(1,1,lr,1).getValues();
  var SM="##START:"+sectionId+"##", EM="##END:"+sectionId+"##";
  var sr=-1,er=-1;
  for (var i=0;i<col.length;i++) {
    if (col[i][0]===SM) sr=i+1;
    if (col[i][0]===EM) er=i+1;
  }
  if (sr<0||er<0) return null;
  var cs=sr+2,ce=er-1; if (cs>ce) return null;
  var nR=ce-cs+1, nC=rsh.getLastColumn();
  var src=rsh.getRange(cs,1,nR,nC);
  var tmp=ss.insertSheet("_TMP_EXPORT_");
  tmp.setColumnWidth(1,960); tmp.setHiddenGridlines(true);
  var dst=tmp.getRange(1,1,nR,nC);
  dst.setValues(src.getValues()); dst.setBackgrounds(src.getBackgrounds());
  dst.setFontColors(src.getFontColors()); dst.setFontWeights(src.getFontWeights());
  if (nC>1) tmp.hideColumns(2,nC-1);
  return tmp;
}

function s2ToPDF(sheet,filename) {
  var id=SpreadsheetApp.getActiveSpreadsheet().getId();
  var url="https://docs.google.com/spreadsheets/d/"+id+
    "/export?format=pdf&gid="+sheet.getSheetId()+
    "&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false"+
    "&sheetnames=false&pagenumbers=false&top_margin=0.5&bottom_margin=0.5"+
    "&left_margin=0.5&right_margin=0.5";
  var res=UrlFetchApp.fetch(url,{
    headers:{Authorization:"Bearer "+ScriptApp.getOAuthToken()},muteHttpExceptions:true});
  if (res.getResponseCode()!==200) throw new Error("PDF export HTTP "+res.getResponseCode());
  return res.getBlob().setName(filename);
}

function s2Email(recipients,reportName,dateStr,blob,driveUrl) {
  var subject="Howrah Division Exception Report — "+reportName+" — "+dateStr;
  var plain="Dear Sir/Ma'am,\n\nPlease find attached the Exception Report for "+reportName+
    " as on "+dateStr+".\n\nDrive link: "+driveUrl+
    "\n\nRegards,\nException Report System\nHowrah Division, Eastern Railway";
  GmailApp.sendEmail(recipients.join(","),subject,plain,{
    attachments:[blob],name:"Exception Report — Howrah Division"});
}

// ── Manual send-all (non-blocking chain) ─────────────────────────────────────
var S2_QUEUE_KEY="S2_REPORT_QUEUE";
var S2_CHAIN_FNS=["s2_SandHump","s2_Buffer","s2_ROB","s2_FOB","s2_Bridge"];

function manualSendAll_Sheet2() {
  try {
    var ui=SpreadsheetApp.getUi();
    var resp=ui.alert("Generate & Email All Reports",
      "All 5 Sheet 2 reports will be generated and emailed one by one in the background.\n\nProceed?",
      ui.ButtonSet.YES_NO);
    if (resp!==ui.Button.YES) return;
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction()==="_s2_runNextInQueue") ScriptApp.deleteTrigger(t);
    });
    PropertiesService.getScriptProperties().setProperty(S2_QUEUE_KEY,JSON.stringify(S2_CHAIN_FNS.slice()));
    ScriptApp.newTrigger("_s2_runNextInQueue").timeBased().at(new Date(Date.now()+30000)).create();
    ui.alert("Started!","Reports running in background. You'll receive emails as each completes.",ui.ButtonSet.OK);
  } catch(_) { console.log("manualSendAll_Sheet2 called."); }
}

function _s2_runNextInQueue() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction()==="_s2_runNextInQueue") ScriptApp.deleteTrigger(t);
  });
  var prop=PropertiesService.getScriptProperties();
  var raw=prop.getProperty(S2_QUEUE_KEY); if (!raw) return;
  var queue=JSON.parse(raw);
  if (!queue||queue.length===0) { prop.deleteProperty(S2_QUEUE_KEY); return; }
  var fnName=queue.shift();
  prop.setProperty(S2_QUEUE_KEY,JSON.stringify(queue));
  try { _sCallFn(fnName); } catch(e) { console.log("S2 queue error in "+fnName+": "+e.message); }
  if (queue.length>0) {
    ScriptApp.newTrigger("_s2_runNextInQueue").timeBased().at(new Date(Date.now()+90000)).create();
  } else {
    prop.deleteProperty(S2_QUEUE_KEY);
    try { GmailApp.sendEmail("enggplanningcellhwh@gmail.com",
      "Sheet 2: All Exception Reports Generated — "+gToday(),
      "All 5 Sheet 2 reports generated and emailed.",
      {name:"Exception Report System — Howrah Division"}); } catch(_) {}
  }
}

// ── Menu ──────────────────────────────────────────────────────────────────────
function buildExceptionMenu() {
  SpreadsheetApp.getUi()
    .createMenu("Exception Report")
    .addItem("Sand Hump Report",         "generateSandHumpReport")
    .addItem("Buffer Report",            "generateBufferReport")
    .addItem("ROB Guard Rail Report",    "generateROBReport")
    .addItem("FOB Guard Rail Report",    "generateFOBReport")
    .addItem("Bridge Guard Rail Report", "generateBridgeReport")
    .addSeparator()
    .addItem("Generate & Email All Reports","manualSendAll_Sheet2")
    .addSeparator()
    .addItem("Open Sheet 1 Daily Report","openSheet1DailyReport")
    .addToUi();
}

// NOTE: onOpen() is intentionally NOT defined here.
// photo upload tool.gs owns onOpen() for the Drive Formula menu.
// Exception Report menu is added via installable trigger (buildExceptionMenu).
function installDailyTrigger_Sheet2() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (["s2_SandHump","s2_Buffer","s2_ROB","s2_FOB","s2_Bridge","buildExceptionMenu"].indexOf(t.getHandlerFunction())>-1)
      ScriptApp.deleteTrigger(t);
  });
  var ss=SpreadsheetApp.getActive();
  ScriptApp.newTrigger("buildExceptionMenu").forSpreadsheet(ss).onOpen().create();
  ["s2_SandHump","s2_Buffer","s2_ROB","s2_FOB","s2_Bridge"].forEach(function(fn) {
    ScriptApp.newTrigger(fn).timeBased().atHour(S2_HOUR).everyDays(1).inTimezone("Asia/Kolkata").create();
  });
  try { SpreadsheetApp.getUi().alert("Sheet 2 triggers installed at "+S2_HOUR+":00 IST."); } catch(_) {}
}