// ============================================================
//  Sheet1_DailyScheduler.gs — Sheet 1
//  Individual report triggers (one per report, non-blocking).
//  Also handles manual "Generate & Email All Reports" button.
// ============================================================

var S1_FOLDER_ID   = "1InCD6mlBHmvuvtgp1ycTY-bNkZxwqbld";
var S1_HOUR        = 9;   // ← Change this to change schedule time
var S1_REPORT_SHEET= "Exception Report";

var S1_RECIPIENTS = [
  "adenbdcer@gmail.com","srdenchwh@gmail.com","dkaeaen@gmail.com",
  "aentrhwh@gmail.com","deepsonamoni@gmail.com","aenndae@gmail.com",
  "adenndaeminicontrol@gmail.com","maityshuvajit22@gmail.com",
  "aenbhper@gmail.com","enggctlhwh@gmail.com","aenazimganj@gmail.com",
  "aenofficial7@gmail.com","aendkaeminicontrol@gmail.com",
  "er.aenrph@gmail.com","aen1llh@gmail.com","adenazminicontrol@gmail.com",
  "aenkatwa.er@gmail.com"
];

// Add new reports here as more scripts are added
var S1_CHAIN_FNS = [
  { id:"BAD_ROAD_SURFACE", name:"Bad_Road_Surface",       fn:"s1_BadRoadSurface"  },
  { id:"TWS_TIE_BAR",      name:"TWS_Tie_Bar",            fn:"s1_TWSTieBar"       },
  { id:"GAPLESS_CMS",      name:"Gapless_Joint_CMS",      fn:"s1_GaplessCMS"      },
  { id:"GJOINT",           name:"Physically_Damaged_GJoint",fn:"s1_GJoint"        }
];

// ── Individual report runners (called by daily triggers) ──────────────────────
function s1_BadRoadSurface() {
  _sRunOne("BAD_ROAD_SURFACE","Bad_Road_Surface","generateBadRoadReport",
           S1_REPORT_SHEET,S1_FOLDER_ID,S1_RECIPIENTS);
}
function s1_TWSTieBar() {
  _sRunOne("TWS_TIE_BAR","TWS_Tie_Bar","generateTWSTieBarReport",
           S1_REPORT_SHEET,S1_FOLDER_ID,S1_RECIPIENTS);
}
function s1_GaplessCMS() {
  _sRunOne("GAPLESS_CMS","Gapless_Joint_CMS","generateGaplessReport",
           S1_REPORT_SHEET,S1_FOLDER_ID,S1_RECIPIENTS);
}
function s1_GJoint() {
  _sRunOne("GJOINT","Physically_Damaged_GJoint","generateGJointReport",
           S1_REPORT_SHEET,S1_FOLDER_ID,S1_RECIPIENTS);
}

// ── Core single-report runner ─────────────────────────────────────────────────
function _sRunOne(sectionId, reportName, fnName, reportSheet, folderId, recipients) {
  var dateStr=cm_fmtDate(new Date());
  try {
    _sCallFn(fnName);
    var tmp=_sExtract(reportSheet,sectionId);
    if (!tmp) throw new Error("Section '"+sectionId+"' not found.");
    var fname=reportName+"_Exception_"+dateStr+".pdf";
    var pdfBlob=_sToPDF(tmp,fname);
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(tmp);
    var dayFolder=_sGetDayFolder(folderId,dateStr);
    var driveFile=dayFolder.createFile(pdfBlob);
    _sEmail(recipients, reportName.replace(/_/g," "), dateStr, pdfBlob, driveFile.getUrl());
    console.log("SUCCESS: "+reportName+" — "+dateStr);
  } catch(e) {
    _sCleanTmp();
    _sErrorEmail("enggplanningcellhwh@gmail.com", reportName, dateStr, e.message);
    console.log("ERROR in "+reportName+": "+e.message);
  }
}

function _sCallFn(fnName) {
  var fn=this[fnName];
  if (typeof fn==="function") fn();
  else throw new Error("Function not found: "+fnName);
}

function _sExtract(reportSheetName, sectionId) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var rsh=ss.getSheetByName(reportSheetName); if (!rsh) return null;
  var lr=rsh.getLastRow(); if (lr<1) return null;
  var col=rsh.getRange(1,1,lr,1).getValues();
  var SM=["##START:"+sectionId+"##","##S:"+sectionId+"##"];
  var EM=["##END:"+sectionId+"##","##E:"+sectionId+"##"];
  var sr=-1,er=-1;
  for (var i=0;i<col.length;i++) {
    if (SM.indexOf(col[i][0])>-1) sr=i+1;
    if (EM.indexOf(col[i][0])>-1) er=i+1;
  }
  if (sr<0||er<0) return null;
  var cs=sr+2,ce=er-1; if (cs>ce) return null;
  var nR=ce-cs+1, nC=rsh.getLastColumn();
  var src=rsh.getRange(cs,1,nR,nC);
  var tmp=ss.insertSheet("_TMP_EXPORT_");
  tmp.setColumnWidth(1,960); tmp.setHiddenGridlines(true);
  var dst=tmp.getRange(1,1,nR,nC);
  dst.setValues(src.getValues());
  dst.setBackgrounds(src.getBackgrounds());
  dst.setFontColors(src.getFontColors());
  dst.setFontWeights(src.getFontWeights());
  if (nC>1) tmp.hideColumns(2,nC-1);
  return tmp;
}

function _sToPDF(sheet,filename) {
  var id=SpreadsheetApp.getActiveSpreadsheet().getId();
  var url="https://docs.google.com/spreadsheets/d/"+id+
    "/export?format=pdf&gid="+sheet.getSheetId()+
    "&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false"+
    "&sheetnames=false&pagenumbers=false&top_margin=0.5&bottom_margin=0.5"+
    "&left_margin=0.5&right_margin=0.5";
  var res=UrlFetchApp.fetch(url,{
    headers:{Authorization:"Bearer "+ScriptApp.getOAuthToken()},
    muteHttpExceptions:true
  });
  if (res.getResponseCode()!==200) throw new Error("PDF export HTTP "+res.getResponseCode());
  return res.getBlob().setName(filename);
}

function _sGetDayFolder(folderId,dateStr) {
  var root=DriveApp.getFolderById(folderId);
  var sub=root.getFoldersByName(dateStr);
  return sub.hasNext()?sub.next():root.createFolder(dateStr);
}

function _sEmail(recipients,reportName,dateStr,blob,driveUrl) {
  var subject="Howrah Division Exception Report — "+reportName+" — "+dateStr;
  var plain="Dear Sir/Ma'am,\n\nPlease find attached the Exception Report for "+reportName+
    " as on "+dateStr+".\n\nGenerated automatically at "+S1_HOUR+":00 AM IST.\n\n"+
    "Google Drive Link: "+driveUrl+"\n\nRegards,\nException Report System\nHowrah Division, Eastern Railway";
  var html="<p>Dear Sir/Ma'am,</p><p>Please find attached the Exception Report for <b>"+
    reportName+"</b> as on <b>"+dateStr+"</b>.</p>"+
    "<p>Generated automatically at "+S1_HOUR+":00 AM IST.</p>"+
    "<p>&#128193; <a href='"+driveUrl+"'>Open in Google Drive</a></p>"+
    "<br><p>Regards,<br><b>Exception Report System</b><br>Howrah Division, Eastern Railway</p>";
  GmailApp.sendEmail(recipients.join(","),subject,plain,{
    htmlBody:html,attachments:[blob],name:"Exception Report — Howrah Division"
  });
}

function _sErrorEmail(recipient,reportName,dateStr,errMsg) {
  try {
    GmailApp.sendEmail(recipient,
      "WARNING: Exception Report Failed — "+reportName+" — "+dateStr,
      "Report '"+reportName+"' failed on "+dateStr+".\n\nError: "+errMsg+
      "\n\nOpen Apps Script → Executions for details.");
  } catch(_) {}
}

function _sCleanTmp() {
  try {
    var t=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("_TMP_EXPORT_");
    if (t) SpreadsheetApp.getActiveSpreadsheet().deleteSheet(t);
  } catch(_) {}
}

// ── Manual "Generate & Email All Reports" — non-blocking chain ────────────────
var S1_QUEUE_KEY="S1_REPORT_QUEUE";

function manualSendAll_Sheet1() {
  try {
    var ui=SpreadsheetApp.getUi();
    var resp=ui.alert("Generate & Email All Reports",
      "All 4 reports will be generated and emailed one by one in the background.\n"+
      "Each runs automatically — no waiting required.\n\nProceed?",
      ui.ButtonSet.YES_NO);
    if (resp!==ui.Button.YES) return;
    // Schedule ALL as triggers — nothing runs inline
    _s1_startQueue(S1_CHAIN_FNS.map(function(s){return s.fn;}));
    ui.alert("Started!",
      "Reports are running in the background.\n"+
      "You will receive emails as each report completes.\n"+
      "A final completion email will be sent when all are done.",
      ui.ButtonSet.OK);
  } catch(_) { console.log("manualSendAll_Sheet1 called."); }
}

function _s1_startQueue(queue) {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction()==="_s1_runNextInQueue") ScriptApp.deleteTrigger(t);
  });
  PropertiesService.getScriptProperties().setProperty(S1_QUEUE_KEY, JSON.stringify(queue));
  ScriptApp.newTrigger("_s1_runNextInQueue")
    .timeBased().at(new Date(Date.now()+30000)).create();
}

function _s1_runNextInQueue() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction()==="_s1_runNextInQueue") ScriptApp.deleteTrigger(t);
  });
  var prop=PropertiesService.getScriptProperties();
  var raw=prop.getProperty(S1_QUEUE_KEY); if (!raw) return;
  var queue=JSON.parse(raw);
  if (!queue||queue.length===0) { prop.deleteProperty(S1_QUEUE_KEY); return; }
  var fnName=queue.shift();
  prop.setProperty(S1_QUEUE_KEY,JSON.stringify(queue));
  try { _sCallFn(fnName); } catch(e) { console.log("Queue error in "+fnName+": "+e.message); }
  if (queue.length>0) {
    ScriptApp.newTrigger("_s1_runNextInQueue")
      .timeBased().at(new Date(Date.now()+90000)).create();
  } else {
    prop.deleteProperty(S1_QUEUE_KEY);
    try {
      GmailApp.sendEmail("enggplanningcellhwh@gmail.com",
        "Sheet 1: All Exception Reports Generated — "+cm_fmtDate(new Date()),
        "All 4 Sheet 1 exception reports have been generated and emailed.",
        {name:"Exception Report System — Howrah Division"});
    } catch(_) {}
  }
}