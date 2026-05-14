// ============================================================================
// PWI at Block Site and certificates.gs
// Exception report for PWI block site certificates and timing compliance
// Integrates with Howrah Division Exception Report System
// ============================================================================

// ----- CONFIGURATION -----
var PBS_SHEET_NAME = "PWI at Block Site and certificates";
var PBS_HEADER_ROW = 3;
var PBS_DATE_ROW = 2;
var PBS_DATA_START_ROW = 4;
var PBS_FIRST_BLOCK_COL = 18; // Column R = 18 (0-indexed)
var PBS_BLOCK_WIDTH = 14; // Each date block spans 14 columns
var PBS_MIN_PRE_BUFFER = 45; // Minutes before block start
var PBS_MIN_POST_BUFFER = 45; // Minutes after block end

// Column offsets within each date block (relative to block start col)
var PBS_COL = {
  BLOCK_SECTION: 0,   // R
  LINE: 1,            // S
  PURPOSE: 2,         // T
  PWI_NAME: 3,        // U
  BLOCK_TIMING: 4,    // V
  PRE_TIME: 5,        // W
  PRE_IMAGE: 6,       // X
  CERT_SAFETY: 7,     // Y
  CERT_AGENCY: 8,     // Z
  CERT_WORK: 9,       // AA
  VIDEO: 10,          // AB
  POST_TIME: 11,      // AC
  POST_IMAGE: 12,     // AD
  REMARKS: 13         // AE
};

// ----- HELPER: Check if cell is placeholder/empty -----
function _pbsIsEmpty(val) {
  if (!val) return true;
  if (val === null || val === undefined) return true;
  var s = String(val).trim();
  if (!s) return true;
  if (/^-+$/.test(s)) return true;
  if (/^(nil|na|n\/a)$/i.test(s)) return true;
  return false;
}

// ----- HELPER: Check if entire row is placeholder -----
function _pbsIsPlaceholderRow(rowData, blockStartCol) {
  for (var i = 0; i < PBS_BLOCK_WIDTH; i++) {
    var val = rowData[blockStartCol + i];
    if (!_pbsIsEmpty(val)) return false;
  }
  return true;
}

// ----- HELPER: Parse date from header -----
function _pbsParseHeaderDate(headerText) {
  if (!headerText) return null;
  var s = String(headerText).trim();
  var m1 = s.match(/&\s*(\d{1,2}\/\d{1,2}\/\d{2,4})/);
  if (m1) return m1[1];
  var m2 = s.match(/(\d{1,2})-(\d{1,2}\/\d{1,2}\/\d{2,4})/);
  if (m2) return m2[2];
  var m3 = s.match(/(\d{1,2}\/\d{1,2}\/\d{2,4})/);
  if (m3) return m3[1];
  return null;
}

// ----- HELPER: Format date as DD/MM/YY for comparison -----
function _pbsFormatDate(date) {
  var d = String(date.getDate()).padStart(2, '0');
  var m = String(date.getMonth() + 1).padStart(2, '0');
  var y = String(date.getFullYear()).slice(-2);
  return d + "/" + m + "/" + y;
}

// ----- HELPER: Parse DD/MM/YY or DD/MM/YYYY to Date object -----
function _pbsParseDate(dateStr) {
  if (!dateStr) return null;
  var parts = dateStr.split('/');
  if (parts.length !== 3) return null;
  var day = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10) - 1;
  var year = parseInt(parts[2], 10);
  if (year < 100) year += 2000;
  return new Date(year, month, day);
}

// ----- HELPER: Parse time HH:MM or H:MM to minutes since midnight -----
function _pbsParseTime(timeStr) {
  if (!timeStr) return null;
  var s = String(timeStr).trim();
  if (s.indexOf('.') >= 0) return null;
  var parts = s.split(':');
  if (parts.length !== 2) return null;
  var h = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10);
  if (isNaN(h) || isNaN(m) || h < 0 || h > 23 || m < 0 || m > 59) return null;
  return h * 60 + m;
}

// ----- HELPER: Parse block timing "HH:MM-HH:MM" -----
function _pbsParseBlockTiming(timingStr) {
  if (!timingStr) return null;
  var s = String(timingStr).trim();
  var parts = s.split('-');
  if (parts.length !== 2) return null;
  var start = _pbsParseTime(parts[0]);
  var end = _pbsParseTime(parts[1]);
  if (start === null || end === null) return null;
  return { start: start, end: end };
}

// ----- HELPER: Calculate time difference handling midnight crossing -----
function _pbsTimeDiff(time1, time2) {
  if (time1 > time2) {
    return (1440 - time1) + time2;
  } else {
    return time2 - time1;
  }
}

// ----- HELPER: Check if cell has photo -----
function _pbsHasPhoto(val, formula) {
  if (!val && !formula) return false;
  if (val !== null && typeof val === "object" && !Array.isArray(val)) return true;
  var s = String(val).trim();
  if (s === "CellImage") return true;
  var l = s.toLowerCase();
  if (l.indexOf("drive.google.com") > -1) return true;
  if (/\.(jpg|jpeg|png|gif)/i.test(s)) return true;
  if (formula) {
    var f = String(formula).toLowerCase();
    if (f.indexOf("image(") > -1 || f.indexOf("hyperlink(") > -1) return true;
  }
  return false;
}

// ----- HELPER: Check if cell has valid video link -----
function _pbsHasVideoLink(val) {
  if (!val) return false;
  var s = String(val).trim().toLowerCase();
  return s.indexOf("drive.google.com") > -1 && (s.indexOf("/file/") > -1 || s.indexOf("/open?id=") > -1);
}

// ----- MAIN FUNCTION: Generate PWI Block Report (Yesterday Only) -----
function generatePWIBlockReport() {
  var ss = SpreadsheetApp.openById("1-EWkp8nq5aL_BKs3MQymexS-1gjrrXtPbLu1ZCZ1_FI");
  var sheet = ss.getSheetByName(PBS_SHEET_NAME);
  if (!sheet) throw new Error("Sheet '" + PBS_SHEET_NAME + "' not found");
  
  var today = new Date();
  var yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);
  var targetDate = _pbsFormatDate(yesterday);
  Logger.log("=== PWI Block Site Report (Yesterday: " + targetDate + ") ===");
  
  var dateRow = sheet.getRange(PBS_DATE_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  var blockCol = -1;
  
  for (var c = PBS_FIRST_BLOCK_COL; c < dateRow.length; c += PBS_BLOCK_WIDTH) {
    var headerDate = _pbsParseHeaderDate(dateRow[c]);
    if (headerDate) {
      var parsedDate = _pbsParseDate(headerDate);
      if (parsedDate && _pbsFormatDate(parsedDate) === targetDate) {
        blockCol = c;
        break;
      }
    }
  }
  
  if (blockCol === -1) {
    Logger.log("No date block found for " + targetDate);
    return { totalExceptions: 0, exceptions: {} };
  }
  
  Logger.log("Found date block at column " + (blockCol + 1));
  
  var lastRow = sheet.getLastRow();
  var allData = sheet.getRange(PBS_DATA_START_ROW, 1, lastRow - PBS_DATA_START_ROW + 1, sheet.getLastColumn()).getValues();
  var allFormulas = sheet.getRange(PBS_DATA_START_ROW, 1, lastRow - PBS_DATA_START_ROW + 1, sheet.getLastColumn()).getFormulas();
  
  var exceptions = {};
  var lastAEN = "";
  var lastPWI = "";
  
  for (var r = 0; r < allData.length; r++) {
    var row = allData[r];
    var formulas = allFormulas[r];
    var rowNum = PBS_DATA_START_ROW + r;
    
    var aenRaw = row[1];
    var pwiRaw = row[2];
    if (!_pbsIsEmpty(aenRaw)) lastAEN = String(aenRaw).trim();
    if (!_pbsIsEmpty(pwiRaw)) lastPWI = String(pwiRaw).trim();
    
    var aen = _pbsNormalizeAEN(lastAEN);
    if (!aen) continue;
    if (_pbsIsPlaceholderRow(row, blockCol)) continue;
    
    var blockSection = row[blockCol + PBS_COL.BLOCK_SECTION];
    var line = row[blockCol + PBS_COL.LINE];
    var purpose = row[blockCol + PBS_COL.PURPOSE];
    var pwiName = row[blockCol + PBS_COL.PWI_NAME];
    var blockTiming = row[blockCol + PBS_COL.BLOCK_TIMING];
    var preTime = row[blockCol + PBS_COL.PRE_TIME];
    var preImage = row[blockCol + PBS_COL.PRE_IMAGE];
    var preImageFormula = formulas[blockCol + PBS_COL.PRE_IMAGE];
    var certSafety = row[blockCol + PBS_COL.CERT_SAFETY];
    var certSafetyFormula = formulas[blockCol + PBS_COL.CERT_SAFETY];
    var certAgency = row[blockCol + PBS_COL.CERT_AGENCY];
    var certAgencyFormula = formulas[blockCol + PBS_COL.CERT_AGENCY];
    var certWork = row[blockCol + PBS_COL.CERT_WORK];
    var certWorkFormula = formulas[blockCol + PBS_COL.CERT_WORK];
    var video = row[blockCol + PBS_COL.VIDEO];
    var postTime = row[blockCol + PBS_COL.POST_TIME];
    var postImage = row[blockCol + PBS_COL.POST_IMAGE];
    var postImageFormula = formulas[blockCol + PBS_COL.POST_IMAGE];
    var remarks = row[blockCol + PBS_COL.REMARKS];
    
    var rowExceptions = [];
    
    if (_pbsIsEmpty(blockSection)) rowExceptions.push("Block section missing");
    if (_pbsIsEmpty(line)) rowExceptions.push("LINE missing");
    if (_pbsIsEmpty(purpose)) rowExceptions.push("Purpose missing");
    if (_pbsIsEmpty(pwiName)) rowExceptions.push("PWI name missing");
    if (_pbsIsEmpty(blockTiming)) rowExceptions.push("Block timing missing");
    
    var blockParsed = _pbsParseBlockTiming(blockTiming);
    if (_pbsIsEmpty(preTime)) {
      rowExceptions.push("Pre-photo time missing");
    } else {
      var preTimeParsed = _pbsParseTime(preTime);
      if (preTimeParsed === null) {
        rowExceptions.push("Pre-photo time invalid format");
      } else if (blockParsed) {
        var diff = _pbsTimeDiff(preTimeParsed, blockParsed.start);
        if (diff < PBS_MIN_PRE_BUFFER) {
          rowExceptions.push("Pre-photo < 45min before block (" + diff + "min)");
        }
      }
    }
    
    if (!_pbsHasPhoto(preImage, preImageFormula)) rowExceptions.push("Pre-photo image missing");
    if (!_pbsHasPhoto(certSafety, certSafetyFormula)) rowExceptions.push("Safety certificate missing");
    if (!_pbsHasPhoto(certAgency, certAgencyFormula)) rowExceptions.push("Agency certificate missing");
    if (!_pbsHasPhoto(certWork, certWorkFormula)) rowExceptions.push("Work certificate missing");
    if (!_pbsHasVideoLink(video)) rowExceptions.push("Video link missing");
    
    if (_pbsIsEmpty(postTime)) {
      rowExceptions.push("Post-photo time missing");
    } else {
      var postTimeParsed = _pbsParseTime(postTime);
      if (postTimeParsed === null) {
        rowExceptions.push("Post-photo time invalid format");
      } else if (blockParsed) {
        var diff2 = postTimeParsed - blockParsed.end;
        if (diff2 < 0) diff2 += 1440;
        if (diff2 < PBS_MIN_POST_BUFFER) {
          rowExceptions.push("Post-photo < 45min after block (" + diff2 + "min)");
        }
      }
    }
    
    if (!_pbsHasPhoto(postImage, postImageFormula)) rowExceptions.push("Post-photo image missing");
    if (_pbsIsEmpty(remarks)) rowExceptions.push("Remarks missing");
    
    if (rowExceptions.length > 0) {
      if (!exceptions[aen]) exceptions[aen] = {};
      for (var e = 0; e < rowExceptions.length; e++) {
        var exType = rowExceptions[e];
        if (!exceptions[aen][exType]) exceptions[aen][exType] = 0;
        exceptions[aen][exType]++;
      }
      Logger.log("Row " + rowNum + " (" + aen + " | " + lastPWI + "): " + rowExceptions.join(", "));
    }
  }
  
  var totalEx = 0;
  for (var a in exceptions) {
    for (var t in exceptions[a]) {
      totalEx += exceptions[a][t];
    }
  }
  
  Logger.log("Total exceptions: " + totalEx);
  _pbsWriteCache("generatePWIBlockReport", { totalExceptions: totalEx, exceptions: exceptions });
  return { totalExceptions: totalEx, exceptions: exceptions };
}

// ----- WEEKLY FUNCTION: Generate PWI Block Report (Last 7 Days) -----
function generatePWIBlockWeeklyReport() {
  var ss = SpreadsheetApp.openById("1-EWkp8nq5aL_BKs3MQymexS-1gjrrXtPbLu1ZCZ1_FI");
  var sheet = ss.getSheetByName(PBS_SHEET_NAME);
  if (!sheet) throw new Error("Sheet '" + PBS_SHEET_NAME + "' not found");
  
  var today = new Date();
  var weekAgo = new Date(today);
  weekAgo.setDate(weekAgo.getDate() - 7);
  Logger.log("=== PWI Block Site Weekly Report (Last 7 Days) ===");
  
  var dateRow = sheet.getRange(PBS_DATE_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  var blockCols = [];
  
  for (var c = PBS_FIRST_BLOCK_COL; c < dateRow.length; c += PBS_BLOCK_WIDTH) {
    var headerDate = _pbsParseHeaderDate(dateRow[c]);
    if (headerDate) {
      var parsedDate = _pbsParseDate(headerDate);
      if (parsedDate && parsedDate >= weekAgo && parsedDate < today) {
        blockCols.push({ col: c, date: _pbsFormatDate(parsedDate) });
      }
    }
  }
  
  if (blockCols.length === 0) {
    Logger.log("No date blocks found in last 7 days");
    return { totalExceptions: 0, exceptions: {} };
  }
  
  Logger.log("Found " + blockCols.length + " date blocks");
  
  var allExceptions = {};
  for (var b = 0; b < blockCols.length; b++) {
    var blockCol = blockCols[b].col;
    var blockDate = blockCols[b].date;
    Logger.log("Processing block: " + blockDate);
    
    var lastRow = sheet.getLastRow();
    var allData = sheet.getRange(PBS_DATA_START_ROW, 1, lastRow - PBS_DATA_START_ROW + 1, sheet.getLastColumn()).getValues();
    var allFormulas = sheet.getRange(PBS_DATA_START_ROW, 1, lastRow - PBS_DATA_START_ROW + 1, sheet.getLastColumn()).getFormulas();
    
    var lastAEN = "";
    var lastPWI = "";
    
    for (var r = 0; r < allData.length; r++) {
      var row = allData[r];
      var formulas = allFormulas[r];
      
      var aenRaw = row[1];
      var pwiRaw = row[2];
      if (!_pbsIsEmpty(aenRaw)) lastAEN = String(aenRaw).trim();
      if (!_pbsIsEmpty(pwiRaw)) lastPWI = String(pwiRaw).trim();
      
      var aen = _pbsNormalizeAEN(lastAEN);
      if (!aen) continue;
      if (_pbsIsPlaceholderRow(row, blockCol)) continue;
      
      var blockTiming = row[blockCol + PBS_COL.BLOCK_TIMING];
      var preTime = row[blockCol + PBS_COL.PRE_TIME];
      var postTime = row[blockCol + PBS_COL.POST_TIME];
      
      var rowExceptions = [];
      
      if (_pbsIsEmpty(row[blockCol + PBS_COL.BLOCK_SECTION])) rowExceptions.push("Block section missing");
      if (_pbsIsEmpty(row[blockCol + PBS_COL.LINE])) rowExceptions.push("LINE missing");
      if (_pbsIsEmpty(row[blockCol + PBS_COL.PURPOSE])) rowExceptions.push("Purpose missing");
      if (_pbsIsEmpty(row[blockCol + PBS_COL.PWI_NAME])) rowExceptions.push("PWI name missing");
      if (_pbsIsEmpty(blockTiming)) rowExceptions.push("Block timing missing");
      
      var blockParsed = _pbsParseBlockTiming(blockTiming);
      if (_pbsIsEmpty(preTime)) {
        rowExceptions.push("Pre-photo time missing");
      } else {
        var preTimeParsed = _pbsParseTime(preTime);
        if (preTimeParsed === null) {
          rowExceptions.push("Pre-photo time invalid format");
        } else if (blockParsed) {
          var diff = _pbsTimeDiff(preTimeParsed, blockParsed.start);
          if (diff < PBS_MIN_PRE_BUFFER) {
            rowExceptions.push("Pre-photo < 45min before block (" + diff + "min)");
          }
        }
      }
      
      if (!_pbsHasPhoto(row[blockCol + PBS_COL.PRE_IMAGE], formulas[blockCol + PBS_COL.PRE_IMAGE])) {
        rowExceptions.push("Pre-photo image missing");
      }
      if (!_pbsHasPhoto(row[blockCol + PBS_COL.CERT_SAFETY], formulas[blockCol + PBS_COL.CERT_SAFETY])) {
        rowExceptions.push("Safety certificate missing");
      }
      if (!_pbsHasPhoto(row[blockCol + PBS_COL.CERT_AGENCY], formulas[blockCol + PBS_COL.CERT_AGENCY])) {
        rowExceptions.push("Agency certificate missing");
      }
      if (!_pbsHasPhoto(row[blockCol + PBS_COL.CERT_WORK], formulas[blockCol + PBS_COL.CERT_WORK])) {
        rowExceptions.push("Work certificate missing");
      }
      if (!_pbsHasVideoLink(row[blockCol + PBS_COL.VIDEO])) {
        rowExceptions.push("Video link missing");
      }
      
      if (_pbsIsEmpty(postTime)) {
        rowExceptions.push("Post-photo time missing");
      } else {
        var postTimeParsed = _pbsParseTime(postTime);
        if (postTimeParsed === null) {
          rowExceptions.push("Post-photo time invalid format");
        } else if (blockParsed) {
          var diff2 = postTimeParsed - blockParsed.end;
          if (diff2 < 0) diff2 += 1440;
          if (diff2 < PBS_MIN_POST_BUFFER) {
            rowExceptions.push("Post-photo < 45min after block (" + diff2 + "min)");
          }
        }
      }
      
      if (!_pbsHasPhoto(row[blockCol + PBS_COL.POST_IMAGE], formulas[blockCol + PBS_COL.POST_IMAGE])) {
        rowExceptions.push("Post-photo image missing");
      }
      if (_pbsIsEmpty(row[blockCol + PBS_COL.REMARKS])) {
        rowExceptions.push("Remarks missing");
      }
      
      if (rowExceptions.length > 0) {
        if (!allExceptions[aen]) allExceptions[aen] = {};
        for (var e = 0; e < rowExceptions.length; e++) {
          var exType = rowExceptions[e];
          if (!allExceptions[aen][exType]) allExceptions[aen][exType] = 0;
          allExceptions[aen][exType]++;
        }
      }
    }
  }
  
  var totalEx = 0;
  for (var a in allExceptions) {
    for (var t in allExceptions[a]) {
      totalEx += allExceptions[a][t];
    }
  }
  
  Logger.log("Total weekly exceptions: " + totalEx);
  return { totalExceptions: totalEx, exceptions: allExceptions };
}

// ----- HELPER: Normalize AEN -----
function _pbsNormalizeAEN(raw) {
  if (!raw) return null;
  var s = String(raw).trim().toUpperCase();
  s = s.replace(/\s*\([^)]*\)/g, '');
  s = s.replace(/\s+/g, '/');
  if (s.indexOf('AEN/') !== 0) s = 'AEN/' + s;
  var aenMap = {
    'AEN/TR/HWH': 'AEN/TR/HWH',
    'AEN/LLH': 'AEN/LLH',
    'AEN/BDC': 'AEN/BDC',
    'AEN/DKAE': 'AEN/DKAE',
    'AEN/BWN': 'AEN/BWN',
    'AEN/BHP': 'AEN/BHP',
    'AEN/RPH': 'AEN/RPH',
    'AEN/NDAE': 'AEN/NDAE',
    'AEN/KWAE': 'AEN/KWAE',
    'AEN/AZ': 'AEN/AZ',
    'AEN/HWH/BMG': 'AEN/TR/HWH',
    'AEN/TR-HWH': 'AEN/TR/HWH'
  };
  return aenMap[s] || s;
}

// ----- HELPER: Write to cache -----
function _pbsWriteCache(scriptFn, summary) {
  var ss = SpreadsheetApp.openById("1-EWkp8nq5aL_BKs3MQymexS-1gjrrXtPbLu1ZCZ1_FI");
  var cacheSheet = ss.getSheetByName("_EXCEPTION_CACHE_");
  if (!cacheSheet) {
    Logger.log("Cache sheet not found");
    return;
  }
  
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy");
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm:ss");
  
  var data = cacheSheet.getDataRange().getValues();
  var rowToUpdate = -1;
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === scriptFn && data[i][1] === today) {
      rowToUpdate = i + 1;
      break;
    }
  }
  
  var slNo = 1;
  var jsonStr = JSON.stringify(summary);
  
  if (rowToUpdate > 0) {
    cacheSheet.getRange(rowToUpdate, 1, 1, 6).setValues([[scriptFn, today, slNo, jsonStr, 1, timestamp]]);
    Logger.log("Updated cache row " + rowToUpdate);
  } else {
    cacheSheet.appendRow([scriptFn, today, slNo, jsonStr, 1, timestamp]);
    Logger.log("Added new cache row");
  }
}

// ----- TEST FUNCTION -----
function testPWIBlockReport() {
  generatePWIBlockReport();
}

function testPWIBlockWeeklyReport() {
  generatePWIBlockWeeklyReport();
}
