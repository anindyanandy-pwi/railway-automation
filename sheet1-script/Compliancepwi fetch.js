// ============================================================
//  CompliancePWI_Fetch.gs — Sheet 1
//  Web App endpoint: receives parsed TMS data from Claude shortcut
//  and writes it into the "Compliance by PWI" tab.
//
//  DEPLOYMENT:
//  1. Deploy as Web App: Execute as Me, Anyone with link
//  2. Paste URL in _CONFIG_ tab → PWI_WEBAPP_URL
//
//  SHEET STRUCTURE (Compliance by PWI tab):
//  Row 1: Title
//  Row 2: Instructions
//  Row 3: blank
//  Row 4: "Compliance only for PT & Xing"
//  Row 5: Date blocks (newest LEFT) — each block = 7 cols
//  Row 6: Column headers repeated for each block
//  Row 7: UNIT rows (SSE/P.WAY/HWH etc. in col B)
//  Row 8+: Data rows per unit
//
//  Each 7-col date block:
//  +0: Section/station (blank — field units fill)
//  +1: Line & Location  (P&T\nLine\nXkm Ym)
//  +2: Inspection to be complied (defect text)
//  +3: Pending Since date
//  +4: Date of compliance in TMS (blank — field units fill)
//  +5: Date of actual compliance (blank — field units fill)
//  +6: Photo of attention (blank — field units fill)
// ============================================================

var PWI_TAB_NAME   = "Compliance by PWI";
var PWI_DATE_ROW   = 5;
var PWI_HDR_ROW    = 6;
var PWI_DATA_START = 7;
var PWI_UNIT_COL   = 2;
var PWI_BLOCK_COLS = 7;
var PWI_DATA_START_COL = 3;

// ── Self-contained helpers (no dependency on Sheet 1 shared functions) ────────
function _pwi_parseDate(v) {
  if (!v) return null;
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  var s = String(v).trim();
  var m;
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return new Date(parseInt(m[3]), parseInt(m[2])-1, parseInt(m[1]));
  m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (m) return new Date(parseInt(m[3]), parseInt(m[2])-1, parseInt(m[1]));
  m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2})$/);
  if (m) return new Date(2000+parseInt(m[3]), parseInt(m[2])-1, parseInt(m[1]));
  return null;
}
function _pwi_fmtDate(d) {
  if (!d) return "";
  try {
    return String(d.getDate()).padStart(2,'0')+"."+
           String(d.getMonth()+1).padStart(2,'0')+"."+
           d.getFullYear();
  } catch(_) { return String(d); }
}
function _pwi_today() { return _pwi_fmtDate(new Date()); }
function _pwi_isEmpty(v) {
  var s = String(v||"").trim().toLowerCase();
  return s===''||s==='-'||s==='--'||s==='---'||s==='na'||s==='nil'||s==='n/a';
}

// Colors
var PWI_C = {
  DATE_HEADER : "#EAD1DC",    // Pink — date row
  COL_HEADER  : "#B6D7A8",    // Green — header row
  UNIT_ROW    : "#C9DAF8",    // Light blue — unit name row
  BLANK_CELL  : "#FFFFFF",    // White — empty cells
  DUPLICATE   : "#FFE599",    // Orange-yellow — existing data found when re-fetching
  DATA_QUALITY: "#FF9900",    // Orange — data quality flag (missing date)
  NEW_DATA    : "#D9EAD3"     // Very light green — freshly written TMS data
};

// ── Web App entry point (POST) ────────────────────────────────────────────────
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);

    // Token validation
    if (!pwi_validateToken(payload.token)) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "error", message: "Invalid or expired token"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var records = payload.records; // Array of defect objects
    var result  = pwi_writeRecordsToSheet(records);

    // Write to cache for daily report integration
    try {
      var summary = _pwi_buildFetchSummary(records);
      if (typeof cm_writeSummaryToCache === 'function')
        cm_writeSummaryToCache("generateCompliancePWIReport", summary, PWI_CACHE_STALE_DAYS);
    } catch(ce) { console.log("PWI cache write error: "+ce.message); }

    return ContentService.createTextOutput(JSON.stringify({
      status: "ok",
      written: result.written,
      skipped: result.skipped,
      duplicates: result.duplicates,
      newBlocks: result.newBlocks
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    console.log("PWI doPost error: "+err.message);
    return ContentService.createTextOutput(JSON.stringify({
      status: "error", message: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Main writer ───────────────────────────────────────────────────────────────
function pwi_writeRecordsToSheet(records) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PWI_TAB_NAME);
  if (!sheet) throw new Error("Sheet '"+PWI_TAB_NAME+"' not found.");

  var result = { written:0, skipped:0, duplicates:0, newBlocks:0 };

  // Group records by inspection date
  var byDate = {};
  records.forEach(function(rec) {
    var d = rec.inspectionDate; // DD/MM/YYYY
    if (!byDate[d]) byDate[d] = [];
    byDate[d].push(rec);
  });

  // Process each date
  var sortedDates = Object.keys(byDate).sort(function(a,b) {
    return _pwi_parseDate(b) - _pwi_parseDate(a); // newest first
  });

  for (var di = 0; di < sortedDates.length; di++) {
    var dateStr  = sortedDates[di];
    var dateRecs = byDate[dateStr];

    // Find or create date block column
    var blockCol = _pwi_findDateBlock(sheet, dateStr);
    if (blockCol < 0) {
      blockCol = _pwi_createDateBlock(sheet, dateStr);
      result.newBlocks++;
    }

    // Group by TMS authority → SSE/P.WAY unit
    var byUnit = {};
    dateRecs.forEach(function(rec) {
      var unit = pwi_tmsAuthToUnit(rec.authority) || rec.authority;
      if (!byUnit[unit]) byUnit[unit] = [];
      byUnit[unit].push(rec);
    });

    // Write each unit's data
    for (var unit in byUnit) {
      var unitRecs = byUnit[unit];
      var unitResult = _pwi_writeUnitData(sheet, unit, blockCol, unitRecs, dateStr);
      result.written    += unitResult.written;
      result.skipped    += unitResult.skipped;
      result.duplicates += unitResult.duplicates;
    }
  }

  return result;
}

// ── Find existing date block column ──────────────────────────────────────────
function _pwi_findDateBlock(sheet, dateStr) {
  var lr = sheet.getLastColumn();
  if (lr < PWI_DATA_START_COL) return -1;
  var dateRow = sheet.getRange(PWI_DATE_ROW, PWI_DATA_START_COL, 1, lr - PWI_DATA_START_COL + 1).getValues()[0];
  for (var i = 0; i < dateRow.length; i++) {
    var cv = String(dateRow[i]||"").trim();
    // Normalize: DD.MM.YY → DD/MM/YYYY and compare
    if (_pwi_datesMatch(cv, dateStr)) return PWI_DATA_START_COL + i;
  }
  return -1;
}

function _pwi_datesMatch(a, b) {
  var da = _pwi_parseDate(a), db = _pwi_parseDate(b);
  if (!da || !db) return false;
  return da.getTime() === db.getTime();
}

// ── Create new date block at chronological position ───────────────────────────
function _pwi_createDateBlock(sheet, dateStr) {
  var lr = sheet.getLastColumn();
  var targetDate = _pwi_parseDate(dateStr);
  var insertCol = PWI_DATA_START_COL; // Default: leftmost

  if (lr >= PWI_DATA_START_COL) {
    var dateRow = sheet.getRange(PWI_DATE_ROW, PWI_DATA_START_COL, 1, lr - PWI_DATA_START_COL + 1).getValues()[0];
    // Find position: dates go newest → oldest left to right
    for (var i = 0; i < dateRow.length; i++) {
      var existDate = _pwi_parseDate(String(dateRow[i]||"").trim());
      if (existDate && existDate < targetDate) {
        insertCol = PWI_DATA_START_COL + i;
        break;
      }
      insertCol = PWI_DATA_START_COL + i + 1;
    }
  }

  // Insert 7 columns at insertCol
  sheet.insertColumnsBefore(insertCol, PWI_BLOCK_COLS);

  // Format date header row (row 5)
  var dateRange = sheet.getRange(PWI_DATE_ROW, insertCol, 1, PWI_BLOCK_COLS);
  dateRange.merge();
  dateRange.setValue(dateStr)
    .setBackground(PWI_C.DATE_HEADER)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // Format column headers (row 6)
  var hdrs = ["Section / station","Line & Location","Inspection to be complied",
              "Pending Since date","Date of compliance in TMS","Date of actual compliance","Photo of attention"];
  sheet.getRange(PWI_HDR_ROW, insertCol, 1, PWI_BLOCK_COLS)
    .setValues([hdrs])
    .setBackground(PWI_C.COL_HEADER)
    .setFontWeight("bold")
    .setWrap(true);

  return insertCol;
}

// ── Write data for one unit into the date block ───────────────────────────────
function _pwi_writeUnitData(sheet, unit, blockCol, recs, dateStr) {
  var result = { written:0, skipped:0, duplicates:0 };
  var unitRow = _pwi_findUnitRow(sheet, unit);
  if (unitRow < 0) {
    console.log("Unit row not found for: "+unit);
    result.skipped += recs.length;
    return result;
  }

  // Get the unit's block end (next unit row - 1)
  var blockEnd = _pwi_findUnitBlockEnd(sheet, unitRow);

  // Group records by P&T number to write sequentially
  var byPT = {};
  var ptOrder = [];
  recs.forEach(function(rec) {
    var pt = rec.ptNumber;
    if (!byPT[pt]) { byPT[pt] = []; ptOrder.push(pt); }
    byPT[pt].push(rec);
  });

  // For each P&T group, find or create rows
  for (var pi = 0; pi < ptOrder.length; pi++) {
    var pt   = ptOrder[pi];
    var ptRecs = byPT[pt];

    for (var ri = 0; ri < ptRecs.length; ri++) {
      var rec = ptRecs[ri];
      var isDupRow = _pwi_findDuplicate(sheet, unit, blockCol, rec, unitRow, blockEnd);

      if (isDupRow > 0) {
        // Existing data found — highlight orange, insert new row below
        sheet.getRange(isDupRow, blockCol, 1, PWI_BLOCK_COLS)
          .setBackground(PWI_C.DUPLICATE);
        // Insert new row
        var insertAt = isDupRow + 1;
        sheet.insertRowAfter(isDupRow);
        blockEnd++; // adjust boundary
        _pwi_writeDataRow(sheet, insertAt, blockCol, rec, pt);
        result.duplicates++;
        result.written++;
      } else {
        // Find next blank row within unit block
        var writeRow = _pwi_findNextBlankRow(sheet, blockCol, unitRow+1, blockEnd);
        if (writeRow < 0) {
          // Need to insert row at end of unit block
          sheet.insertRowAfter(blockEnd);
          writeRow = blockEnd + 1;
          blockEnd++;
        }
        _pwi_writeDataRow(sheet, writeRow, blockCol, rec, pt);
        result.written++;
      }
    }
  }

  return result;
}

// ── Write one data row ────────────────────────────────────────────────────────
function _pwi_writeDataRow(sheet, rowNum, blockCol, rec, ptNumber) {
  // Col +0: Section/station — BLANK (field units fill)
  sheet.getRange(rowNum, blockCol).setValue("").setBackground(PWI_C.BLANK_CELL);

  // Col +1: Line & Location = "P&T\nLine\nXkm Ym"
  // P&T number always shown on every row for easy identification
  var pt = (ptNumber && ptNumber.trim()) ? ptNumber.trim() : (rec.ptNumber || "");
  var lineAndLoc = pt + "\n" + rec.line + "\n" +
                   rec.locationFromKm + "km " + rec.locationFromM + "m";
  sheet.getRange(rowNum, blockCol+1)
    .setValue(lineAndLoc)
    .setBackground(PWI_C.NEW_DATA)
    .setWrap(true);

  // Col +2: Inspection to be complied (defect text only)
  sheet.getRange(rowNum, blockCol+2)
    .setValue(rec.itemNeedingAttention)
    .setBackground(PWI_C.NEW_DATA)
    .setWrap(true);

  // Col +3: Pending Since date
  sheet.getRange(rowNum, blockCol+3)
    .setValue(rec.inspectionDate)
    .setBackground(PWI_C.NEW_DATA)
    .setHorizontalAlignment("center");

  // Col +4,+5,+6: BLANK — field units fill
  sheet.getRange(rowNum, blockCol+4, 1, 3)
    .clearContent()
    .setBackground(PWI_C.BLANK_CELL);
}

// ── Duplicate detection ───────────────────────────────────────────────────────
// Key: Unit + Line&Location (starts with P&T) + Defect text + Inspection date
function _pwi_findDuplicate(sheet, unit, blockCol, rec, unitRowStart, blockEnd) {
  if (blockEnd <= unitRowStart) return -1;
  var numRows = blockEnd - unitRowStart;
  if (numRows < 1) return -1;

  var colB   = sheet.getRange(unitRowStart+1, PWI_UNIT_COL, numRows, 1).getValues();
  var colLoc = sheet.getRange(unitRowStart+1, blockCol+1, numRows, 1).getValues();
  var colDef = sheet.getRange(unitRowStart+1, blockCol+2, numRows, 1).getValues();
  var colDt  = sheet.getRange(unitRowStart+1, blockCol+3, numRows, 1).getValues();

  for (var i = 0; i < numRows; i++) {
    var rowLoc = String(colLoc[i][0]||"").trim();
    var rowDef = String(colDef[i][0]||"").trim();
    var rowDt  = String(colDt[i][0]||"").trim();

    if (!rowLoc && !rowDef) continue; // truly blank row

    var locMatch = rowLoc.indexOf(rec.ptNumber) > -1 &&
                   rowLoc.indexOf(rec.line) > -1;
    var defMatch = rowDef === rec.itemNeedingAttention.trim();
    var dtMatch  = _pwi_datesMatch(rowDt, rec.inspectionDate);

    if (locMatch && defMatch && dtMatch) return unitRowStart + 1 + i;
  }
  return -1;
}

// ── Find unit row ─────────────────────────────────────────────────────────────
function _pwi_findUnitRow(sheet, unit) {
  var lr = sheet.getLastRow();
  if (lr < PWI_DATA_START) return -1;
  var colB = sheet.getRange(PWI_DATA_START, PWI_UNIT_COL, lr-PWI_DATA_START+1, 1).getValues();
  var unitLower = unit.toLowerCase().trim();
  for (var i = 0; i < colB.length; i++) {
    var v = String(colB[i][0]||"").trim().toLowerCase();
    if (v === unitLower) return PWI_DATA_START + i;
  }
  return -1;
}

// ── Find end of unit block (row before next unit or last row) ─────────────────
function _pwi_findUnitBlockEnd(sheet, unitRow) {
  var lr = sheet.getLastRow();
  if (unitRow >= lr) return unitRow;
  var colB = sheet.getRange(unitRow+1, PWI_UNIT_COL, lr-unitRow, 1).getValues();
  for (var i = 0; i < colB.length; i++) {
    if (String(colB[i][0]||"").trim()) return unitRow + i;
  }
  return lr;
}

// ── Find next blank row in block ──────────────────────────────────────────────
function _pwi_findNextBlankRow(sheet, blockCol, startRow, endRow) {
  if (startRow > endRow) return -1;
  var numRows = endRow - startRow + 1;
  var vals = sheet.getRange(startRow, blockCol+1, numRows, 3).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (!vals[i][0] && !vals[i][1] && !vals[i][2]) return startRow + i;
  }
  return -1;
}

// ── Date parser (handles DD/MM/YYYY, DD.MM.YYYY, DD.MM.YY) ───────────────────
function _pwi_parseDate(v) {
  if (!v) return null;
  var s = String(v).trim();
  var m;
  // DD/MM/YYYY or DD-MM-YYYY
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return new Date(parseInt(m[3]), parseInt(m[2])-1, parseInt(m[1]));
  // DD.MM.YYYY
  m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (m) return new Date(parseInt(m[3]), parseInt(m[2])-1, parseInt(m[1]));
  // DD.MM.YY
  m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2})$/);
  if (m) return new Date(2000+parseInt(m[3]), parseInt(m[2])-1, parseInt(m[1]));
  return null;
}

// ── Build summary for cache ────────────────────────────────────────────────────
function _pwi_buildFetchSummary(records) {
  var byUnit = {};
  records.forEach(function(rec) {
    var unit = pwi_tmsAuthToUnit(rec.authority) || rec.authority;
    var aen  = pwi_unitToAEN(unit) || "UNASSIGNED";
    if (!byUnit[aen]) byUnit[aen] = {};
    var pt   = rec.ptNumber;
    if (!byUnit[aen][pt]) byUnit[aen][pt] = 0;
    byUnit[aen][pt]++;
  });

  var exceptions = {};
  var total = 0;
  for (var aen in byUnit) {
    var pts = Object.keys(byUnit[aen]).length;
    var cnt = Object.values(byUnit[aen]).reduce(function(a,b){ return a+b; }, 0);
    exceptions[aen] = {};
    exceptions[aen]["Defects fetched from TMS"] = cnt;
    total += cnt;
  }
  return { totalExceptions: total, exceptions: exceptions };
}

// ── Menu trigger: prepare token and instruct user ────────────────────────────
function pwiStartTMSFetch() {
  var info = pwi_prepareToken();
  if (!info.webAppUrl || info.webAppUrl === "PASTE_YOUR_DEPLOYED_WEBAPP_URL_HERE") {
    SpreadsheetApp.getUi().alert(
      "Web App URL not set.\n\n" +
      "Please:\n1. Deploy CompliancePWI_Fetch.gs as a Web App\n" +
      "2. Paste the URL in _CONFIG_ tab → PWI_WEBAPP_URL row\n" +
      "3. Run this again."
    );
    return;
  }
  // Token is already shown to user by pwi_prepareToken()
}

// ── Fallback: process JSON files uploaded to Drive folder ─────────────────────
// Run this from menu after uploading tms_records_*.json files to Drive
function pwi_processUploadedFiles() {
  var ui       = SpreadsheetApp.getUi();
  var folderId = "1K_92lYmySldpNze9DMjROrFDt-TJztH_";
  var BATCH    = 15; // records per run — safe within 6-min limit

  // ── Step 1: Load all records from JSON files in Drive ──────────────────────
  var folder;
  try { folder = DriveApp.getFolderById(folderId); }
  catch(e) { ui.alert("Could not access Drive folder:\n"+e.message); return; }

  var props         = PropertiesService.getScriptProperties();
  var allRecordsRaw = props.getProperty("PWI_IMPORT_RECORDS");
  var startIndex    = parseInt(props.getProperty("PWI_IMPORT_INDEX")||"0");
  var filesProcessed= parseInt(props.getProperty("PWI_IMPORT_FILES")||"0");
  var totalWritten  = parseInt(props.getProperty("PWI_IMPORT_WRITTEN")||"0");
  var totalDups     = parseInt(props.getProperty("PWI_IMPORT_DUPS")||"0");
  var totalBlocks   = parseInt(props.getProperty("PWI_IMPORT_BLOCKS")||"0");

  // Fresh start — load from Drive files
  if (!allRecordsRaw || startIndex === 0) {
    var allRecords = [];
    filesProcessed = 0;
    var files = folder.getFilesByType("application/json");
    while (files.hasNext()) {
      var file = files.next();
      var name = file.getName();
      if (name.indexOf("tms_records_") < 0) continue;
      try {
        var content = file.getBlob().getDataAsString();
        var data    = JSON.parse(content);
        var recs    = data.records || data;
        if (Array.isArray(recs)) {
          allRecords = allRecords.concat(recs);
          filesProcessed++;
          // Move to _processed_ subfolder
          try {
            var pf = folder.getFoldersByName("_processed_");
            var pFolder = pf.hasNext() ? pf.next() : folder.createFolder("_processed_");
            file.moveTo(pFolder);
          } catch(_) {}
        }
      } catch(pe) { console.log("File read error: "+pe.message); }
    }

    if (allRecords.length === 0) {
      ui.alert("No tms_records_*.json files found in Drive folder.\n\n"+
               "Upload file to:\nhttps://drive.google.com/drive/folders/"+folderId);
      return;
    }

    // Store in script properties for continuation
    props.setProperty("PWI_IMPORT_RECORDS",  JSON.stringify(allRecords));
    props.setProperty("PWI_IMPORT_INDEX",    "0");
    props.setProperty("PWI_IMPORT_FILES",    String(filesProcessed));
    props.setProperty("PWI_IMPORT_WRITTEN",  "0");
    props.setProperty("PWI_IMPORT_DUPS",     "0");
    props.setProperty("PWI_IMPORT_BLOCKS",   "0");
    allRecordsRaw = JSON.stringify(allRecords);
    startIndex    = 0;
    totalWritten  = 0; totalDups = 0; totalBlocks = 0;
  }

  var allRecords = JSON.parse(allRecordsRaw);
  var total      = allRecords.length;
  var endIndex   = Math.min(startIndex + BATCH, total);
  var batch      = allRecords.slice(startIndex, endIndex);

  // ── Step 2: Process this batch ─────────────────────────────────────────────
  var result = pwi_writeRecordsToSheet(batch);
  totalWritten += result.written;
  totalDups    += result.duplicates;
  totalBlocks  += result.newBlocks;

  // ── Step 3: Check if more batches remain ───────────────────────────────────
  if (endIndex < total) {
    // Save progress and schedule continuation
    props.setProperty("PWI_IMPORT_INDEX",   String(endIndex));
    props.setProperty("PWI_IMPORT_WRITTEN", String(totalWritten));
    props.setProperty("PWI_IMPORT_DUPS",    String(totalDups));
    props.setProperty("PWI_IMPORT_BLOCKS",  String(totalBlocks));

    ui.alert(
      "Batch Progress: "+endIndex+" / "+total+" records processed.\n\n"+
      "Click OK then run 'Import TMS Files from Drive' again to continue.\n"+
      "(Repeat until complete)"
    );
  } else {
    // All done — clean up
    props.deleteProperty("PWI_IMPORT_RECORDS");
    props.deleteProperty("PWI_IMPORT_INDEX");
    props.deleteProperty("PWI_IMPORT_FILES");
    props.deleteProperty("PWI_IMPORT_WRITTEN");
    props.deleteProperty("PWI_IMPORT_DUPS");
    props.deleteProperty("PWI_IMPORT_BLOCKS");

    // Update cache
    try {
      var summary = _pwi_buildFetchSummary(allRecords);
      if (typeof cm_writeSummaryToCache === 'function')
        cm_writeSummaryToCache("generateCompliancePWIReport", summary, PWI_CACHE_STALE_DAYS);
    } catch(ce) { console.log("Cache error: "+ce.message); }

    ui.alert(
      "✅ TMS Data Import Complete!\n\n"+
      "Files processed: "+filesProcessed+"\n"+
      "Total records: "+total+"\n"+
      "Rows written: "+totalWritten+"\n"+
      "Duplicates found: "+totalDups+"\n"+
      "New date blocks created: "+totalBlocks
    );
  }
}