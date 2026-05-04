// ============================================================
//  CompliancePWI_Config.gs — Sheet 1
//  One-time migration: adds PWI_UNIT_MAP section to _CONFIG_ tab
//  and updates the SOP tab.
//
// Safe fallback if CacheManager.gs constants not loaded yet
var _PWI_CONFIG_TAB = (typeof CM_TAB_CONFIG !== 'undefined') ? CM_TAB_CONFIG : "_CONFIG_";
//  Run ONCE on existing install: addPWIUnitMapSection()
//  Fresh install: handled by setupAllSystemTabs() which calls this
// ============================================================

// Section header used in _CONFIG_ tab
var PWI_UNIT_MAP_SECTION = "PWI_UNIT_MAP";

// Stale threshold for PWI compliance cache (weekly report = 8 days)
var PWI_CACHE_STALE_DAYS = 8;

// TMS section code map (name → numeric code, km range)
// Used by Chrome shortcut to make DWR calls
var PWI_TMS_SECTIONS = [
  { name: "ADL-DKAE",                  code: "1334", kmFrom: "7",   kmTo: "14"  },
  { name: "AZ-NFK",                    code: "1348", kmFrom: "177", kmTo: "181" },
  { name: "AZ-NHT",                    code: "1326", kmFrom: "2",   kmTo: "45"  },
  { name: "Bally-Bandel",              code: "1322", kmFrom: "8",   kmTo: "39"  },
  { name: "BDC-KWAE Jn",               code: "1324", kmFrom: "0",   kmTo: "104" },
  { name: "BDC-SKG",                   code: "1323", kmFrom: "39",  kmTo: "95"  },
  { name: "BLY-SKG",                   code: "1319", kmFrom: "8",   kmTo: "95"  },
  { name: "BMGA-SNT",                  code: "0941", kmFrom: "70",  kmTo: "72"  },
  { name: "BWN-KAN",                   code: "1336", kmFrom: "106", kmTo: "119" },
  { name: "BWN-KWAE",                  code: "1331", kmFrom: "0",   kmTo: "52"  },
  { name: "BZL - RCD",                 code: "1337", kmFrom: "0",   kmTo: "0"   },
  { name: "DDJ-DKAE",                  code: "1306", kmFrom: "9",   kmTo: "14"  },
  { name: "Howrah - SRC",              code: "0142", kmFrom: "0",   kmTo: "1"   },
  { name: "HWH-BLY",                   code: "1321", kmFrom: "0",   kmTo: "8"   },
  { name: "HWH-BLY Chord line",        code: "1461", kmFrom: "0",   kmTo: "8"   },
  { name: "KAN-SNT",                   code: "1327", kmFrom: "0",   kmTo: "70"  },
  { name: "KAN-UDL",                   code: "0929", kmFrom: "119", kmTo: "120" },
  { name: "KWAE-AMP",                  code: "1330", kmFrom: "0",   kmTo: "50"  },
  { name: "KWAE-AZ",                   code: "1325", kmFrom: "104", kmTo: "177" },
  { name: "LLH-BRMH",                  code: "1333", kmFrom: "4",   kmTo: "7"   },
  { name: "Magra-Tribeni Avoiding line", code: "1447", kmFrom: "0", kmTo: "3"   },
  { name: "NH-BDC",                    code: "1298", kmFrom: "2",   kmTo: "8"   },
  { name: "NHT-BDAG Link",             code: "1329", kmFrom: "112", kmTo: "172" },
  { name: "RCD-BTNG",                  code: "1335", kmFrom: "9",   kmTo: "13"  },
  { name: "RPH - DUMK",                code: "1476", kmFrom: "100", kmTo: "162" },
  { name: "SHE-VSU",                   code: "1332", kmFrom: "0",   kmTo: "117" },
  { name: "SKG-BWN",                   code: "1320", kmFrom: "95",  kmTo: "106" },
  { name: "SNT-NHT",                   code: "1328", kmFrom: "70",  kmTo: "112" }
];

// ── Migration function ────────────────────────────────────────────────────────
function addPWIUnitMapSection() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName(_PWI_CONFIG_TAB);
  if (!cfg) {
    SpreadsheetApp.getUi().alert("_CONFIG_ tab not found. Run setupAllSystemTabs() first.");
    return;
  }

  // Check if already added
  var vals = cfg.getDataRange().getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === PWI_UNIT_MAP_SECTION) {
      SpreadsheetApp.getUi().alert("PWI_UNIT_MAP section already exists in _CONFIG_. No changes made.");
      return;
    }
  }

  // Append after last row with a blank spacer
  var lr = cfg.getLastRow();
  cfg.getRange(lr+1, 1).setValue("");

  // Section header
  var hdrRow = lr + 2;
  cfg.getRange(hdrRow, 1).setValue(PWI_UNIT_MAP_SECTION)
     .setBackground("#B6D7A8").setFontWeight("bold");
  cfg.getRange(hdrRow, 2).setValue("SR DEN / AEN level")
     .setBackground("#B6D7A8").setFontWeight("bold");
  cfg.getRange(hdrRow, 3).setValue("SSE/P.WAY Unit")
     .setBackground("#B6D7A8").setFontWeight("bold");
  cfg.getRange(hdrRow, 4).setValue("TMS Auth 1 (SSE level)")
     .setBackground("#B6D7A8").setFontWeight("bold");
  cfg.getRange(hdrRow, 5).setValue("TMS Auth 2 (JE level)")
     .setBackground("#B6D7A8").setFontWeight("bold");
  cfg.getRange(hdrRow, 6).setValue("TMS Auth 3")
     .setBackground("#B6D7A8").setFontWeight("bold");
  cfg.getRange(hdrRow, 7).setValue("TMS Auth 4")
     .setBackground("#B6D7A8").setFontWeight("bold");
  cfg.getRange(hdrRow, 8).setValue("TMS Auth 5 (add more cols as needed)")
     .setBackground("#B6D7A8").setFontWeight("bold");

  // Column header row
  var colHdr = hdrRow + 1;
  cfg.getRange(colHdr, 1).setValue("SR_DEN").setFontStyle("italic");
  cfg.getRange(colHdr, 2).setValue("AEN").setFontStyle("italic");
  cfg.getRange(colHdr, 3).setValue("SSE_PWAY_UNIT").setFontStyle("italic");
  cfg.getRange(colHdr, 4, 1, 5).setValues([["TMS_AUTH_1","TMS_AUTH_2","TMS_AUTH_3","TMS_AUTH_4","TMS_AUTH_5"]]).setFontStyle("italic");

  // Sample rows (fill in progressively)
  var sample = [
    ["SR DEN HQ","AEN/TR/HWH","SSE/P.WAY/HWH","SSE/PW/1/HWH","JE/PW/2/HWH","","",""],
    ["SR DEN HQ","AEN/TR/HWH","SSE/P.WAY/BMG","JE/PW/1/BMG","SSE/PW/YD/BMG","","",""],
    ["SR DEN I/HWH","AEN/LLH","SSE/P.WAY/LLH","SSE/PW/LLH","JE/PW/SRP","SSE/PW/WS","",""],
    ["SR DEN I/HWH","AEN/LLH","SSE/P.WAY/HGY","SSE/PW/BHR","JE/PW/NKL","","",""],
    ["SR DEN I/HWH","AEN/LLH","SSE/P.WAY/TAK","JE/PW/TAK","SSE/PW/AMBG","JE/P.Way/VSU","",""],
    ["SR DEN I/HWH","AEN/BDC","SSE/P.WAY/BDC","SSE/PW/ML/BDC","SSE/PW/YD/BDC","","",""],
    ["SR DEN I/HWH","AEN/BDC","SSE/P.WAY/PDA","JE/PW/PDA","JE/PW/MYM","","",""],
    ["SR DEN II/HWH","AEN/DKAE","SSE/P.WAY/DKAE","SSE/PW/DKAE (Sectional)","SSE/PW/2/RCD","","",""],
    ["SR DEN II/HWH","AEN/DKAE","SSE/P.WAY/KQU","JE/PW/JOX","JE/PW/KQU/Sec","","",""],
    ["SR DEN II/HWH","AEN/DKAE","SSE/P.WAY/GRAE","JE/PW/BMAE (Sec.)","SSE/PW/GRAE(Sec.)","SSE/PW/JRAE(Sec.)","",""],
    ["SR DEN II/HWH","AEN/BWN","SSE/P.WAY/BWN(E)","JE/PW/SKG","SSE/PW/GRP","","",""],
    ["SR DEN II/HWH","AEN/BWN","SSE/P.WAY/BWN(W)","JE/YD/BWN","JE/PW/TIT","JE/PW/KAN","SSE/PW/BGNA",""],
    ["SR DEN IV/HWH","AEN/BHP","SSE/P.WAY/BHP","JE/PW/GKH","JE/PW/BHP","","",""],
    ["SR DEN IV/HWH","AEN/BHP","SSE/P.Way/SNT","SSE/PW/SNT(Sec)","","","",""],
    ["SR DEN IV/HWH","AEN/BHP","SSE/P.way/AMP","","","","",""],
    ["SR DEN IV/HWH","AEN/RPH","SSE/P.WAY/RPH","JE/PW/RPH","JE/PW/DUMK at PRGR","JE/P.Way/Dumka at SKIP","",""],
    ["SR DEN IV/HWH","AEN/RPH","SSE/P.WAY/NHT","JE/PW/NHT","JE/PW/MRR","","",""],
    ["SR DEN IV/HWH","AEN/RPH","SSE/P.WAY/PKR","JE/PW/1/PKR","JE/PW/2/PKR","","",""],
    ["SR DEN III/HWH","AEN/KWAE","SSE/P.WAY/KWAE(ML)","JE/PW/ML/KWAE","JE/PW/SALE","","",""],
    ["SR DEN III/HWH","AEN/KWAE","SSE/P.WAY/KWAE(BL)","JE/PW/KNHR","JE/PW/KCY","","",""],
    ["SR DEN III/HWH","AEN/AZ","SSE/P.WAY/AZ","JE/PW/AZ","JE/PW/MGAE","","",""],
    ["SR DEN III/HWH","AEN/AZ","SSE/P.WAY/BZLE","JE/PW/BZLE","SSE/PW/KGLE","","",""],
    ["SR DEN III/HWH","AEN/NDAE","SSE/P.WAY/NDAE","JE/PW/NDAE","SSE/PW/PTAE","","",""],
    ["SR DEN III/HWH","AEN/NDAE","SSE/P.Way/TBAE","SSE/PW/TBAE","JE/PW/ABKA","","",""]
  ];

  var dataStartRow = colHdr + 1;
  cfg.getRange(dataStartRow, 1, sample.length, 8).setValues(sample);

  // PWI_EXCLUDE_BEFORE setting
  var exclRow = dataStartRow + sample.length + 1;
  cfg.getRange(exclRow, 1).setValue("PWI_EXCLUDE_BEFORE")
     .setBackground("#FFF2CC").setFontWeight("bold");
  cfg.getRange(exclRow, 2).setValue("01.04.2026")
     .setBackground("#FFF2CC");
  cfg.getRange(exclRow, 3).setValue("← Data before this date is ignored by exception scanner")
     .setFontStyle("italic").setFontColor("#666666");

  // PWI_EXCEPTION_THRESHOLDS
  var thrRow = exclRow + 1;
  cfg.getRange(thrRow, 1).setValue("PWI_PENDING_DAYS")
     .setBackground("#FFF2CC").setFontWeight("bold");
  cfg.getRange(thrRow, 2).setValue("15")
     .setBackground("#FFF2CC");
  cfg.getRange(thrRow, 3).setValue("← Days after inspection date before 'TMS not entered' exception fires")
     .setFontStyle("italic").setFontColor("#666666");

  var thrRow2 = thrRow + 1;
  cfg.getRange(thrRow2, 1).setValue("PWI_COMPLIANCE_DAYS")
     .setBackground("#FFF2CC").setFontWeight("bold");
  cfg.getRange(thrRow2, 2).setValue("10")
     .setBackground("#FFF2CC");
  cfg.getRange(thrRow2, 3).setValue("← Days after TMS compliance date before 'actual work not done' exception fires (from day 11)")
     .setFontStyle("italic").setFontColor("#666666");

  var thrRow3 = thrRow2 + 1;
  cfg.getRange(thrRow3, 1).setValue("PWI_ACTUAL_VS_TMS_GRACE")
     .setBackground("#FFF2CC").setFontWeight("bold");
  cfg.getRange(thrRow3, 2).setValue("3")
     .setBackground("#FFF2CC");
  cfg.getRange(thrRow3, 3).setValue("← Grace days: actual date filled but TMS blank (allows 3 days before exception)")
     .setFontStyle("italic").setFontColor("#666666");

  // PWI Web App URL placeholder
  var urlRow = thrRow3 + 2;
  cfg.getRange(urlRow, 1).setValue("PWI_WEBAPP_URL")
     .setBackground("#CFE2F3").setFontWeight("bold");
  cfg.getRange(urlRow, 2).setValue("PASTE_YOUR_DEPLOYED_WEBAPP_URL_HERE")
     .setBackground("#CFE2F3");
  cfg.getRange(urlRow, 3).setValue("← Deploy CompliancePWI_Fetch.gs as Web App, paste URL here")
     .setFontStyle("italic").setFontColor("#666666");

  // TMS Sections config (for Chrome shortcut reference)
  var secHdr = urlRow + 2;
  cfg.getRange(secHdr, 1).setValue("PWI_TMS_SECTIONS")
     .setBackground("#EAD1DC").setFontWeight("bold");
  cfg.getRange(secHdr, 2).setValue("Section Name")
     .setBackground("#EAD1DC").setFontWeight("bold");
  cfg.getRange(secHdr, 3).setValue("Section Code")
     .setBackground("#EAD1DC").setFontWeight("bold");
  cfg.getRange(secHdr, 4).setValue("KM From")
     .setBackground("#EAD1DC").setFontWeight("bold");
  cfg.getRange(secHdr, 5).setValue("KM To")
     .setBackground("#EAD1DC").setFontWeight("bold");

  var sectionData = PWI_TMS_SECTIONS.map(function(s) {
    return ["", s.name, s.code, s.kmFrom, s.kmTo];
  });
  cfg.getRange(secHdr+1, 1, sectionData.length, 5).setValues(sectionData);

  // Update SOP tab
  _pwiUpdateSOP(ss);

  SpreadsheetApp.getUi().alert(
    "PWI_UNIT_MAP section added to _CONFIG_.\n\n" +
    "Next steps:\n" +
    "1. Fill in correct TMS auth codes for each SSE/P.WAY unit\n" +
    "2. Deploy CompliancePWI_Fetch.gs as Web App → paste URL in PWI_WEBAPP_URL\n" +
    "3. Run TMS fetch from Claude shortcut in Brave\n\n" +
    "See SYSTEM SOP tab for full instructions."
  );
}

// ── Helper: update SOP tab ────────────────────────────────────────────────────
function _pwiUpdateSOP(ss) {
  var sop = ss.getSheetByName("SYSTEM SOP");
  if (!sop) return;
  var lr = sop.getLastRow();
  sop.getRange(lr+2, 1).setValue("== COMPLIANCE BY PWI SYSTEM ==")
     .setFontWeight("bold").setBackground("#B6D7A8");
  var sopText = [
    ["PWI_UNIT_MAP CONFIG FORMAT:"],
    ["Each row: SR_DEN | AEN | SSE/P.WAY_UNIT | TMS_AUTH_1 | TMS_AUTH_2 | ... (up to 10 authorities)"],
    ["Example: SR DEN I/HWH | AEN/LLH | SSE/P.WAY/LLH | SSE/PW/LLH | JE/PW/SRP | SSE/PW/WS"],
    [""],
    ["TMS FETCH PROCEDURE:"],
    ["1. Open Brave browser, navigate to https://ircep.gov.in/TMS/HomeIframe.jsp"],
    ["2. Log in (enter credentials + OTP when prompted)"],
    ["3. Run the Claude TMS Shortcut from Claude sidebar"],
    ["4. Shortcut automatically fetches all 28 sections and populates the sheet"],
    ["5. Review orange-highlighted rows (potential duplicates) before deleting"],
    [""],
    ["EXCEPTION SCANNING:"],
    ["Run generateCompliancePWIReport() from Exception Report menu"],
    ["Scans date range: 1st of prev month (if today ≤ 7th) or 1st of current month"],
    ["Exception types:"],
    ["  Ex1: Pending > 15 days, no TMS compliance date"],
    ["  Ex2: TMS date filled, actual work+photo missing after 10 days"],
    ["  Ex3: Photo missing (actual date filled, photo column blank)"],
    ["  Ex4: Actual date filled, TMS date blank (> 3 day grace)"],
    ["  Ex5: Data quality — Pending Since date missing"],
    [""],
    ["MIGRATION LOG:"],
    ["v1: addPWIUnitMapSection() — adds PWI_UNIT_MAP, thresholds, section codes to _CONFIG_"]
  ];
  sop.getRange(lr+3, 1, sopText.length, 1).setValues(sopText);
}

// ── Helper: read PWI_UNIT_MAP from config ─────────────────────────────────────
function pwi_readUnitMap() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName(_PWI_CONFIG_TAB);
  if (!cfg) return {};
  var vals = cfg.getDataRange().getValues();
  var inSection = false, inData = false;
  var unitMap = {}; // SSE_UNIT → { srden, aen, tmsAuths: [] }

  for (var i = 0; i < vals.length; i++) {
    var col0 = String(vals[i][0]).trim();
    if (col0 === PWI_UNIT_MAP_SECTION) { inSection = true; continue; }
    if (!inSection) continue;
    if (col0 === "SR_DEN") { inData = true; continue; }
    // Stop at next section
    if (col0 && col0.indexOf("_") === 0) break;
    // Skip threshold rows
    if (col0.indexOf("PWI_") === 0) continue;

    if (inData && vals[i][2]) {
      var srden = String(vals[i][0]).trim();
      var aen   = String(vals[i][1]).trim();
      var unit  = String(vals[i][2]).trim();
      var auths = [];
      for (var c = 3; c < Math.min(vals[i].length, 12); c++) {
        if (vals[i][c] && String(vals[i][c]).trim()) auths.push(String(vals[i][c]).trim());
      }
      unitMap[unit] = { srden: srden, aen: aen, tmsAuths: auths };
    }
  }
  return unitMap;
}

// ── Helper: read PWI thresholds from config ───────────────────────────────────
function pwi_readThresholds() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName(_PWI_CONFIG_TAB);
  if (!cfg) return { pendingDays:15, complianceDays:10, graceDays:3, excludeBefore:null };
  var vals = cfg.getDataRange().getValues();
  var result = { pendingDays:15, complianceDays:10, graceDays:3, excludeBefore:null };
  for (var i = 0; i < vals.length; i++) {
    var k = String(vals[i][0]).trim(), v = vals[i][1];
    if (k === "PWI_PENDING_DAYS")         result.pendingDays    = parseInt(v)||15;
    if (k === "PWI_COMPLIANCE_DAYS")      result.complianceDays = parseInt(v)||10;
    if (k === "PWI_ACTUAL_VS_TMS_GRACE")  result.graceDays      = parseInt(v)||3;
    if (k === "PWI_EXCLUDE_BEFORE")       result.excludeBefore  = _pwi_parseDate(String(v).trim());
    if (k === "PWI_WEBAPP_URL")           result.webAppUrl      = String(v).trim();
  }
  return result;
}

// ── Helper: find TMS authority → SSE/P.WAY unit (reverse lookup) ──────────────
function pwi_tmsAuthToUnit(tmsAuth) {
  var unitMap = pwi_readUnitMap();
  var auth = String(tmsAuth||"").trim().toLowerCase();
  for (var unit in unitMap) {
    var auths = unitMap[unit].tmsAuths;
    for (var i = 0; i < auths.length; i++) {
      if (auths[i].toLowerCase() === auth) return unit;
    }
    // Also match SSE unit name directly
    if (unit.toLowerCase() === auth) return unit;
  }
  return null;
}

// ── Helper: get AEN for a SSE/P.WAY unit ─────────────────────────────────────
function pwi_unitToAEN(unit) {
  var unitMap = pwi_readUnitMap();
  return unitMap[unit] ? unitMap[unit].aen : null;
}

// ── One-time token management ─────────────────────────────────────────────────
function pwi_generateToken() {
  var token = Utilities.getUuid();
  var expiry = new Date(Date.now() + 120 * 60 * 1000); // 2 hours
  PropertiesService.getScriptProperties().setProperty("PWI_TOKEN", JSON.stringify({
    token: token, expiry: expiry.getTime()
  }));
  return token;
}

function pwi_validateToken(token) {
  var raw = PropertiesService.getScriptProperties().getProperty("PWI_TOKEN");
  if (!raw) return false;
  try {
    var stored = JSON.parse(raw);
    if (stored.token !== token) return false;
    if (Date.now() > stored.expiry) {
      PropertiesService.getScriptProperties().deleteProperty("PWI_TOKEN");
      return false;
    }
    // Keep token valid until expiry — allows retry if POST fails
    return true;
  } catch(e) { return false; }
}

// Generate token and show to user (called from menu before running shortcut)
function pwi_prepareToken() {
  var token = pwi_generateToken();
  var cfg   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(_PWI_CONFIG_TAB);
  var webAppUrl = "";
  if (cfg) {
    var vals = cfg.getDataRange().getValues();
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][0]).trim() === "PWI_WEBAPP_URL") { webAppUrl = String(vals[i][1]).trim(); break; }
    }
  }
  SpreadsheetApp.getUi().alert(
    "TMS Fetch Token Generated\n\n" +
    "Token (valid 30 min): " + token + "\n\n" +
    "Web App URL: " + (webAppUrl || "NOT SET — deploy web app first") + "\n\n" +
    "The Claude TMS Shortcut will use these automatically.\n" +
    "Run the shortcut in Brave now."
  );
  return { token: token, webAppUrl: webAppUrl };
}