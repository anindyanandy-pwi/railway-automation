/**
 * GPS Cross-check – results store, settings and dashboard.
 *
 * Lives inside a Google Sheet you own ("GPS Cross-check – Results").
 * Deploy > New deployment > Web app:  Execute as: Me   |   Who has access: Anyone
 *  - The daily GitHub run sends results here (protected by RESULTS_KEY).
 *  - The dashboard opens at  <web app URL>?key=<DASHBOARD_KEY>
 *
 * Script properties (Project Settings > Script properties):
 *   RESULTS_KEY     long random text, same value as the GitHub secret RESULTS_KEY
 *   DASHBOARD_KEY   password for opening the dashboard
 *   GITHUB_REPO     anindyanandy-pwi/railway-automation
 *   GITHUB_TOKEN    (optional) fine-grained token with "Actions: write" on that repo, for the Run now button
 */

var RESULTS = 'Results';
var HASHES = 'PhotoHashes';
var SETTINGS = 'Settings';
var RUNS = 'Runs';
var WORKFLOW = 'gps-crosscheck.yml';

var COLS = ['run_date', 'check', 'duty_date', 'status', 'sr_den', 'aen', 'pwi', 'beat', 'names',
  'gps_in_sheet', 'vendor', 'portal_device', 'roster', 'sheet_start', 'sheet_end', 'sheet_km', 'sheet_speed',
  'gps_start', 'gps_end', 'gps_km', 'gps_max_speed', 'gps_avg_speed', 'start_photo_time', 'end_photo_time',
  'actual_start', 'actual_start_from', 'actual_end', 'findings_text', 'portal_text',
  'start_photo', 'end_photo', 'sheet_link', 'findings_json', 'portal_json'];

var SETTING_HELP = {
  start_tolerance_min: 'Allowed difference in start time (minutes)',
  end_tolerance_min: 'Allowed difference in end time (minutes)',
  km_min_percent: 'GPS km must be at least this % of the beat km',
  sheet_km_tolerance_percent: 'Sheet km vs GPS km: allowed difference (%)',
  max_walking_speed_kmph: 'Above this speed (km/h) = high speed',
  speed_spike_points: 'High speed must last at least this many GPS points (else GPS jitter)',
  photo_vs_gps_max_km: 'Photo location vs GPS position: allowed distance (km)',
  photo_max_delay_min: 'Stamp older than WhatsApp time by more than this (minutes) = old photo?',
  photo_ahead_tolerance_min: 'Stamp later than WhatsApp time by more than this (minutes) = altered',
  gps_gap_switched_off_min: 'No GPS point for this long (minutes) = switched off'
};
var SETTING_DEFAULTS = {
  start_tolerance_min: 15, end_tolerance_min: 15, km_min_percent: 90, sheet_km_tolerance_percent: 10,
  max_walking_speed_kmph: 5, speed_spike_points: 2, photo_vs_gps_max_km: 1.0, photo_max_delay_min: 120,
  photo_ahead_tolerance_min: 2, gps_gap_switched_off_min: 30
};

/** Run once from the editor after pasting the code. Creates the tabs and the two keys. */
function setup() {
  var ss = SpreadsheetApp.getActive();
  sheet_(RESULTS, COLS);
  sheet_(HASHES, ['date', 'drive_id', 'role', 'dhash', 'batch']);
  sheet_(RUNS, ['run_date', 'received_at', 'records', 'red', 'amber', 'check', 'ok']);
  var s = sheet_(SETTINGS, ['setting', 'value', 'meaning']);
  if (s.getLastRow() < 2) {
    var rows = Object.keys(SETTING_DEFAULTS).map(function (k) { return [k, SETTING_DEFAULTS[k], SETTING_HELP[k]]; });
    s.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  var p = PropertiesService.getScriptProperties();
  if (!p.getProperty('RESULTS_KEY')) p.setProperty('RESULTS_KEY', Utilities.getUuid() + Utilities.getUuid());
  if (!p.getProperty('DASHBOARD_KEY')) p.setProperty('DASHBOARD_KEY', Utilities.getUuid().slice(0, 13));
  if (!p.getProperty('GITHUB_REPO')) p.setProperty('GITHUB_REPO', 'anindyanandy-pwi/railway-automation');
  Logger.log('RESULTS_KEY (put this in GitHub secret RESULTS_KEY): ' + p.getProperty('RESULTS_KEY'));
  Logger.log('DASHBOARD_KEY (add ?key=... to the web app URL): ' + p.getProperty('DASHBOARD_KEY'));
  ss.toast('Setup done. Open View > Logs (Execution log) to copy the two keys.');
}

function sheet_(name, header) {
  var ss = SpreadsheetApp.getActive();
  var s = ss.getSheetByName(name) || ss.insertSheet(name);
  if (s.getLastRow() === 0) {
    s.appendRow(header);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, header.length).setFontWeight('bold');
  }
  return s;
}

function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================================ web endpoints

function doGet(e) {
  var p = e.parameter || {};
  if (p.action) {
    if (p.key !== prop_('RESULTS_KEY')) return json_({ ok: false, error: 'bad key' });
    if (p.action === 'settings') return json_({ ok: true, settings: getSettings() });
    if (p.action === 'hashes') return json_({ ok: true, hashes: getHashes_(Number(p.days || 30)) });
    return json_({ ok: false, error: 'unknown action' });
  }
  if (p.key !== prop_('DASHBOARD_KEY')) {
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif">Access key needed. ' +
      'Open the dashboard with the link that ends in <b>?key=…</b></p>');
  }
  var t = HtmlService.createTemplateFromFile('Dashboard');
  t.key = p.key;
  return t.evaluate().setTitle('GPS Cross-check').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  if ((e.parameter || {}).key !== prop_('RESULTS_KEY')) return json_({ ok: false, error: 'bad key' });
  var lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    var data = JSON.parse(e.postData.contents);
    saveRun_(data);
    return json_({ ok: true, saved: data.records.length });
  } finally {
    lock.releaseLock();
  }
}

function saveRun_(data) {
  var s = sheet_(RESULTS, COLS);
  var runDate = data.run_date;
  // a re-run for the same date replaces the earlier rows of that date
  deleteWhere_(s, function (r) { return r[0] === runDate; });
  var rows = data.records.map(function (r) {
    r.findings_text = r.findings.map(function (f) { return '• ' + f.message; }).join('\n');
    r.portal_text = r.portal_exceptions.map(function (v) { return v.exception + ': ' + v.verdict + ' – ' + v.reason; }).join('\n');
    r.findings_json = JSON.stringify(r.findings);
    r.portal_json = JSON.stringify(r.portal_exceptions);
    return COLS.map(function (c) { var v = r[c]; return v === null || v === undefined ? '' : v; });
  });
  if (rows.length) s.getRange(s.getLastRow() + 1, 1, rows.length, COLS.length).setNumberFormat('@').setValues(rows);

  var h = sheet_(HASHES, ['date', 'drive_id', 'role', 'dhash', 'batch']);
  var known = {};
  if (h.getLastRow() > 1) h.getRange(2, 2, h.getLastRow() - 1, 1).getValues().forEach(function (x) { known[x[0]] = 1; });
  var hrows = [];
  data.records.forEach(function (r) {
    (r.photo_hashes || []).forEach(function (p) {
      if (!known[p.drive_id]) { known[p.drive_id] = 1; hrows.push([r.duty_date, p.drive_id, p.role, p.dhash, r.pwi + ' ' + r.beat]); }
    });
  });
  if (hrows.length) h.getRange(h.getLastRow() + 1, 1, hrows.length, 5).setNumberFormat('@').setValues(hrows);

  var c = { RED: 0, AMBER: 0, CHECK: 0, OK: 0 };
  data.records.forEach(function (r) { c[r.status] = (c[r.status] || 0) + 1; });
  sheet_(RUNS, ['run_date', 'received_at', 'records', 'red', 'amber', 'check', 'ok'])
    .appendRow([runDate, data.generated_at, data.records.length, c.RED, c.AMBER, c.CHECK, c.OK]);
}

function deleteWhere_(s, test) {
  var n = s.getLastRow();
  if (n < 2) return;
  var vals = s.getRange(2, 1, n - 1, 1).getDisplayValues();
  // delete bottom-up in contiguous blocks
  for (var i = vals.length - 1; i >= 0;) {
    if (!test([vals[i][0]])) { i--; continue; }
    var end = i;
    while (i >= 0 && test([vals[i][0]])) i--;
    s.deleteRows(i + 2, end - i);
  }
}

function getHashes_(days) {
  var h = SpreadsheetApp.getActive().getSheetByName(HASHES);
  if (!h || h.getLastRow() < 2) return [];
  var since = new Date(Date.now() - days * 86400000);
  return h.getRange(2, 1, h.getLastRow() - 1, 5).getDisplayValues()
    .filter(function (r) { return new Date(r[0]) >= since; })
    .map(function (r) { return { date: r[0], drive_id: r[1], role: r[2], dhash: r[3], batch: r[4] }; });
}

// ============================================================================ used by the dashboard

function getSettings() {
  var s = sheet_(SETTINGS, ['setting', 'value', 'meaning']);
  var out = {};
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 2).getValues().forEach(function (r) { out[r[0]] = r[1]; });
  return out;
}

function settingsForDashboard(key) {
  check_(key);
  var cur = getSettings();
  return Object.keys(SETTING_DEFAULTS).map(function (k) {
    return { key: k, value: cur[k] === undefined || cur[k] === '' ? SETTING_DEFAULTS[k] : cur[k], help: SETTING_HELP[k] };
  });
}

function saveSettings(key, values) {
  check_(key);
  var s = sheet_(SETTINGS, ['setting', 'value', 'meaning']);
  var rows = Object.keys(SETTING_DEFAULTS).map(function (k) {
    var v = Number(values[k]);
    if (isNaN(v) || v < 0) throw new Error('Invalid value for ' + k);
    return [k, v, SETTING_HELP[k]];
  });
  s.getRange(2, 1, Math.max(s.getLastRow() - 1, 1), 3).clearContent();
  s.getRange(2, 1, rows.length, 3).setValues(rows);
  return true;
}

function check_(key) { if (key !== prop_('DASHBOARD_KEY')) throw new Error('Not allowed'); }

function listRuns(key) {
  check_(key);
  var s = SpreadsheetApp.getActive().getSheetByName(RUNS);
  if (!s || s.getLastRow() < 2) return [];
  var seen = {}, out = [];
  s.getRange(2, 1, s.getLastRow() - 1, 7).getDisplayValues().reverse().forEach(function (r) {
    if (!seen[r[0]]) { seen[r[0]] = 1; out.push({ run_date: r[0], at: r[1], records: r[2], red: r[3], amber: r[4], check: r[5], ok: r[6] }); }
  });
  return out.slice(0, 120);
}

function getRun(key, runDate) {
  check_(key);
  var s = SpreadsheetApp.getActive().getSheetByName(RESULTS);
  if (!s || s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, COLS.length).getDisplayValues()
    .filter(function (r) { return r[0] === runDate; })
    .map(function (r) {
      var o = {};
      COLS.forEach(function (c, i) { o[c] = r[i]; });
      o.findings = JSON.parse(o.findings_json || '[]');
      o.portal_exceptions = JSON.parse(o.portal_json || '[]');
      delete o.findings_json; delete o.portal_json;
      return o;
    });
}

/** Excel / PDF of what the dashboard is showing. Saved in Drive folder "GPS Cross-check exports". */
function exportRows(key, runDate, rows, format) {
  check_(key);
  var head = ['Check', 'Duty date', 'Status', 'Sr DEN', 'AEN', 'PWI', 'Beat / batch', 'Names', 'GPS', 'Roster',
    'Sheet start', 'Sheet end', 'Sheet km', 'GPS start', 'GPS end', 'GPS km', 'GPS max speed',
    'Start photo', 'End photo', 'Actual start', 'Findings', 'Portal exceptions – verdict'];
  var data = [head].concat(rows.map(function (r) {
    return [r.check, r.duty_date, r.status, r.sr_den, r.aen, r.pwi, r.beat, r.names, r.gps_in_sheet, r.roster,
      r.sheet_start, r.sheet_end, r.sheet_km, r.gps_start, r.gps_end, r.gps_km, r.gps_max_speed,
      r.start_photo_time, r.end_photo_time, (r.actual_start + ' ' + (r.actual_start_from ? '(' + r.actual_start_from + ')' : '')).trim(),
      r.findings_text, r.portal_text];
  }));
  var name = 'GPS cross-check ' + runDate;
  var tmp = SpreadsheetApp.create(name);
  var sh = tmp.getSheets()[0];
  sh.getRange(1, 1, data.length, head.length).setValues(data);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#e8eaed');
  sh.getRange(1, 1, data.length, head.length).setWrap(true).setVerticalAlignment('top').setFontSize(9);
  sh.setColumnWidths(1, head.length, 90);
  sh.setColumnWidth(21, 320); sh.setColumnWidth(22, 260); sh.setColumnWidth(7, 160);
  for (var i = 2; i <= data.length; i++) {
    var st = data[i - 1][2];
    var col = st === 'RED' ? '#fce8e6' : st === 'AMBER' ? '#fef7e0' : st === 'CHECK' ? '#e8f0fe' : '#e6f4ea';
    sh.getRange(i, 3).setBackground(col);
  }
  SpreadsheetApp.flush();
  var id = tmp.getId();
  var url = 'https://docs.google.com/spreadsheets/d/' + id + '/export?' + (format === 'pdf'
    ? 'format=pdf&portrait=false&size=A4&fitw=true&gridlines=true&printtitle=false&sheetnames=false&fzr=true&top_margin=0.4&bottom_margin=0.4&left_margin=0.3&right_margin=0.3'
    : 'format=xlsx');
  var blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob()
    .setName(name + (format === 'pdf' ? '.pdf' : '.xlsx'));
  var folders = DriveApp.getFoldersByName('GPS Cross-check exports');
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('GPS Cross-check exports');
  var file = folder.createFile(blob);
  DriveApp.getFileById(id).setTrashed(true);
  return { url: file.getUrl(), name: file.getName() };
}

/** Run now: starts the GitHub job directly if a token is set, else returns the page with the Run button. */
function runNow(key, runDate) {
  check_(key);
  var repo = prop_('GITHUB_REPO');
  var page = 'https://github.com/' + repo + '/actions/workflows/' + WORKFLOW;
  var token = prop_('GITHUB_TOKEN');
  if (!token) return { started: false, open: page };
  var res = UrlFetchApp.fetch('https://api.github.com/repos/' + repo + '/actions/workflows/' + WORKFLOW + '/dispatches', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json' },
    payload: JSON.stringify({ ref: 'main', inputs: { date: runDate || '' } })
  });
  if (res.getResponseCode() !== 204) return { started: false, open: page, error: 'GitHub said ' + res.getResponseCode() };
  return { started: true, open: page };
}
