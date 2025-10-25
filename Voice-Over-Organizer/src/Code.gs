// @ts-nocheck
// Voice Tracker — Universal, per-spreadsheet settings

// All settings are stored in DocumentProperties as key "VT_SETTINGS".
// Nothing hardcoded: users can paste folder URLs/IDs and choose a filename rule.

// ===========================
// Settings helpers
// ===========================
function defaultSettings_() {
  return {
    voiceFolderId: '',
    auditionFolderId: '',
    // Filename → Character extraction
    // Choose ONE mode below:
    filenameMode: 'regex',              // 'regex' | 'delimiter'
    regexPattern: ' - (.+?) -',         // capture group for character
    regexGroup: 1,                      // which capture group (1-based)
    delimiter: ' - ',                   // used when filenameMode='delimiter'
    charIndex: 2,                       // 1-based index of token (e.g., "Name - Character - ..." => 2)
    // Voice → Auditions status mapping
    voiceToAuditionMap: { "Current": "Booked", "Past": "Submitted" },
    // Auditions statuses that manual edits should never be overwritten by sync
    dontOverwrite: ["Passed"]
  };
}

function getSettings() {
  var props = PropertiesService.getDocumentProperties();
  var raw = props.getProperty('VT_SETTINGS');
  var s = defaultSettings_();
  if (raw) {
    try {
      var parsed = JSON.parse(raw);
      for (var k in parsed) s[k] = parsed[k];
    } catch (e) {}
  }
  return s;
}

function saveSettings(payload) {
  var s = getSettings();

  // Accept URL or raw ID
  var vid = extractIdFromInput_(payload.voiceFolder || '');
  var aid = extractIdFromInput_(payload.auditionFolder || '');

  // Validate if provided
  if (vid) DriveApp.getFolderById(vid); // throws if invalid
  if (aid) DriveApp.getFolderById(aid);

  s.voiceFolderId = vid;
  s.auditionFolderId = aid;

  // filename mode
  s.filenameMode = (payload.filenameMode === 'delimiter') ? 'delimiter' : 'regex';
  s.regexPattern = String(payload.regexPattern || s.regexPattern);
  s.regexGroup   = Math.max(1, parseInt(payload.regexGroup || s.regexGroup, 10));
  s.delimiter    = String(payload.delimiter || s.delimiter);
  s.charIndex    = Math.max(1, parseInt(payload.charIndex || s.charIndex, 10));

  // mapping
  var map = payload.voiceToAuditionMap || {};
  s.voiceToAuditionMap = {
    "Current": map.Current || "Booked",
    "Past": map.Past || "Submitted"
  };

  // never-overwrite list
  s.dontOverwrite = Array.isArray(payload.dontOverwrite) && payload.dontOverwrite.length
    ? payload.dontOverwrite.map(String)
    : ["Passed"];

  PropertiesService.getDocumentProperties().setProperty('VT_SETTINGS', JSON.stringify(s));
  return true;
}

function extractIdFromInput_(input) {
  if (!input) return '';
  var s = String(input).trim();
  // /drive/folders/<id> or ?id=<id>
  var m = s.match(/(?:\/folders\/|[\?\&]id=)([a-zA-Z0-9_-]+)/);
  if (m && m[1]) return m[1];
  // Fallback: the longest Drive-looking token
  var t = s.match(/[-\w]{25,}/);
  return t ? t[0] : s;
}

function ensureConfigured_(needsAuditions) {
  var s = getSettings();
  if (!s.voiceFolderId) throw new Error('Set your Voice folder in the Settings tab.');
  if (needsAuditions && !s.auditionFolderId) throw new Error('Set your Auditions folder in the Settings tab.');
  return s;
}

// ===========================
// Sheet helpers & headers
// ===========================
var VOICE_HEADERS = ["Folder", "File Name", "Character", "Date Added", "Mime Type", "File Link", "Final Link", "Status"];
var AUDITION_HEADERS = ["Folder", "File Name", "Character", "Date Added", "Mime Type", "File Link", "Status"];

function arraysEqual_(a, b) {
  if (!a || !b || a.length !== b.length) return false;
  for (var i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

function getVoiceSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    if (sh.getName() === 'Auditions') continue;
    var head = sh.getRange(2, 1, 1, 8).getValues()[0] || [];
    if (arraysEqual_(head, VOICE_HEADERS)) return sh;
  }
  var active = ss.getActiveSheet();
  var candidate = active.getRange(2, 1, 1, 8).getValues()[0] || [];
  if (candidate[6] === "Final Link" && candidate[7] === "Status") return active;
  return null;
}

function getAuditionSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Auditions');
  if (!sheet) {
    sheet = ss.insertSheet('Auditions');
    sheet.getRange(2, 1, 1, AUDITION_HEADERS.length).setValues([AUDITION_HEADERS]);
  }
  return sheet;
}

// Extract URL from a cell (HYPERLINK / richtext / raw)
function getUrlFromCell(range) {
  var formula = range.getFormula();
  if (formula && /^=HYPERLINK\(/i.test(formula)) {
    var m = formula.match(/=HYPERLINK\("([^"]+)"/i);
    if (m && m[1]) return m[1];
  }
  var rich = range.getRichTextValue && range.getRichTextValue();
  if (rich) {
    var runs = rich.getRuns();
    for (var i = 0; i < runs.length; i++) {
      var link = runs[i].getLinkUrl();
      if (link) return link;
    }
  }
  var val = range.getValue();
  if (typeof val === 'string' && /^https?:\/\//i.test(val)) return val;
  return '';
}

// Character extraction (universal)
function extractCharacterFromName_(fileName, settings) {
  try {
    if (settings.filenameMode === 'delimiter') {
      var parts = String(fileName).split(settings.delimiter || ' - ');
      var idx = Math.max(0, (settings.charIndex || 1) - 1);
      var got = (parts[idx] || '').trim();
      return got || 'Unknown';
    } else {
      var re = new RegExp(settings.regexPattern || ' - (.+?) -');
      var m = String(fileName).match(re);
      var g = Math.max(1, settings.regexGroup || 1);
      if (m && m[g]) return m[g].trim();
    }
  } catch (e) {}
  return 'Unknown';
}

// Collect files recursively with character extraction
function collectFiles(folder, path, filesData, allCharacters, settings) {
  var folderName = folder.getName();
  var currentPath = path ? path + "/" + folderName : folderName;

  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    var characterName = extractCharacterFromName_(fileName, settings);

    allCharacters.add(characterName);
    filesData.push({
      folderPath: currentPath,
      fileName: fileName,
      characterName: characterName,
      dateAdded: file.getDateCreated(),
      mimeType: file.getMimeType(),
      url: file.getUrl()
    });
  }

  var subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    collectFiles(subfolders.next(), currentPath, filesData, allCharacters, settings);
  }
}

// ===========================
// Voice Projects (OG)
// ===========================
function updateVoiceProjects() {
  var s = ensureConfigured_(false);
  var folder = DriveApp.getFolderById(s.voiceFolderId);
  var sheet = getVoiceSheet() || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Preserve existing Final Link (G) & Status (H)
  var existingData = {};
  var oldLastRow = sheet.getLastRow();
  if (oldLastRow >= 3) {
    var oldVals = sheet.getRange(3, 1, oldLastRow - 2, 8).getValues();
    var oldForm = sheet.getRange(3, 7, oldLastRow - 2, 1).getFormulas();
    for (var i = 0; i < oldVals.length; i++) {
      var fileName = oldVals[i][1];
      var character = oldVals[i][2];
      var finalLinkValue = oldVals[i][6];
      var statusValue = oldVals[i][7];
      var finalLinkFormula = (oldForm && oldForm[i] && oldForm[i][0]) ? oldForm[i][0] : '';
      if (character && fileName) {
        existingData[character + '|' + fileName] = {
          finalLinkFormula: finalLinkFormula,
          finalLinkValue: finalLinkValue,
          status: statusValue
        };
      }
    }
  }

  // Collect from Drive
  var allCharacters = new Set();
  var filesData = [];
  collectFiles(folder, "", filesData, allCharacters, s);

  // Ensure rows
  var requiredRows = filesData.length + 2;
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }

  // Clear A:H to avoid leftovers
  if (oldLastRow >= 3) {
    var clearCount = oldLastRow - 2;
    sheet.getRange(3, 1, clearCount, 8).clearContent().clearFormat();
  }

  // Title & headers
  var titleRange = sheet.getRange(1, 1, 1, 8);
  try { if (!titleRange.isPartOfMerge()) titleRange.merge(); } catch (e) { try { titleRange.merge(); } catch(e2) {} }
  titleRange.setValue("🎙️ Voice Projects Overview")
    .setFontWeight("bold").setFontSize(14).setFontFamily("Roboto")
    .setBackground("#81C784").setFontColor("#000000")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  sheet.getRange(2, 1, 1, VOICE_HEADERS.length).setValues([VOICE_HEADERS])
    .setFontWeight("bold").setFontFamily("Roboto").setFontSize(10)
    .setBackground("#C8E6C9").setFontColor("#000000")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, false, false, "#A5D6A7", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Build rows
  var out = new Array(filesData.length);
  for (var i = 0; i < filesData.length; i++) {
    var d = filesData[i];
    var key = d.characterName + '|' + d.fileName;
    var store = existingData[key] || {};
    var gCell = '';
    if (store.finalLinkFormula) gCell = store.finalLinkFormula;
    else if (store.finalLinkValue && typeof store.finalLinkValue === 'string') {
      var v = store.finalLinkValue.trim();
      gCell = /^https?:\/\//i.test(v) ? ('=HYPERLINK("' + v + '","Final Production")') : v;
    }
    var hCell = store.status ? store.status : "Current";

    out[i] = [
      d.folderPath, d.fileName, d.characterName, d.dateAdded, d.mimeType,
      '=HYPERLINK("' + d.url + '","🎵 View File")',
      gCell, hCell
    ];
  }

  if (filesData.length > 0) {
    sheet.getRange(3, 1, filesData.length, 8).setValues(out);
    var dataRange = sheet.getRange(3, 1, filesData.length, 8);
    dataRange
      .setFontFamily("Roboto").setFontSize(10)
      .setHorizontalAlignment("left").setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true, "#E0E0E0", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(3, 4, filesData.length, 1).setHorizontalAlignment("center");
    sheet.getRange(3, 8, filesData.length, 1).setHorizontalAlignment("center");

    var bg = [];
    for (var r = 0; r < filesData.length; r++) {
      bg.push(new Array(8).fill(r % 2 === 0 ? "#F1F8E9" : "#FFFFFF"));
    }
    dataRange.setBackgrounds(bg);

    sheet.autoResizeColumns(1, 8);
    var minW = [150, 200, 130, 100, 120, 150, 150, 100];
    for (var c = 1; c <= 8; c++) if (sheet.getColumnWidth(c) < minW[c - 1]) sheet.setColumnWidth(c, minW[c - 1]);
    dataRange.setWrap(true);
  }

  // Persist characters for UI dropdowns
  PropertiesService.getDocumentProperties().setProperty('CHARACTERS', JSON.stringify(Array.from(allCharacters)));

  // Keep auditions in sync (if configured)
  try { syncAuditionsWithProjects(sheet, getAuditionSheet()); } catch (e) { Logger.log(e); }
}

// ===========================
// Insights API (for Sidebar)
// ===========================
function getStatsAndRecent(limit) {
  limit = limit || 12;
  var now = new Date();
  var recent7  = new Date(now.getTime() - 7  * 24 * 60 * 60 * 1000);
  var recent30 = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);

  var out = {
    voice: { total: 0, current: 0, past: 0, recent7: 0, recent30: 0 },
    auditions: { total: 0, pending: 0, submitted: 0, booked: 0, passed: 0, recent7: 0, recent30: 0 },
    recent: []
  };

  var voiceSheet = getVoiceSheet();
  if (voiceSheet) {
    var last = voiceSheet.getLastRow();
    if (last >= 3) {
      var vals = voiceSheet.getRange(3, 1, last - 2, 8).getValues();
      for (var i = 0; i < vals.length; i++) {
        var row = vals[i];
        var fileName = row[1], character = row[2], date = row[3], status = (row[7] || '').toString();
        out.voice.total++;
        if (status === 'Current') out.voice.current++; else if (status === 'Past') out.voice.past++;
        if (date && date instanceof Date) {
          if (date >= recent7)  out.voice.recent7++;
          if (date >= recent30) out.voice.recent30++;
          out.recent.push({ type:'Voice', date:date.toISOString(), character:character||'', fileName:fileName||'', status:status||'' });
        }
      }
    }
  }

  var audSheet = getAuditionSheet();
  if (audSheet) {
    var lastA = audSheet.getLastRow();
    if (lastA >= 4) {
      var valsA = audSheet.getRange(4, 1, lastA - 3, 7).getValues();
      for (var j = 0; j < valsA.length; j++) {
        var rowA = valsA[j];
        var fileNameA = rowA[1], characterA = rowA[2], dateA = rowA[3], statA = (rowA[6] || '').toString();
        out.auditions.total++;
        var key = statA.toLowerCase();
        if (key === 'pending') out.auditions.pending++;
        else if (key === 'submitted') out.auditions.submitted++;
        else if (key === 'booked') out.auditions.booked++;
        else if (key === 'passed') out.auditions.passed++;
        if (dateA && dateA instanceof Date) {
          if (dateA >= recent7)  out.auditions.recent7++;
          if (dateA >= recent30) out.auditions.recent30++;
          out.recent.push({ type:'Audition', date:dateA.toISOString(), character:characterA||'', fileName:fileNameA||'', status:statA||'' });
        }
      }
    }
  }

  out.recent = out.recent.filter(function(x){return x.date;})
                         .sort(function(a,b){return new Date(b.date)-new Date(a.date);})
                         .slice(0, limit);
  return out;
}

// ===========================
// Auditions
// ===========================
function updateAuditions() {
  var s = ensureConfigured_(true);
  var folder = DriveApp.getFolderById(s.auditionFolderId);
  var sheet = getAuditionSheet();

  var preserve = {};
  var lastRow = sheet.getLastRow();
  if (lastRow >= 3) {
    var vals = sheet.getRange(3, 1, lastRow - 2, 7).getValues();
    for (var i = 0; i < vals.length; i++) {
      var fileName = vals[i][1];
      var charName = vals[i][2] || '';
      preserve[charName + '|' + fileName] = vals[i][6];
    }
  }

  var filesData = [];
  collectFiles(folder, "", filesData, new Set(), s);

  // Reset layout fast
  sheet.setFrozenRows(0);
  sheet.clearContents().clearFormats();
  if (sheet.getMaxRows() < 4) sheet.insertRowsAfter(sheet.getMaxRows(), 4 - sheet.getMaxRows());
  var totalRows = sheet.getMaxRows();
  if (totalRows > 4) sheet.deleteRows(5, totalRows - 4);
  sheet.setFrozenRows(Math.min(3, sheet.getMaxRows() - 1));

  // Title & headers
  sheet.getRange("A1:G1").merge().setValue("🎧 Audition Tracker")
    .setFontSize(16).setFontWeight("bold").setFontFamily("Roboto")
    .setBackground("#2E7D32").setFontColor("#FFFFFF")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("A2:G2").merge().setValue("Automatically updated from your Google Drive audition folder")
    .setFontSize(10).setFontColor("#E8F5E9").setHorizontalAlignment("center");

  sheet.getRange(3, 1, 1, AUDITION_HEADERS.length).setValues([AUDITION_HEADERS])
    .setFontFamily("Roboto").setFontSize(10).setFontWeight("bold")
    .setBackground("#A5D6A7").setFontColor("#1B5E20")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, false, false, "#81C784", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  var out = new Array(filesData.length);
  for (var i = 0; i < filesData.length; i++) {
    var f = filesData[i];
    var key = f.characterName + '|' + f.fileName;
    var status = preserve[key] || "Pending";
    out[i] = [
      f.folderPath, f.fileName, f.characterName, f.dateAdded, f.mimeType,
      '=HYPERLINK("' + f.url + '","🎵 View Audition")',
      status
    ];
  }

  if (filesData.length > 0) {
    sheet.getRange(4, 1, filesData.length, 7).setValues(out);
    var dataRange = sheet.getRange(4, 1, filesData.length, 7);
    dataRange
      .setFontFamily("Roboto").setFontSize(10)
      .setHorizontalAlignment("left").setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true, "#C8E6C9", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(4, 4, filesData.length, 1).setHorizontalAlignment("center");
    sheet.getRange(4, 7, filesData.length, 1).setHorizontalAlignment("center");

    var bg = [];
    for (var r = 0; r < filesData.length; r++) bg.push(new Array(7).fill(r % 2 === 0 ? "#E8F5E9" : "#FFFFFF"));
    dataRange.setBackgrounds(bg);
    sheet.autoResizeColumns(1, 7);
  }

  var statusRange = sheet.getRange(4, 7, Math.max(1, filesData.length), 1);
  var rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Pending").setBackground("#FFF9C4").setFontColor("#F57F17").setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Submitted").setBackground("#BBDEFB").setFontColor("#0D47A1").setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Booked").setBackground("#C8E6C9").setFontColor("#1B5E20").setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Passed").setBackground("#FFCDD2").setFontColor("#B71C1C").setRanges([statusRange]).build()
  ];
  sheet.setConditionalFormatRules(rules);

  addAuditionStatusDropdown();

  try { syncAuditionsWithProjects(getVoiceSheet(), sheet); } catch (e) { Logger.log(e); }
}

function addAuditionStatusDropdown() {
  var sheet = getAuditionSheet();
  var lastRow = Math.max(4, sheet.getLastRow());
  var range = sheet.getRange(4, 7, lastRow - 3);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Pending", "Submitted", "Booked", "Passed"], true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
}

// ===========================
// Sync Auditions ⇄ Voice Projects
// ===========================
function buildVoiceIndex_(voiceSheet) {
  var idx = {};
  if (!voiceSheet) return idx;
  var last = voiceSheet.getLastRow();
  if (last < 3) return idx;
  var vals = voiceSheet.getRange(3, 1, last - 2, 8).getValues();
  for (var i = 0; i < vals.length; i++) {
    var fileName = vals[i][1], charName = vals[i][2], status = vals[i][7];
    if (!fileName || !charName) continue;
    idx[charName + '|' + fileName] = status;
  }
  return idx;
}

/**
 * Mapping is configurable in Settings:
 *   e.g., {"Current":"Booked","Past":"Submitted"}
 * Never overwrite statuses listed in settings.dontOverwrite (default: ["Passed"]).
 */
function syncAuditionsWithProjects(voiceSheet, auditionSheet) {
  var settings = getSettings();
  voiceSheet = voiceSheet || getVoiceSheet();
  auditionSheet = auditionSheet || getAuditionSheet();
  if (!voiceSheet || !auditionSheet) return { booked: 0, submitted: 0 };

  var map = settings.voiceToAuditionMap || {"Current":"Booked","Past":"Submitted"};
  var protect = {};
  (settings.dontOverwrite || ["Passed"]).forEach(function(x){ protect[String(x||'')] = true; });

  var vmap = buildVoiceIndex_(voiceSheet);
  var last = auditionSheet.getLastRow();
  if (last < 4) return { booked: 0, submitted: 0 };

  var counters = { booked: 0, submitted: 0 };

  for (var r = 4; r <= last; r++) {
    var fileName = auditionSheet.getRange(r, 2).getValue();
    var charName = auditionSheet.getRange(r, 3).getValue();
    if (!fileName || !charName) continue;

    var key = charName + '|' + fileName;
    var voiceStatus = vmap[key];
    if (!voiceStatus) continue;

    var target = map[voiceStatus];
    if (!target) continue;

    var currentAud = auditionSheet.getRange(r, 7).getValue();
    if (protect[currentAud]) continue;

    if (currentAud !== target) {
      auditionSheet.getRange(r, 7).setValue(target);
      if (target === 'Booked') counters.booked++;
      if (target === 'Submitted') counters.submitted++;
    }
  }
  return counters;
}

function syncNow() { return syncAuditionsWithProjects(getVoiceSheet(), getAuditionSheet()); }

// ===========================
// Filters & Search
// ===========================
function filterByStatus(status) {
  var sheet = getVoiceSheet() || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  status = (status || '').toString().trim();
  var all = !status || status.toLowerCase() === 'all';

  var rows = lastRow - 2;
  sheet.showRows(3, rows);
  if (all) return;

  var vals = sheet.getRange(3, 8, rows, 1).getValues(); // H
  var target = status.toLowerCase();
  var start = null;
  for (var i = 0; i < rows; i++) {
    var v = (vals[i][0] || '').toString().trim().toLowerCase();
    var mismatch = v !== target;
    if (mismatch && start === null) start = 3 + i;
    if (!mismatch && start !== null) { sheet.hideRows(start, (3 + i) - start); start = null; }
  }
  if (start !== null) sheet.hideRows(start, (lastRow + 1) - start);
}

// Auditions: combined status+search
function filterAuditionStatus(status, query) {
  var sheet = getAuditionSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return;

  status = (status || '').toString().trim();
  query  = (query  || '').toString().trim().toLowerCase();

  var allStatus = !status || status.toLowerCase() === 'all';
  var dataRows = lastRow - 3;
  sheet.showRows(4, dataRows);
  if (allStatus && !query) return;

  var meta = sheet.getRange(4, 1, dataRows, 3).getValues(); // A..C
  var stats = sheet.getRange(4, 7, dataRows, 1).getValues(); // G
  var target = status.toLowerCase();

  var hideStart = null;
  for (var i = 0; i < dataRows; i++) {
    var rowFolder = (meta[i][0] || '').toString();
    var rowFile   = (meta[i][1] || '').toString();
    var rowChar   = (meta[i][2] || '').toString();
    var rowStat   = (stats[i][0] || '').toString();

    var matchStatus = allStatus || rowStat.toLowerCase().trim() === target;
    var haystack = (rowFolder + ' ' + rowFile + ' ' + rowChar).toLowerCase();
    var matchQuery  = !query || haystack.indexOf(query) !== -1;

    var keep = matchStatus && matchQuery;
    if (!keep && hideStart === null) hideStart = 4 + i;
    if (keep && hideStart !== null) { sheet.hideRows(hideStart, (4 + i) - hideStart); hideStart = null; }
  }
  if (hideStart !== null) sheet.hideRows(hideStart, (lastRow + 1) - hideStart);
}

function searchVoiceProjects(character, folderKeyword, fileKeyword) {
  var sheet = getVoiceSheet() || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  for (var i = 3; i <= lastRow; i++) {
    var charName = (sheet.getRange(i, 3).getValue() || "").toLowerCase();
    var folder = (sheet.getRange(i, 1).getValue() || "").toLowerCase();
    var fileName = (sheet.getRange(i, 2).getValue() || "").toLowerCase();
    var visible = true;
    if (character && !charName.includes(character.toLowerCase())) visible = false;
    if (folderKeyword && !folder.includes(folderKeyword.toLowerCase())) visible = false;
    if (fileKeyword && !fileName.includes(fileKeyword.toLowerCase())) visible = false;
    sheet.showRows(i);
    if (!visible) sheet.hideRows(i);
  }
}

// ===========================
// Bulk update helpers
// ===========================
function getFilesByCharacters(charNamesArray) {
  var sheet = getVoiceSheet() || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var out = [];
  if (lastRow < 3 || !charNamesArray || charNamesArray.length === 0) return out;
  for (var i = 3; i <= lastRow; i++) {
    var rowChar = sheet.getRange(i, 3).getValue();
    if (charNamesArray.indexOf(rowChar) !== -1) {
      var fileName = sheet.getRange(i, 2).getValue();
      var fileUrl = getUrlFromCell(sheet.getRange(i, 6));
      out.push({ fileName: fileName, url: fileUrl });
    }
  }
  return out;
}

function applyBulkLinkAndStatus(charNames, fileUrls, link, status) {
  var sheet = getVoiceSheet() || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return 0;

  charNames = Array.isArray(charNames) ? charNames : (charNames ? [charNames] : []);
  fileUrls = Array.isArray(fileUrls) ? fileUrls : (fileUrls ? [fileUrls] : []);

  var urlSet = {};
  for (var u = 0; u < fileUrls.length; u++) if (fileUrls[u]) urlSet[fileUrls[u]] = true;
  var filterByFiles = Object.keys(urlSet).length > 0;

  var updated = 0;
  for (var i = 3; i <= lastRow; i++) {
    var rowChar = sheet.getRange(i, 3).getValue();
    if (charNames.indexOf(rowChar) === -1) continue;
    var rowFileUrl = getUrlFromCell(sheet.getRange(i, 6));
    if (filterByFiles && !urlSet[rowFileUrl]) continue;

    if (link && link.trim() !== "") sheet.getRange(i, 7).setFormula('=HYPERLINK("' + link + '","Final Production")');
    if (status && status.trim() !== "") sheet.getRange(i, 8).setValue(status);
    updated++;
  }
  try { syncAuditionsWithProjects(sheet, getAuditionSheet()); } catch (e) { Logger.log(e); }
  return updated;
}

function applyLinkAndStatusForCharacter(charName, fileUrl, link, status) {
  if (!charName) throw new Error("No character selected.");
  var sheet = getVoiceSheet() || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) throw new Error("No data rows available.");

  var updated = 0;
  for (var i = 3; i <= lastRow; i++) {
    var rowChar = sheet.getRange(i, 3).getValue();
    var rowFileUrl = getUrlFromCell(sheet.getRange(i, 6));
    if (rowChar === charName && (!fileUrl || (rowFileUrl && rowFileUrl.indexOf(fileUrl) !== -1))) {
      if (link && link.trim() !== "") sheet.getRange(i, 7).setFormula('=HYPERLINK("' + link + '","Final Production")');
      if (status && status.trim() !== "") sheet.getRange(i, 8).setValue(status);
      updated++;
    }
  }
  if (updated === 0) throw new Error("No matching rows found to update.");
  try { syncAuditionsWithProjects(sheet, getAuditionSheet()); } catch (e) { Logger.log(e); }
  return updated;
}

// ===========================
// CSV export & utilities
// ===========================
function exportVisibleAuditionsCSV() { return exportVisibleToCSV_(getAuditionSheet(), 'Auditions_Visible.csv', 4); }
function exportVisibleVoiceCSV() {
  var sh = getVoiceSheet() || SpreadsheetApp.getActiveSheet();
  return exportVisibleToCSV_(sh, 'Voice_Visible.csv', 3);
}
function exportVisibleToCSV_(sheet, filename, dataStartRow) {
  var lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) throw new Error('No data to export.');
  var lastCol = sheet.getLastColumn();
  var rows = [];
  rows.push(sheet.getRange(2, 1, 1, lastCol).getDisplayValues()[0]); // headers
  for (var r = dataStartRow; r <= lastRow; r++) {
    if (sheet.isRowHiddenByUser(r) || sheet.isRowHiddenByFilter(r)) continue;
    rows.push(sheet.getRange(r, 1, 1, lastCol).getDisplayValues()[0]);
  }
  if (rows.length <= 1) throw new Error('No visible rows to export.');

  var csv = rows.map(function(row) {
    return row.map(function(v) {
      v = (v == null ? '' : String(v));
      return '"' + v.replace(/"/g, '""') + '"';
    }).join(',');
  }).join('\r\n');

  var blob = Utilities.newBlob(csv, 'text/csv', filename);
  var s = getSettings();
  if (!s.voiceFolderId) throw new Error('Set Voice folder in Settings first.');
  DriveApp.getFolderById(s.voiceFolderId).createFile(blob);
  return filename;
}

// ===========================
// Triggers, menu & PDF
// ===========================
function installAutoRefresh() {
  removeAutoRefresh();
  ScriptApp.newTrigger('updateVoiceProjects').timeBased().everyHours(1).create();
}
function removeAutoRefresh() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'updateVoiceProjects') ScriptApp.deleteTrigger(t);
  });
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Voice Tracker')
    .addItem('Refresh Voice', 'updateVoiceProjects')
    .addItem('Refresh Auditions', 'updateAuditions')
    .addItem('Sync Auditions from Voice', 'syncNow')
    .addItem('Show Tracker Panel', 'showPopup')
    .addItem('Install Auto-Refresh (hourly)', 'installAutoRefresh')
    .addItem('Remove Auto-Refresh', 'removeAutoRefresh')
    .addItem('Export Visible Auditions CSV', 'exportVisibleAuditionsCSV')
    .addItem('Export Visible Voice CSV', 'exportVisibleVoiceCSV')
    .addItem('Export Summary PDF', 'exportSummaryPDF')
    .addToUi();
}

function onEdit(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() === 'Auditions') return; // let manual edits live
  if (sheet.getName() !== 'Auditions' && e.range.getColumn() === 8 && e.range.getRow() >= 3) {
    try { syncAuditionsWithProjects(getVoiceSheet(), getAuditionSheet()); } catch (err) { Logger.log(err); }
  }
}

function showPopup() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Voice Project Tracker');
  SpreadsheetApp.getUi().showSidebar(html);
}

function exportSummaryPDF() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = getSettings();
  if (!s.voiceFolderId) throw new Error('Set Voice folder in Settings first.');
  var folder = DriveApp.getFolderById(s.voiceFolderId);
  var pdfName = "VoiceProjectSummary_" + new Date().toISOString().slice(0,10) + ".pdf";
  var blob = ss.getAs('application/pdf').setName(pdfName);
  folder.createFile(blob);
  return pdfName;
}
