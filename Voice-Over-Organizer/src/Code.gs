// @ts-nocheck
// ========== MAIN SCRIPT ==========
//
// 🔁 Update sheet from Drive folder
//
// =====================
// Configuration (GLOBAL)
// =====================
var VOICE_FOLDER_ID = '1lq_rbzqRj_keQ3XXwtWUYG7z6hNGTxK7'; // main voice projects folder
var AUDITION_FOLDER_ID = '1YtKGcqfCC3zR5XriayoMw-SpvmA2Exmk'; // auditions folder

// Header fingerprints (to locate sheets reliably)
var VOICE_HEADERS = ["Folder", "File Name", "Character", "Date Added", "Mime Type", "File Link", "Final Link", "Status"];
var AUDITION_HEADERS = ["Folder", "File Name", "Character", "Date Added", "Mime Type", "File Link", "Status"];

// ====== Utilities ======
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
  // Fallback to active sheet if it looks like VOICE (has Final Link + Status in G/H)
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

// Extract URL from a cell (works for HYPERLINK() / RichText / raw)
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

// 🔎 Collect files recursively
function collectFiles(folder, path, filesData, allCharacters) {
  var folderName = folder.getName();
  var currentPath = path ? path + "/" + folderName : folderName;

  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    var match = fileName.match(/Fred French - (.+?) -/);
    var characterName = match ? match[1].trim() : "Unknown";
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
    collectFiles(subfolders.next(), currentPath, filesData, allCharacters);
  }
}

// =====================
// Voice Projects (OG)
// =====================
function updateVoiceProjects() {
  var folderId = VOICE_FOLDER_ID;
  var sheet = getVoiceSheet() || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // === Preserve G/H keyed by Character|FileName (from existing sheet)
  var existingData = {};
  var oldLastRow = sheet.getLastRow();
  if (oldLastRow >= 3) {
    var oldVals = sheet.getRange(3, 1, oldLastRow - 2, 8).getValues();   // A..H
    var oldForm = sheet.getRange(3, 7, oldLastRow - 2, 1).getFormulas(); // G (formulas)
    for (var i = 0; i < oldVals.length; i++) {
      var fileName = oldVals[i][1];       // B
      var character = oldVals[i][2];      // C
      var finalLinkValue = oldVals[i][6]; // G (value if no formula)
      var statusValue = oldVals[i][7];    // H
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

  // === Collect fresh Drive data
  var allCharacters = new Set();
  var filesData = [];
  collectFiles(DriveApp.getFolderById(folderId), "", filesData, allCharacters);

  // === Ensure enough rows
  var requiredRows = filesData.length + 2; // header row 2 + data from row 3
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }

  // === Clear old data block (A:H) to avoid leftovers
  if (oldLastRow >= 3) {
    var clearCount = oldLastRow - 2;
    sheet.getRange(3, 1, clearCount, 8).clearContent().clearFormat();
  }

  // === Title & headers (unchanged)
  var titleRange = sheet.getRange(1, 1, 1, 8);
  try { if (!titleRange.isPartOfMerge()) titleRange.merge(); } catch (e) { try { titleRange.merge(); } catch(e2) {} }
  titleRange.setValue("🎙️ Voice Projects Overview")
    .setFontWeight("bold").setFontSize(14).setFontFamily("Roboto")
    .setBackground("#81C784").setFontColor("#000000")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  var VOICE_HEADERS = ["Folder", "File Name", "Character", "Date Added", "Mime Type", "File Link", "Final Link", "Status"];
  sheet.getRange(2, 1, 1, VOICE_HEADERS.length).setValues([VOICE_HEADERS])
    .setFontWeight("bold").setFontFamily("Roboto").setFontSize(10)
    .setBackground("#C8E6C9").setFontColor("#000000")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, false, false, "#A5D6A7", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // === Build rows A:H in memory (formulas inline)
  var out = new Array(filesData.length);
  for (var i = 0; i < filesData.length; i++) {
    var d = filesData[i];
    var key = d.characterName + '|' + d.fileName;
    var store = existingData[key] || {};

    // Column G: keep prior formula if present; else keep prior value; if that value is a URL, convert to hyperlink for consistency
    var gCell = '';
    if (store.finalLinkFormula) {
      gCell = store.finalLinkFormula; // already a formula
    } else if (store.finalLinkValue && typeof store.finalLinkValue === 'string') {
      var v = store.finalLinkValue.trim();
      if (/^https?:\/\//i.test(v)) gCell = '=HYPERLINK("' + v + '","Final Production")';
      else gCell = v;
    }

    // Column H: keep prior status or default "Current"
    var hCell = store.status ? store.status : "Current";

    out[i] = [
      d.folderPath,                              // A
      d.fileName,                                // B
      d.characterName,                           // C
      d.dateAdded,                               // D
      d.mimeType,                                // E
      '=HYPERLINK("' + d.url + '","🎵 View File")', // F (formula)
      gCell,                                     // G (formula or value)
      hCell                                      // H
    ];
  }

  // === Batch write all rows in one call
  if (filesData.length > 0) {
    sheet.getRange(3, 1, filesData.length, 8).setValues(out);
  }

  // === Format the written block once
  if (filesData.length > 0) {
    var dataRange = sheet.getRange(3, 1, filesData.length, 8);
    dataRange
      .setFontFamily("Roboto").setFontSize(10)
      .setHorizontalAlignment("left").setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true, "#E0E0E0", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(3, 4, filesData.length, 1).setHorizontalAlignment("center"); // Date
    sheet.getRange(3, 8, filesData.length, 1).setHorizontalAlignment("center"); // Status

    // Alternating row colors in one go
    var bg = [];
    for (var r = 0; r < filesData.length; r++) {
      var color = (r % 2 === 0) ? "#F1F8E9" : "#FFFFFF";
      var rowColors = new Array(8).fill(color);
      bg.push(rowColors);
    }
    dataRange.setBackgrounds(bg);

    // Column widths (single pass)
    sheet.autoResizeColumns(1, 8);
    var minWidths = [150, 200, 130, 100, 120, 150, 150, 100];
    for (var col = 1; col <= 8; col++) {
      if (sheet.getColumnWidth(col) < minWidths[col - 1]) {
        sheet.setColumnWidth(col, minWidths[col - 1]);
      }
    }
    dataRange.setWrap(true);
  }

  // Save character list for UI
  PropertiesService.getDocumentProperties()
    .setProperty('CHARACTERS', JSON.stringify(Array.from(allCharacters)));

  // Sync auditions after refresh
  try { syncAuditionsWithProjects(sheet, getAuditionSheet()); } catch (e) { Logger.log(e); }
}

function getStatsAndRecent(limit) {
  limit = limit || 12;
  var now = new Date();
  var recent7  = new Date(now.getTime() - 7  * 24 * 60 * 60 * 1000);
  var recent30 = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);

  var out = {
    voice: { total: 0, current: 0, past: 0, recent7: 0, recent30: 0 },
    auditions: { total: 0, pending: 0, submitted: 0, booked: 0, passed: 0, recent7: 0, recent30: 0 },
    recent: [] // {type, date(ISO), character, fileName, status}
  };

  // --- Voice (OG)
  var voiceSheet = getVoiceSheet();
  if (voiceSheet) {
    var last = voiceSheet.getLastRow();
    if (last >= 3) {
      var vals = voiceSheet.getRange(3, 1, last - 2, 8).getValues(); // A..H
      for (var i = 0; i < vals.length; i++) {
        var row = vals[i];
        var fileName = row[1], character = row[2];
        var date = row[3]; // Date Added
        var status = (row[7] || '').toString();

        out.voice.total++;
        if (status === 'Current') out.voice.current++;
        else if (status === 'Past') out.voice.past++;

        if (date && date instanceof Date) {
          if (date >= recent7)  out.voice.recent7++;
          if (date >= recent30) out.voice.recent30++;
          out.recent.push({
            type: 'Voice',
            date: date.toISOString(),
            character: character || '',
            fileName: fileName || '',
            status: status || ''
          });
        }
      }
    }
  }

  // --- Auditions
  var audSheet = getAuditionSheet();
  if (audSheet) {
    var lastA = audSheet.getLastRow();
    if (lastA >= 4) {
      var valsA = audSheet.getRange(4, 1, lastA - 3, 7).getValues(); // A..G
      for (var j = 0; j < valsA.length; j++) {
        var rowA = valsA[j];
        var fileNameA = rowA[1], characterA = rowA[2];
        var dateA = rowA[3]; // Date Added
        var statA = (rowA[6] || '').toString();

        out.auditions.total++;
        var key = statA.toLowerCase();
        if (key === 'pending')   out.auditions.pending++;
        else if (key === 'submitted') out.auditions.submitted++;
        else if (key === 'booked')    out.auditions.booked++;
        else if (key === 'passed')    out.auditions.passed++;

        if (dateA && dateA instanceof Date) {
          if (dateA >= recent7)  out.auditions.recent7++;
          if (dateA >= recent30) out.auditions.recent30++;
          out.recent.push({
            type: 'Audition',
            date: dateA.toISOString(),
            character: characterA || '',
            fileName: fileNameA || '',
            status: statA || ''
          });
        }
      }
    }
  }

  // Sort recent by date desc and trim
  out.recent = out.recent
    .filter(function(x){ return x.date; })
    .sort(function(a,b){ return new Date(b.date) - new Date(a.date); })
    .slice(0, limit);

  return out;
}

// =====================
// Auditions
// =====================
function updateAuditions() {
  var sheet = getAuditionSheet();
  var folder = DriveApp.getFolderById(AUDITION_FOLDER_ID);

  // Preserve existing Status keyed by Character|FileName
  var preserve = {};
  var lastRow = sheet.getLastRow();
  if (lastRow >= 3) {
    var vals = sheet.getRange(3, 1, lastRow - 2, 7).getValues(); // A..G
    for (var i = 0; i < vals.length; i++) {
      var fileName = vals[i][1];   // B
      var charName = vals[i][2]||''; // C
      preserve[charName + '|' + fileName] = vals[i][6]; // G
    }
  }

  // Collect audition files
  var filesData = [];
  collectFiles(folder, "", filesData, new Set());

  // Reset layout (fast)
  sheet.setFrozenRows(0);
  sheet.clearContents().clearFormats();
  if (sheet.getMaxRows() < 4) sheet.insertRowsAfter(sheet.getMaxRows(), 4 - sheet.getMaxRows());
  var totalRows = sheet.getMaxRows();
  if (totalRows > 4) sheet.deleteRows(5, totalRows - 4);
  sheet.setFrozenRows(Math.min(3, sheet.getMaxRows() - 1));

  // Title & subtitle
  sheet.getRange("A1:G1").merge().setValue("🎧 Audition Tracker")
    .setFontSize(16).setFontWeight("bold").setFontFamily("Roboto")
    .setBackground("#2E7D32").setFontColor("#FFFFFF")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("A2:G2").merge().setValue("Automatically updated from your Google Drive audition folder")
    .setFontSize(10).setFontColor("#E8F5E9").setHorizontalAlignment("center");

  // Headers
  var AUDITION_HEADERS = ["Folder", "File Name", "Character", "Date Added", "Mime Type", "File Link", "Status"];
  sheet.getRange(3, 1, 1, AUDITION_HEADERS.length).setValues([AUDITION_HEADERS])
    .setFontFamily("Roboto").setFontSize(10).setFontWeight("bold")
    .setBackground("#A5D6A7").setFontColor("#1B5E20")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, false, false, "#81C784", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Build rows A..G in memory
  var out = new Array(filesData.length);
  for (var i = 0; i < filesData.length; i++) {
    var f = filesData[i];
    var key = f.characterName + '|' + f.fileName;
    var status = preserve[key] || "Pending";
    out[i] = [
      f.folderPath,                               // A
      f.fileName,                                 // B
      f.characterName,                            // C
      f.dateAdded,                                // D
      f.mimeType,                                 // E
      '=HYPERLINK("' + f.url + '","🎵 View Audition")', // F (formula)
      status                                      // G
    ];
  }

  // Batch write the data block
  if (filesData.length > 0) {
    sheet.getRange(4, 1, filesData.length, 7).setValues(out);
  }

  // Format in one pass
  if (filesData.length > 0) {
    var dataRange = sheet.getRange(4, 1, filesData.length, 7);
    dataRange
      .setFontFamily("Roboto").setFontSize(10)
      .setHorizontalAlignment("left").setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true, "#C8E6C9", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(4, 4, filesData.length, 1).setHorizontalAlignment("center");
    sheet.getRange(4, 7, filesData.length, 1).setHorizontalAlignment("center");

    var bg = [];
    for (var r = 0; r < filesData.length; r++) {
      var color = (r % 2 === 0) ? "#E8F5E9" : "#FFFFFF";
      bg.push(new Array(7).fill(color));
    }
    dataRange.setBackgrounds(bg);

    sheet.autoResizeColumns(1, 7);
  }

  // Conditional formatting (unchanged behavior)
  var statusRange = sheet.getRange(4, 7, Math.max(1, filesData.length), 1);
  var rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Pending").setBackground("#FFF9C4").setFontColor("#F57F17").setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Submitted").setBackground("#BBDEFB").setFontColor("#0D47A1").setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Booked").setBackground("#C8E6C9").setFontColor("#1B5E20").setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Passed").setBackground("#FFCDD2").setFontColor("#B71C1C").setRanges([statusRange]).build()
  ];
  sheet.setConditionalFormatRules(rules);

  // Dropdowns once (covers plenty of rows)
  addAuditionStatusDropdown();

  // Auto-sync after refresh
  try { syncAuditionsWithProjects(getVoiceSheet(), sheet); } catch (e) { Logger.log(e); }
}


function addAuditionStatusDropdown() {
  var sheet = getAuditionSheet();
  var lastRow = Math.max(4, sheet.getLastRow());
  var range = sheet.getRange(4, 7, lastRow - 3); // Column G
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Pending", "Submitted", "Booked", "Passed"], true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
}

// =======================
// Sync Auditions ⇄ Voice Projects
// =======================
function buildVoiceIndex_(voiceSheet) {
  var idx = {};
  if (!voiceSheet) return idx;
  var last = voiceSheet.getLastRow();
  if (last < 3) return idx;
  var vals = voiceSheet.getRange(3, 1, last - 2, 8).getValues(); // A..H
  for (var i = 0; i < vals.length; i++) {
    var fileName = vals[i][1]; // B
    var charName = vals[i][2]; // C
    var status   = vals[i][7]; // H
    if (!fileName || !charName) continue;
    idx[charName + '|' + fileName] = status;
  }
  return idx;
}

/**
 * Sync rule:
 *   Voice: Current  -> Auditions: Booked
 *   Voice: Past     -> Auditions: Submitted
 *   Never overwrite "Passed".
 * Returns {booked:n, submitted:n}
 */
function syncAuditionsWithProjects(voiceSheet, auditionSheet) {
  voiceSheet = voiceSheet || getVoiceSheet();
  auditionSheet = auditionSheet || getAuditionSheet();
  if (!voiceSheet || !auditionSheet) return { booked: 0, submitted: 0 };

  var map = buildVoiceIndex_(voiceSheet);
  var last = auditionSheet.getLastRow();
  if (last < 4) return { booked: 0, submitted: 0 };

  var booked = 0, submitted = 0;

  for (var r = 4; r <= last; r++) {
    var fileName = auditionSheet.getRange(r, 2).getValue(); // B
    var charName = auditionSheet.getRange(r, 3).getValue(); // C
    if (!fileName || !charName) continue;

    var key = charName + '|' + fileName;
    var voiceStatus = map[key];
    if (!voiceStatus) continue;

    var currentAud = auditionSheet.getRange(r, 7).getValue(); // G
    if (currentAud === "Passed") continue; // don't overwrite manual "Passed"

    if (voiceStatus === "Current") {
      if (currentAud !== "Booked") {
        auditionSheet.getRange(r, 7).setValue("Booked");
        booked++;
      }
    } else if (voiceStatus === "Past") {
      if (currentAud !== "Submitted") {
        auditionSheet.getRange(r, 7).setValue("Submitted");
        submitted++;
      }
    }
  }
  return { booked: booked, submitted: submitted };
}

// Handy button for UI
function syncNow() {
  return syncAuditionsWithProjects(getVoiceSheet(), getAuditionSheet());
}

// Filter + Quick Search for Auditions (status + text query)
// - status: "All" | "Pending" | "Submitted" | "Booked" | "Passed"
// - query: free text; matches Folder (A), File Name (B), Character (C)
function filterAuditionStatus(status, query) {
  var sheet = getAuditionSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return;

  status = (status || '').toString().trim();
  query  = (query  || '').toString().trim().toLowerCase();

  var allStatus = !status || status.toLowerCase() === 'all';
  var dataRows = lastRow - 3;
  if (dataRows <= 0) return;

  // Reset visibility first
  sheet.showRows(4, dataRows);

  // If no filtering needed, we're done
  if (allStatus && !query) return;

  // Read A..C (Folder, File, Character) and G (Status) in one shot
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

    // Hide in contiguous blocks (fewer API calls)
    if (!keep && hideStart === null) hideStart = 4 + i;
    if (keep && hideStart !== null) {
      sheet.hideRows(hideStart, (4 + i) - hideStart);
      hideStart = null;
    }
  }
  if (hideStart !== null) {
    sheet.hideRows(hideStart, (lastRow + 1) - hideStart);
  }
}

// =======================
// Bulk update Final Link & Status (Voice)
// =======================
function getFilesByCharacters(charNamesArray) {
  var sheet = getVoiceSheet() || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var out = [];
  if (lastRow < 3 || !charNamesArray || charNamesArray.length === 0) return out;

  for (var i = 3; i <= lastRow; i++) {
    var rowChar = sheet.getRange(i, 3).getValue(); // C
    if (charNamesArray.indexOf(rowChar) !== -1) {
      var fileName = sheet.getRange(i, 2).getValue(); // B
      var fileUrl = getUrlFromCell(sheet.getRange(i, 6)); // F
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
    var rowChar = sheet.getRange(i, 3).getValue(); // C
    if (charNames.indexOf(rowChar) === -1) continue;

    var rowFileUrl = getUrlFromCell(sheet.getRange(i, 6)); // F
    if (filterByFiles && !urlSet[rowFileUrl]) continue;

    if (link && link.trim() !== "") {
      sheet.getRange(i, 7).setFormula('=HYPERLINK("' + link + '","Final Production")');
    }
    if (status && status.trim() !== "") {
      sheet.getRange(i, 8).setValue(status);
    }
    updated++;
  }

  // After bulk updates, keep auditions in sync
  try { syncAuditionsWithProjects(sheet, getAuditionSheet()); } catch (e) { Logger.log(e); }
  return updated;
}

// Single-row updater
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

  // After single update, sync auditions
  try { syncAuditionsWithProjects(sheet, getAuditionSheet()); } catch (e) { Logger.log(e); }
  return updated;
}

// =======================
// Voice search + helpers
// =======================
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

function installAutoRefresh() {
  removeAutoRefresh();
  ScriptApp.newTrigger('updateVoiceProjects')
    .timeBased()
    .everyHours(1)          // tweak to .everyMinutes(15) if you prefer
    .create();
}

function removeAutoRefresh() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'updateVoiceProjects') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function autoHyperlinkFinalLink() {
  var sheet = getVoiceSheet() || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  var finalLinks = sheet.getRange(3, 7, lastRow - 2, 1).getValues(); // G
  for (var i = 0; i < finalLinks.length; i++) {
    var url = finalLinks[i][0];
    if (url && typeof url === 'string' && !/^=HYPERLINK\(/i.test(url) && /^https?:\/\//i.test(url)) {
      sheet.getRange(i + 3, 7).setFormula('=HYPERLINK("' + url + '","Final Production")');
    }
  }
}

// =======================
// Dashboard bits
// =======================
function getCharacters() {
  var chars = PropertiesService.getDocumentProperties().getProperty('CHARACTERS');
  return chars ? JSON.parse(chars) : [];
}

function updateSummary(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  var statusValues = sheet.getRange(3, 8, lastRow - 2, 1).getValues();
  var dateValues = sheet.getRange(3, 4, lastRow - 2, 1).getValues();
  var now = new Date();
  var recentThreshold = 7; // days

  var currentCount = 0;
  var pastCount = 0;
  var recentCount = 0;

  for (var i = 0; i < statusValues.length; i++) {
    var status = statusValues[i][0];
    var date = dateValues[i][0];
    if (status === "Current") currentCount++;
    else if (status === "Past") pastCount++;

    if (date) {
      var diffDays = (now - new Date(date)) / (1000 * 60 * 60 * 24);
      if (diffDays <= recentThreshold) recentCount++;
    }
  }

  sheet.getRange("A1:H1").merge();
  sheet.getRange("A1").setValue(
    "🎙️ Voice Project Tracker Dashboard | Current: " + currentCount + " | Past: " + pastCount + " | Recent: " + recentCount
  )
    .setBackground("#2E7D32")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(14)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
}

function exportVisibleAuditionsCSV() {
  return exportVisibleToCSV_(getAuditionSheet(), 'Auditions_Visible.csv', 4);
}
function exportVisibleVoiceCSV() {
  var sh = getVoiceSheet() || SpreadsheetApp.getActiveSheet();
  return exportVisibleToCSV_(sh, 'Voice_Visible.csv', 3);
}

function exportVisibleToCSV_(sheet, filename, dataStartRow) {
  var lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) throw new Error('No data to export.');
  var lastCol = sheet.getLastColumn();

  // Build rows: headers (row 2) + only visible data rows
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
  DriveApp.getFolderById(VOICE_FOLDER_ID).createFile(blob);
  return filename;
}

// =======================
// Menu + Hooks
// =======================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Voice Tracker')
    .addItem('Refresh Auditions', 'updateAuditions')
    .addItem('Sync Auditions from Voice', 'syncNow')
    .addItem('Show Tracker Panel', 'showPopup')
    .addItem('Refresh Sheet', 'updateVoiceProjects')
    .addItem('Export Summary PDF', 'exportSummaryPDF')
    .addItem('Install Auto-Refresh (hourly)', 'installAutoRefresh')
    .addItem('Remove Auto-Refresh', 'removeAutoRefresh')
    .addItem('Export Visible Auditions CSV', 'exportVisibleAuditionsCSV')
    .addItem('Export Visible Voice CSV', 'exportVisibleVoiceCSV')
    .addToUi();

}

function onEdit(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() === 'Auditions') return; // user can edit audition status; we don't auto-overwrite on edit

  // If user changes OG Status manually, push sync
  if (sheet.getName() !== 'Auditions' && e.range.getColumn() === 8 && e.range.getRow() >= 3) {
    try { syncAuditionsWithProjects(getVoiceSheet(), getAuditionSheet()); } catch (err) { Logger.log(err); }
  }
}

function showPopup() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Voice Project Tracker');
  SpreadsheetApp.getUi().showSidebar(html);
}

// =======================
// PDF Export (unchanged)
// =======================
function exportSummaryPDF() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var folder = DriveApp.getFolderById(VOICE_FOLDER_ID);
  var pdfName = "VoiceProjectSummary_" + new Date().toISOString().slice(0,10) + ".pdf";
  var blob = ss.getAs('application/pdf').setName(pdfName);
  folder.createFile(blob);
  return pdfName;
}
