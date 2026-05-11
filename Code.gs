/**
 * KAL ACADEMY - Smart File Registry
 * Updated: Description extractor now handles CamelCase and filters out metadata codes.
 *
 * Column layout (Docs sheet):
 *  A(1)  Row number
 *  B(2)  Human-Readable Description
 *  C(3)  File Name  [edit only this column]
 *  D(4)  File Type
 *  E(5)  Current Version
 *  F(6)  Current Folder
 *  G(7)  Link
 *  H(8)  For Who
 *  I(9)  KAL Name Conversion Check
 *  J(10) Destination Drive
 *  K(11) Preferred KAL Template  [dropdown]
 *  L(12) Abstract               [rightmost / AI-generated]
 */

// ── Column indices (1-based) ──────────────────────────────────────────────────
const COL = Object.freeze({
  ROW_NUM:    1,
  DESC:       2,
  FILENAME:   3,
  FILETYPE:   4,
  VERSION:    5,
  FOLDER:     6,
  LINK:       7,
  FOR_WHO:    8,
  KAL_CHECK:  9,
  DEST_DRIVE: 10,
  TEMPLATE:   11,
  ABSTRACT:   12
});

const LAST_COL   = COL.ABSTRACT; // rightmost data column
const DATA_START = 2;            // row 1 = header; data begins on row 2

// ── Lifecycle ─────────────────────────────────────────────────────────────────

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💠 KAL File System')
    .addItem('🎯 Audit & Sync Selected File', 'updateSelectedInfo')
    .addItem('🔄 Audit & Sync All Files',     'updateAllInfo')
    .addSeparator()
    .addItem('🏗️ Create Selected File',   'createSelectedFile')
    .addItem('🗑️ Remove Selected Version', 'removeSelectedFile')
    .addItem('☢️ Remove All Versions',     'removeAllVersions')
    .addSeparator()
    .addItem('🔧 Rebuild Header Row',  'buildHeaderRow')
    .addItem('📖 Show The User Guide', 'showUserGuide')
    .addToUi();
  try { updateTemplateDropdown(); } catch (_) { /* non-fatal on open */ }
}

// ── Header builder ────────────────────────────────────────────────────────────

/**
 * Writes the hardcoded header row (row 1) with formatting and the KAL logo.
 * Safe to re-run: it always overwrites row 1.
 */
function buildHeaderRow() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // ── Row geometry ──────────────────────────────────────────────────────────
  sheet.setRowHeight(1, 64);

  // ── Header labels (col A left blank — logo is inserted separately) ────────
  const labels = [[
    '',                           // A  → logo
    'Human-Readable\nDescription',                                                                 // B
    'File Name (Full Naming Guide)\n\n[DRIVE]-[ENTITY]_[DOCTYPE]_[Human-Readable CamelCase Description]', // C
    'File Type',                  // D
    'Current\nVersion',           // E
    'Current\nFolder',            // F
    'Link',                       // G
    'For\nWho',                   // H
    'KAL\nName Conversion\nCheck',// I
    'Destination\nDrive',         // J
    'Preferred\nKAL Template',    // K
    'Abstract'                    // L
  ]];

  const headerRange = sheet.getRange(1, 1, 1, LAST_COL);
  headerRange.setValues(labels);

  // ── Shared formatting ─────────────────────────────────────────────────────
  headerRange
    .setBackground('#1155CC')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  // ── Col A: Kiji logo (cell-bound image from KAL_LOGO_BASE64) ─────────────
  try {
    const cellImage = SpreadsheetApp.newCellImage()
      .setSourceUrl('data:image/png;base64,' + KAL_LOGO_BASE64)
      .setAltTextDescription('KAL Logo')
      .build();
    sheet.getRange(1, COL.ROW_NUM).setValue(cellImage);
  } catch (e) {
    console.error('buildHeaderRow logo: ' + e.message);
    // Fallback: insert as a floating image anchored to A1
    try {
      const blob = Utilities.newBlob(Utilities.base64Decode(KAL_LOGO_BASE64), 'image/png', 'kal_logo.png');
      sheet.insertImage(blob, COL.ROW_NUM, 1, 4, 4);
    } catch (e2) {
      console.error('buildHeaderRow logo fallback: ' + e2.message);
    }
  }

  ui.alert('✅ Header row rebuilt successfully!');
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/** Returns the URL from a cell by reading rich-text link first, then plain value. */
function getUrlFromCell(sheetName, cellAddress) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return null;
    const cell = sheet.getRange(cellAddress);
    const url  = cell.getRichTextValue().getLinkUrl() || cell.getValue().toString().trim();
    return url || null;
  } catch (_) { return null; }
}

/** Extracts a Google Drive file/folder ID from a URL. */
function getIdFromUrl(url) {
  if (!url) return null;
  const m = url.match(/[-\w]{25,}/);
  return m ? m[0] : null;
}

/** Maps a MIME type to a short human-readable label. */
function formatMimeType(mime) {
  const MAP = {
    'application/vnd.google-apps.document':     'G-Doc',
    'application/vnd.google-apps.spreadsheet':  'G-Sheet',
    'application/vnd.google-apps.presentation': 'G-Slide',
    'application/pdf': 'PDF'
  };
  return MAP[mime] || 'File';
}

// ── 1. DROPDOWN FETCHER ───────────────────────────────────────────────────────

function getTemplateList() {
  const folderId = getIdFromUrl(getUrlFromCell('Settings', 'B2'));
  if (!folderId) return [];
  try {
    const names = [];
    const files = DriveApp.getFolderById(folderId).getFiles();
    while (files.hasNext()) names.push(files.next().getName());
    return names;
  } catch (e) {
    console.error('getTemplateList: ' + e.message);
    return [];
  }
}

// ── 2. CORE AUDIT ENGINE ──────────────────────────────────────────────────────

/**
 * Audits one row and writes results to the sheet.
 *
 * @param {Sheet}    sheet
 * @param {number}   r                 1-based row index
 * @param {Object}   driveUrlLookup    DRIVE_CODE → folder URL
 * @param {Set}      validEntities
 * @param {Set}      validDocs
 * @param {string[]} templateList
 * @param {string}   [preloadedName]   pre-fetched col C value (avoids an extra read in batch mode)
 * @param {boolean}  [applyBg=true]    set false in batch mode; caller batch-writes backgrounds
 * @returns {string|null}              background colour for this row
 */
function processAuditForRow(sheet, r, driveUrlLookup, validEntities, validDocs, templateList, preloadedName, applyBg) {
  if (applyBg === undefined) applyBg = true;

  const baseName = (preloadedName !== undefined)
    ? preloadedName
    : sheet.getRange(r, COL.FILENAME).getValue().toString().trim();

  const dropdownCell = sheet.getRange(r, COL.TEMPLATE);

  // ── Empty row: clear everything and exit ────────────────────────────────
  if (!baseName) {
    sheet.getRange(r, COL.ROW_NUM, 1, 2).clearContent();                        // A–B
    sheet.getRange(r, COL.FILETYPE, 1, LAST_COL - COL.FILETYPE + 1).clearContent(); // D–L
    dropdownCell.clearDataValidations();
    if (applyBg) sheet.getRange(r, COL.DESC, 1, LAST_COL - COL.DESC + 1).setBackground(null);
    return null;
  }

  // ── Smart description extractor (CamelCase splitter) ────────────────────
  const nameParts = baseName.split(/[-_]/);

  if (nameParts.length >= 4) {
    const raw   = nameParts.slice(3).join('');
    const clean = raw
      .replace(/([a-z])([A-Z])/g, '$1 $2')
      .replace(/([A-Z])([A-Z][a-z])/g, '$1 $2')
      .trim();
    sheet.getRange(r, COL.DESC).setValue(clean);
  }

  const driveRaw   = (nameParts[0] || '').trim();
  const entityRaw  = (nameParts[1] || '').trim();
  const docTypeRaw = (nameParts[2] || '').trim();

  const driveCode   = driveRaw.toUpperCase();
  const entityCode  = entityRaw.toUpperCase();
  const docTypeCode = docTypeRaw.toUpperCase();

  // ── Validation (O(1) Set lookups) ───────────────────────────────────────
  const diagnostics = [];

  if (!driveUrlLookup[driveCode]) {
    diagnostics.push('Invalid Drive Code');
  } else if (driveRaw !== driveCode) {
    diagnostics.push('Drive Code must be UPPERCASE');
  }

  if (!validEntities.has(entityCode)) {
    diagnostics.push('Unregistered Entity');
  } else if (entityRaw !== entityCode) {
    diagnostics.push('Entity Code must be UPPERCASE');
  }

  if (!validDocs.has(docTypeCode)) {
    diagnostics.push('Invalid DocType');
  } else if (docTypeRaw !== docTypeCode) {
    diagnostics.push('DocType must be UPPERCASE');
  }

  const status = diagnostics.length > 0 ? diagnostics.join(' | ') : 'OK';
  sheet.getRange(r, COL.KAL_CHECK).setValue(status);

  // ── Col J: Destination Drive ─────────────────────────────────────────────
  const driveUrl = driveUrlLookup[driveCode];
  if (driveUrl) {
    sheet.getRange(r, COL.DEST_DRIVE)
      .setFormula('=HYPERLINK("' + driveUrl + '", "' + driveCode + ' Drive")');
  } else {
    sheet.getRange(r, COL.DEST_DRIVE).clearContent();
  }

  // ── Drive file lookup ────────────────────────────────────────────────────
  const info = GET_SMART_DETAILS(baseName);
  sheet.getRange(r, COL.FILETYPE, 1, COL.FOR_WHO - COL.FILETYPE + 1).clearContent(); // D–H

  if (info.fileLink !== 'Not Found') {
    // Batch-write plain values D and E in one call
    sheet.getRange(r, COL.FILETYPE, 1, 2).setValues([[info.type, info.version]]);
    sheet.getRange(r, COL.FOLDER)
      .setFormula('=HYPERLINK("' + info.folderLink + '", "' + info.folderName + '")');
    sheet.getRange(r, COL.LINK)
      .setFormula('=HYPERLINK("' + info.fileLink + '", "Link")');
    sheet.getRange(r, COL.FOR_WHO).setValue(entityCode);
    dropdownCell.clearContent().clearDataValidations();
  } else {
    sheet.getRange(r, COL.LINK).setValue('File not found');
    if (templateList && templateList.length > 0) {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(templateList)
        .setAllowInvalid(false)
        .build();
      dropdownCell.setDataValidation(rule);
    }
  }

  // ── Col L: Abstract (AI-generated summary) ───────────────────────────────
  sheet.getRange(r, COL.ABSTRACT).setFormula(
    '=AI("Based ONLY on description \'"&B' + r + '&"\' and filename \'"&C' + r + '&"\', write a two-sentence summary.")'
  );

  // ── Row background ───────────────────────────────────────────────────────
  const color = status !== 'OK' ? '#f4cccc' : (info.version === 'FINAL' ? '#d9ead3' : null);
  if (applyBg) sheet.getRange(r, COL.DESC, 1, LAST_COL - COL.DESC + 1).setBackground(color);
  return color;
}

// ── 3. SYNC LOGIC ─────────────────────────────────────────────────────────────

/**
 * Reads the Levels sheet once and returns lookup structures.
 * Uses Sets for O(1) entity/docType validation.
 *
 * @returns {{ driveUrlLookup: Object, validEntities: Set, validDocs: Set }}
 */
function getLevelsData() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const levelsSheet = ss.getSheetByName('Levels');
  if (!levelsSheet) throw new Error('Sheet "Levels" not found. Check your spreadsheet setup.');

  const lastRow = levelsSheet.getLastRow();
  if (lastRow < 2) return { driveUrlLookup: {}, validEntities: new Set(), validDocs: new Set() };

  // Single range read covering both values and rich-text in one fetch
  const range     = levelsSheet.getRange(1, 1, lastRow, 7);
  const lvValues  = range.getValues();
  const lvRich    = range.getRichTextValues();

  const driveUrlLookup = {};
  const validEntities  = new Set();
  const validDocs      = new Set();

  for (let j = 1; j < lvValues.length; j++) {
    const driveCode = String(lvValues[j][0]).toUpperCase().trim();
    const hiddenUrl = lvRich[j][1] ? lvRich[j][1].getLinkUrl() : null;
    if (driveCode && hiddenUrl) driveUrlLookup[driveCode] = hiddenUrl;

    const entity = String(lvValues[j][3]).toUpperCase().trim();
    if (entity) validEntities.add(entity);

    const docType = String(lvValues[j][6]).toUpperCase().trim();
    if (docType) validDocs.add(docType);
  }

  return { driveUrlLookup, validEntities, validDocs };
}

/**
 * Audits every data row.
 * Performance strategy:
 *  - Reads the entire filename column (C) in one batch API call before the loop.
 *  - Passes each baseName into processAuditForRow to skip a per-row read.
 *  - Collects background colours and row numbers, then writes them in two
 *    batch API calls after the loop instead of N individual calls.
 */
function updateAllInfo() {
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) return;

  let data, templateList;
  try {
    data         = getLevelsData();
    templateList = getTemplateList();
  } catch (e) {
    SpreadsheetApp.getUi().alert('⚠️ Setup error: ' + e.message);
    return;
  }

  const numRows   = lastRow - DATA_START + 1;
  // Batch-read all filenames in col C (one API call)
  const fileNames = sheet.getRange(DATA_START, COL.FILENAME, numRows, 1).getValues();

  const bgColors = []; // 2-D array for batch setBackgrounds
  const rowNums  = []; // 2-D array for batch setValues (col A)
  let errors = 0;

  for (let i = 0; i < numRows; i++) {
    const r        = DATA_START + i;
    const baseName = String(fileNames[i][0]).trim();
    try {
      // applyBg = false: we batch-write backgrounds after the loop
      const color = processAuditForRow(
        sheet, r,
        data.driveUrlLookup, data.validEntities, data.validDocs,
        templateList, baseName, false
      );
      const numBgCols = LAST_COL - COL.DESC + 1;
      bgColors.push(Array(numBgCols).fill(color));
      rowNums.push([baseName ? r - 1 : '']);
    } catch (e) {
      console.error('Row ' + r + ': ' + e.message);
      bgColors.push(Array(LAST_COL - COL.DESC + 1).fill('#fff2cc')); // amber = script error
      rowNums.push(['?']);
      errors++;
    }
  }

  // Two batch writes replace N×11 individual setBackground calls + N setValue calls
  sheet.getRange(DATA_START, COL.DESC, numRows, LAST_COL - COL.DESC + 1).setBackgrounds(bgColors);
  sheet.getRange(DATA_START, COL.ROW_NUM, numRows, 1).setValues(rowNums);

  if (errors > 0) {
    SpreadsheetApp.getUi().alert(
      '⚠️ ' + errors + ' row(s) hit errors. Open View → Logs for details.'
    );
  }
}

function updateSelectedInfo() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const r     = sheet.getActiveRange().getRow();
    if (r < DATA_START) return;
    const data         = getLevelsData();
    const templateList = getTemplateList();
    processAuditForRow(sheet, r, data.driveUrlLookup, data.validEntities, data.validDocs, templateList);
  } catch (e) {
    SpreadsheetApp.getUi().alert('⚠️ Audit error: ' + e.message);
  }
}

// ── 4. DRIVE SEARCH ENGINE ────────────────────────────────────────────────────

/**
 * Finds the latest versioned Drive file whose name starts with baseName.
 *
 * Performance: regex is compiled once before the iteration, not inside the loop.
 *
 * @param {string} baseName
 * @returns {{ type, version, folderName, folderLink, fileLink }}
 */
function GET_SMART_DETAILS(baseName) {
  const NOT_FOUND = {
    type: '', version: 'Not Found',
    folderName: 'Not Found', folderLink: '', fileLink: 'Not Found'
  };
  if (!baseName) return NOT_FOUND;

  // Compile regex once outside the while loop
  const escaped   = baseName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const versionRe = new RegExp(escaped + '.*[vV](\\d+|FINAL)', 'i');

  let latestVerNum = -1;
  const details    = Object.assign({}, NOT_FOUND);

  try {
    const files = DriveApp.searchFiles("title contains '" + baseName + "'");
    while (files.hasNext()) {
      const file  = files.next();
      const match = file.getName().match(versionRe);
      if (!match) continue;

      const vSuffix = match[1].toUpperCase();
      const vNum    = vSuffix === 'FINAL' ? 9999 : parseInt(vSuffix, 10);
      if (vNum <= latestVerNum) continue;

      latestVerNum     = vNum;
      details.version  = vSuffix;
      details.fileLink = file.getUrl();
      details.type     = formatMimeType(file.getMimeType());

      const parents = file.getParents();
      if (parents.hasNext()) {
        const p = parents.next();
        details.folderName = p.getName();
        details.folderLink = p.getUrl();
      }
    }
  } catch (e) {
    console.error('GET_SMART_DETAILS("' + baseName + '"): ' + e.message);
  }

  return details;
}

// ── 5. FILE MANAGEMENT ────────────────────────────────────────────────────────

function createSelectedFile() {
  const ui    = SpreadsheetApp.getUi();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const r     = sheet.getActiveRange().getRow();
  if (r < DATA_START) return;

  const baseName     = sheet.getRange(r, COL.FILENAME).getValue().toString().trim();
  const templateName = sheet.getRange(r, COL.TEMPLATE).getValue().toString().trim();

  if (!baseName) { ui.alert('🛑 No filename found in this row.'); return; }

  let data, templateList;
  try {
    data         = getLevelsData();
    templateList = getTemplateList();
  } catch (e) {
    ui.alert('⚠️ Setup error: ' + e.message);
    return;
  }

  processAuditForRow(sheet, r, data.driveUrlLookup, data.validEntities, data.validDocs, templateList);
  const nameStatus = sheet.getRange(r, COL.KAL_CHECK).getValue().toString().trim();

  if (nameStatus !== 'OK') {
    ui.alert('🛑 File name has issues:\n' + nameStatus);
    return;
  }
  if (!templateName) {
    ui.alert('🛑 Select a template from the dropdown (col K) first.');
    return;
  }
  if (GET_SMART_DETAILS(baseName).fileLink !== 'Not Found') {
    ui.alert('🛑 File already exists in Drive.');
    return;
  }

  try {
    const destId = getIdFromUrl(getUrlFromCell('Settings', 'A2'));
    if (!destId) throw new Error('Destination folder URL missing or invalid in Settings!A2.');

    const tempFolderId = getIdFromUrl(getUrlFromCell('Settings', 'B2'));
    if (!tempFolderId) throw new Error('Template folder URL missing or invalid in Settings!B2.');

    const dateStr       = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd');
    const finalFileName = baseName + '_' + dateStr + '_v1';

    const templateFiles = DriveApp.getFolderById(tempFolderId).getFilesByName(templateName);
    if (!templateFiles.hasNext()) {
      ui.alert('🛑 Template "' + templateName + '" not found in the templates folder.');
      return;
    }
    templateFiles.next().makeCopy(finalFileName, DriveApp.getFolderById(destId));
    updateSelectedInfo();
    ui.alert('✅ "' + finalFileName + '" created successfully!');
  } catch (e) {
    ui.alert('❌ Creation failed: ' + e.message);
  }
}

function removeSelectedFile() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const r     = sheet.getActiveRange().getRow();

  let fileUrl = null;
  try { fileUrl = sheet.getRange(r, COL.LINK).getRichTextValue().getLinkUrl(); } catch (_) {}
  if (!fileUrl) { ui.alert('🛑 No file link found in this row.'); return; }

  const fileId = getIdFromUrl(fileUrl);
  if (!fileId) { ui.alert('🛑 Could not parse a file ID from the link.'); return; }

  const response = ui.alert('⚠️ Warning', 'Move THIS version to Trash?', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    updateSelectedInfo();
  } catch (e) {
    ui.alert('❌ Could not trash the file: ' + e.message);
  }
}

function removeAllVersions() {
  const ui       = SpreadsheetApp.getUi();
  const sheet    = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const r        = sheet.getActiveRange().getRow();
  const baseName = sheet.getRange(r, COL.FILENAME).getValue().toString().trim();
  if (!baseName) return;

  const response = ui.alert(
    '☢️ NUCLEAR WARNING',
    'Trash EVERY version of: ' + baseName,
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  let trashed = 0, failed = 0;
  try {
    const files = DriveApp.searchFiles("title contains '" + baseName + "'");
    while (files.hasNext()) {
      try {
        files.next().setTrashed(true);
        trashed++;
      } catch (e) {
        console.error('removeAllVersions trash: ' + e.message);
        failed++;
      }
    }
  } catch (e) {
    ui.alert('❌ Drive search failed: ' + e.message);
    return;
  }

  updateSelectedInfo();
  ui.alert('Trashed ' + trashed + ' file(s).' + (failed > 0 ? ' (' + failed + ' failed — check Logs)' : ''));
}

// ── 6. UI ─────────────────────────────────────────────────────────────────────

function showUserGuide() {
  try {
    const template = HtmlService.createTemplateFromFile('Sidebar');
    template.logoBase64 = KAL_LOGO_BASE64;
    const html = template.evaluate().setTitle('KAL File System').setWidth(350);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Could not load the User Guide: ' + e.message);
  }
}

// ── 7. ACADEMY SIDEBAR DATA ───────────────────────────────────────────────────

function getAcademyCodes() {
  const EMPTY = { drives: [], entities: [], docTypes: [], examples: [] };
  try {
    const levelsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Levels');
    if (!levelsSheet) return EMPTY;

    const lastRow = levelsSheet.getLastRow();
    if (lastRow < 3) return EMPTY;

    const values   = levelsSheet.getRange(3, 1, lastRow - 2, 10).getValues();
    const drives   = [];
    const entities = [];
    const docTypes = [];
    const examples = [];

    values.forEach(row => {
      if (row[0]) drives.push({ code: row[0], name: row[2] });
      if (row[3]) entities.push({ code: row[3], name: row[4] });
      if (row[6]) docTypes.push({ code: row[6], name: row[7] });
      if (row[8]) examples.push({ input: row[8], output: row[9] });
    });

    return { drives, entities, docTypes, examples };
  } catch (e) {
    console.error('getAcademyCodes: ' + e.message);
    return EMPTY;
  }
}
