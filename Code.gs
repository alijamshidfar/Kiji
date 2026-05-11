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
 *
 * NOTE: COL, LAST_COL, and DATA_START are defined in Constants.gs.
 */

// ── Lifecycle ─────────────────────────────────────────────────────────────────

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💠 KAL File System')
    .addItem('🎯 Audit & Sync Selected File',    'updateSelectedInfo')
    .addItem('🔄 Audit & Sync All Files',         'updateAllInfo')
    .addItem('🔍 Search For Missing KAL Files',   'searchMissingKALFiles')
    .addSeparator()
    .addItem('🏗️ Create Selected File',   'createSelectedFile')
    .addItem('🗑️ Remove Selected Version', 'removeSelectedFile')
    .addItem('☢️ Remove All Versions',     'removeAllVersions')
    .addSeparator()
    .addItem('📖 Show The User Guide', 'showUserGuide')
    .addToUi();
  try { updateTemplateDropdown(); } catch (_) { /* non-fatal on open */ }
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

// ── 3. MISSING FILE SEARCH ───────────────────────────────────────────────────

/**
 * Scans Google Drive for KAL files whose base names are not yet in the
 * registry, then inserts them directly below their matching drive-code
 * section in the sheet.
 *
 * Algorithm:
 *  1. Read drive codes (in sheet order) from the Levels tab.
 *  2. Batch-read col C once to build: (a) the set of registered base names,
 *     (b) the last row of each drive-code section.
 *  3. For each drive code, search Drive for files whose title starts with
 *     CODE- and matches the KAL naming pattern. Deduplicate by base name.
 *  4. Insert missing rows bottom-up (highest section row first) so earlier
 *     row indices stay valid throughout the insertion loop.
 *  5. After all insertions, run processAuditForRow on every new row so the
 *     full audit data (type, version, folder link, status …) is populated.
 */
function searchMissingKALFiles() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // ── 1. Drive codes (ordered as they appear in Levels) ────────────────────
  let driveCodes;
  try {
    driveCodes = getDriveCodesOrdered();
  } catch (e) {
    ui.alert('⚠️ ' + e.message);
    return;
  }
  if (!driveCodes.length) {
    ui.alert('⚠️ No drive codes found in the Levels sheet.');
    return;
  }

  // ── 2. Single batch-read of col C ────────────────────────────────────────
  const lastRow       = sheet.getLastRow();
  const existingNames = new Set();          // registered base names (UPPER)
  const sectionEndRow = {};                 // CODE → last row of that section
  driveCodes.forEach(c => { sectionEndRow[c] = 0; });

  if (lastRow >= DATA_START) {
    sheet.getRange(DATA_START, COL.FILENAME, lastRow - DATA_START + 1, 1)
      .getValues()
      .forEach((cell, i) => {
        const name = cell[0].toString().trim();
        if (!name) return;
        existingNames.add(name.toUpperCase());
        const rowIdx = DATA_START + i;
        for (const code of driveCodes) {
          if (name.toUpperCase().startsWith(code + '-')) {
            if (rowIdx > sectionEndRow[code]) sectionEndRow[code] = rowIdx;
            break;
          }
        }
      });
  }

  // ── 3. Search Drive per code ──────────────────────────────────────────────
  const missing = {};
  driveCodes.forEach(c => { missing[c] = []; });
  let searchErrors = 0;

  for (const code of driveCodes) {
    const prefix    = code + '-';
    const seenBases = new Set();
    try {
      const iter = DriveApp.searchFiles("title contains '" + prefix + "'");
      while (iter.hasNext()) {
        const fileName = iter.next().getName();

        // Drive's "contains" is too broad — must actually start with the prefix
        if (!fileName.toUpperCase().startsWith(prefix)) continue;

        // Must follow KAL pattern: CODE-ENTITY_DOCTYPE_Description[…]
        if (!isKALFileName(fileName)) continue;

        const base = extractKALBaseName(fileName);
        if (!base) continue;

        const key = base.toUpperCase();
        if (existingNames.has(key) || seenBases.has(key)) continue;

        seenBases.add(key);
        missing[code].push(base);
      }
    } catch (e) {
      console.error('searchMissingKALFiles [' + code + ']: ' + e.message);
      searchErrors++;
    }
    missing[code].sort(); // alphabetical within each section
  }

  const totalMissing = driveCodes.reduce((s, c) => s + missing[c].length, 0);
  if (totalMissing === 0) {
    let msg = '✅ No missing files found — the registry is up to date!';
    if (searchErrors > 0) msg += '\n\n⚠️ ' + searchErrors + ' code(s) had search errors (View → Logs).';
    ui.alert(msg);
    return;
  }

  // ── 4. Insert missing rows bottom-up (highest section row first) ─────────
  // Sorting descending by section-end row keeps pre-computed indices valid
  // because insertions above a section don't shift rows below it.
  const baseLastRow = sheet.getLastRow();
  const insertOrder = driveCodes
    .filter(c => missing[c].length > 0)
    .sort((a, b) => (sectionEndRow[b] || baseLastRow) - (sectionEndRow[a] || baseLastRow));

  const newRowStart = {}; // CODE → first newly inserted row index (for audit)

  for (const code of insertOrder) {
    const names    = missing[code];
    const afterRow = sectionEndRow[code] || sheet.getLastRow();

    sheet.insertRowsAfter(afterRow, names.length);
    sheet.getRange(afterRow + 1, COL.FILENAME, names.length, 1)
         .setValues(names.map(n => [n]));

    newRowStart[code] = afterRow + 1;
  }

  // ── 5. Audit every newly inserted row ────────────────────────────────────
  let levelsData, templateList;
  try {
    levelsData   = getLevelsData();
    templateList = getTemplateList();
  } catch (e) {
    console.error('searchMissingKALFiles audit setup: ' + e.message);
  }

  if (levelsData) {
    for (const code of insertOrder) {
      const count = missing[code].length;
      for (let i = 0; i < count; i++) {
        const r = newRowStart[code] + i;
        try {
          processAuditForRow(
            sheet, r,
            levelsData.driveUrlLookup, levelsData.validEntities, levelsData.validDocs,
            templateList
          );
          sheet.getRange(r, COL.ROW_NUM).setValue(r - 1);
        } catch (e) {
          console.error('searchMissingKALFiles audit row ' + r + ': ' + e.message);
        }
      }
    }
  }

  // ── 6. Summary alert ─────────────────────────────────────────────────────
  let summary = '✅ Added ' + totalMissing + ' missing file(s):\n\n';
  driveCodes.forEach(c => {
    if (missing[c].length) summary += '  • ' + c + '-  (' + missing[c].length + ' file(s))\n';
  });
  if (searchErrors > 0) summary += '\n⚠️ ' + searchErrors + ' code(s) had search errors (View → Logs).';
  ui.alert(summary);
}

/** Returns true when the filename follows the KAL pattern CODE-ENTITY_DOCTYPE_Desc… */
function isKALFileName(name) {
  return /^[A-Za-z]{2,4}-[A-Za-z]+_[A-Za-z]+_[A-Za-z0-9]/i.test(name);
}

/** Strips _YYYYMMDD_v{n|FINAL} (and trailing whitespace) to return the base name. */
function extractKALBaseName(name) {
  const m = name.match(/^(.+?)(?:_\d{8}_v(?:\d+|FINAL))?\s*$/i);
  return m ? m[1].trim() : name.trim();
}

/**
 * Returns unique drive codes from column A of the Levels sheet (skipping
 * the header row), preserving the order they appear in the sheet.
 */
function getDriveCodesOrdered() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Levels');
  if (!sheet) throw new Error('Sheet "Levels" not found.');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const codes = [];
  const seen  = new Set();
  sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(row => {
    const code = String(row[0]).toUpperCase().trim();
    if (code && !seen.has(code)) { seen.add(code); codes.push(code); }
  });
  return codes;
}

// ── 4. SYNC LOGIC ─────────────────────────────────────────────────────────────

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
