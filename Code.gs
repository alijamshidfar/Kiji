/**
 * KAL ACADEMY - Smart File Registry
 *
 * Column layout (Registry sheet):
 *  A(1)  Row number          B(2)  Human-Readable Description
 *  C(3)  File Name           D(4)  File Type
 *  E(5)  Current Version     F(6)  Current Folder
 *  G(7)  Link                H(8)  For Who
 *  I(9)  KAL Name Check      J(10) Destination Drive
 *  K(11) Preferred Template  L(12) Abstract  [AI-generated]
 *  M(13) Owner               [optional – email for notifications]
 *
 * All global constants (COL, LAST_COL, DATA_START, SHEET, COLOR,
 * KAL_LOGO_BASE64) live in Constants.gs.
 */

// ── Lifecycle ─────────────────────────────────────────────────────────────────

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💠 KAL File System')

    // ── Top-level registry actions ──────────────────────────────────────────
    .addItem('🎯 Audit & Sync Selected File',  'updateSelectedInfo')
    .addItem('🔄 Audit & Sync All Files',       'updateAllInfo')
    .addItem('🔍 Search For Missing KAL Files', 'searchMissingKALFiles')
    .addSeparator()

    // ── File Operations sub-menu ────────────────────────────────────────────
    .addSubMenu(ui.createMenu('📁 File Operations')
      .addItem('🏗️ Create Selected File',    'createSelectedFile')
      .addItem('📂 Open Current Folder',     'openCurrentFolder')
      .addItem('🚚 Move to Destination Drive','moveToDestinationDrive')
      .addItem('🏁 Promote to vFINAL & Move', 'promoteToFinalAndMove')
      .addItem('🗑️ Remove Selected Version',  'removeSelectedFile')
      .addItem('☢️ Remove All Versions',      'removeAllVersions'))
    .addSeparator()

    // ── Maintenance sub-menu ────────────────────────────────────────────────
    .addSubMenu(ui.createMenu('🧹 Maintenance')
      .addItem('🧹 Keep Only Latest Version', 'keepOnlyLatestVersion')
      .addItem('📦 Archive Older Versions',   'archiveOlderVersions')
      .addItem('🛠️ Repair Broken Links',      'repairBrokenLinks')
      .addItem('🧼 Clear All Diagnostics',    'clearAllDiagnostics'))
    .addSeparator()

    // ── Reports & View sub-menu ─────────────────────────────────────────────
    .addSubMenu(ui.createMenu('📊 Reports & View')
      .addItem('📊 Generate Health Report', 'generateHealthReport')
      .addItem('📤 Export Registry to PDF', 'exportRegistryToPDF')
      .addItem('🌓 Toggle Compact View',    'toggleCompactView')
      .addItem('👁️ Toggle Done Rows',        'toggleDoneRows'))
    .addSeparator()

    // ── Codes reference sub-menu ────────────────────────────────────────────
    .addSubMenu(ui.createMenu('📋 Show Codes')
      .addItem('📋 DRIVE Codes',   'showDriveCodes')
      .addItem('📋 ENTITY Codes',  'showEntityCodes')
      .addItem('📋 DOCTYPE Codes', 'showDocTypeCodes'))
    .addSeparator()

    .addItem('🔔 Notify Owner',        'notifyOwner')
    .addItem('📖 Show The User Guide', 'showUserGuide')
    .addToUi();

  try { updateTemplateDropdown(); } catch (_) { /* non-fatal on open */ }
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/** Returns the URL from a cell: rich-text link first, then plain value. */
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

/** Convenience wrapper: toast a message on the active spreadsheet. */
function toast(msg, title, sec) {
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, title || '💠 KAL', sec || 4);
}

// ── 1. DROPDOWN FETCHER ───────────────────────────────────────────────────────

function getTemplateList() {
  const folderId = getIdFromUrl(getUrlFromCell(SHEET.SETTINGS, 'B2'));
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
 * Priority colour model:
 *   🔴 Red    (#f4cccc) – Priority 1: structural / naming error
 *   🟡 Yellow (#fff2cc) – Priority 2: valid name, file not found in Drive
 *   🟢 Green  (#d9ead3) – Priority 3: file found and vFINAL
 *   none                – file found, version < FINAL
 *
 * Error messages use the format  "Invalid Drive Code (XX), Invalid DocType (YY)"
 * (comma-separated, actual code shown in parentheses).
 *
 * @param {boolean} [applyBg=true]   pass false in batch mode; caller writes backgrounds
 * @returns {string|null}            background colour for this row
 */
function processAuditForRow(sheet, r, driveUrlLookup, validEntities, validDocs, templateList, preloadedName, applyBg) {
  if (applyBg === undefined) applyBg = true;

  const baseName     = (preloadedName !== undefined)
    ? preloadedName
    : sheet.getRange(r, COL.FILENAME).getValue().toString().trim();
  const dropdownCell = sheet.getRange(r, COL.TEMPLATE);

  // ── Empty row ─────────────────────────────────────────────────────────────
  if (!baseName) {
    sheet.getRange(r, COL.ROW_NUM, 1, 2).clearContent();
    sheet.getRange(r, COL.FILETYPE, 1, LAST_COL - COL.FILETYPE + 1).clearContent();
    dropdownCell.clearDataValidations();
    if (applyBg) sheet.getRange(r, COL.DESC, 1, LAST_COL - COL.DESC + 1).setBackground(null);
    return null;
  }

  // ── Smart CamelCase description extractor ─────────────────────────────────
  const nameParts = baseName.split(/[-_]/);
  if (nameParts.length >= 4) {
    const clean = nameParts.slice(3).join('')
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

  // ── Validation — Priority 1 error messages with code in parentheses ───────
  const diagnostics = [];

  if (!driveUrlLookup[driveCode]) {
    diagnostics.push('Invalid Drive Code (' + driveRaw + ')');
  } else if (driveRaw !== driveCode) {
    diagnostics.push('Drive Code must be UPPERCASE (' + driveRaw + ')');
  }

  if (!validEntities.has(entityCode)) {
    diagnostics.push('Invalid Entity Code (' + entityRaw + ')');
  } else if (entityRaw !== entityCode) {
    diagnostics.push('Entity Code must be UPPERCASE (' + entityRaw + ')');
  }

  if (!validDocs.has(docTypeCode)) {
    diagnostics.push('Invalid DocType (' + docTypeRaw + ')');
  } else if (docTypeRaw !== docTypeCode) {
    diagnostics.push('DocType must be UPPERCASE (' + docTypeRaw + ')');
  }

  const status = diagnostics.length > 0 ? diagnostics.join(', ') : 'OK';
  sheet.getRange(r, COL.KAL_CHECK).setValue(status);

  // ── Col J: Destination Drive ──────────────────────────────────────────────
  const driveUrl = driveUrlLookup[driveCode];
  if (driveUrl) {
    sheet.getRange(r, COL.DEST_DRIVE)
      .setFormula('=HYPERLINK("' + driveUrl + '", "' + driveCode + ' Drive")');
  } else {
    sheet.getRange(r, COL.DEST_DRIVE).clearContent();
  }

  // ── Drive file details ────────────────────────────────────────────────────
  const info = GET_SMART_DETAILS(baseName);
  sheet.getRange(r, COL.FILETYPE, 1, COL.FOR_WHO - COL.FILETYPE + 1).clearContent();

  if (info.fileLink !== 'Not Found') {
    sheet.getRange(r, COL.FILETYPE, 1, 2).setValues([[info.type, info.version]]);
    sheet.getRange(r, COL.FOLDER).setFormula('=HYPERLINK("' + info.folderLink + '", "' + info.folderName + '")');
    sheet.getRange(r, COL.LINK).setFormula('=HYPERLINK("' + info.fileLink + '", "Link")');
    sheet.getRange(r, COL.FOR_WHO).setValue(entityCode);
    dropdownCell.clearContent().clearDataValidations();
  } else {
    sheet.getRange(r, COL.LINK).setValue('File not found');
    if (templateList && templateList.length > 0) {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(templateList).setAllowInvalid(false).build();
      dropdownCell.setDataValidation(rule);
    }
  }

  // ── Col L: Abstract ───────────────────────────────────────────────────────
  sheet.getRange(r, COL.ABSTRACT).setFormula(
    '=AI("Based ONLY on description \'"&B' + r + '&"\' and filename \'"&C' + r + '&"\', write a two-sentence summary.")'
  );

  // ── Priority colour ───────────────────────────────────────────────────────
  let color;
  if (diagnostics.length > 0)          color = COLOR.RED;     // 🔴 structural error
  else if (info.fileLink === 'Not Found') color = COLOR.YELLOW; // 🟡 file missing
  else if (info.version  === 'FINAL')    color = COLOR.GREEN;   // 🟢 finalized
  else                                   color = COLOR.NONE;

  if (applyBg) {
    sheet.getRange(r, COL.DESC, 1, LAST_COL - COL.DESC + 1).setBackground(color);
    // Sequential number: count non-blank filename rows from DATA_START up to r
    const above = sheet.getRange(DATA_START, COL.FILENAME, r - DATA_START + 1, 1).getValues();
    sheet.getRange(r, COL.ROW_NUM).setValue(above.filter(c => c[0]).length);
  }
  return color;
}

// ── 3. MISSING FILE SEARCH ───────────────────────────────────────────────────

/**
 * Scans Drive for KAL files not yet in the registry and inserts them below
 * their matching drive-code section, then audits each new row immediately.
 */
function searchMissingKALFiles() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  let driveCodes;
  try { driveCodes = getDriveCodesOrdered(); }
  catch (e) { toast(e.message, '⚠️ Error', 6); return; }
  if (!driveCodes.length) { toast('No drive codes found in the Codes sheet.', '⚠️ Warning', 5); return; }

  // Batch-read col C once — builds existing-names Set AND section-end-row map
  const lastRow       = sheet.getLastRow();
  const existingNames = new Set();
  const sectionEndRow = {};
  driveCodes.forEach(c => { sectionEndRow[c] = 0; });

  if (lastRow >= DATA_START) {
    sheet.getRange(DATA_START, COL.FILENAME, lastRow - DATA_START + 1, 1)
      .getValues().forEach((cell, i) => {
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

  // Drive search per code
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
        if (!fileName.toUpperCase().startsWith(prefix)) continue;
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
    missing[code].sort();
  }

  const totalMissing = driveCodes.reduce((s, c) => s + missing[c].length, 0);
  if (totalMissing === 0) {
    toast('Registry is up to date!' + (searchErrors ? ' (' + searchErrors + ' search errors — check Logs)' : ''), '✅ All Good', 5);
    return;
  }

  // Insert bottom-up so pre-computed row indices stay valid
  const baseLastRow = sheet.getLastRow();
  const insertOrder = driveCodes
    .filter(c => missing[c].length > 0)
    .sort((a, b) => (sectionEndRow[b] || baseLastRow) - (sectionEndRow[a] || baseLastRow));

  const newRowStart = {};
  for (const code of insertOrder) {
    const names    = missing[code];
    const afterRow = sectionEndRow[code] || sheet.getLastRow();
    sheet.insertRowsAfter(afterRow, names.length);
    sheet.getRange(afterRow + 1, COL.FILENAME, names.length, 1).setValues(names.map(n => [n]));
    newRowStart[code] = afterRow + 1;
  }

  // Immediately audit new rows
  let levelsData, templateList;
  try { levelsData = getLevelsData(); templateList = getTemplateList(); }
  catch (e) { console.error('searchMissingKALFiles audit setup: ' + e.message); }

  if (levelsData) {
    for (const code of insertOrder) {
      const count = missing[code].length;
      for (let i = 0; i < count; i++) {
        const r = newRowStart[code] + i;
        try {
          processAuditForRow(sheet, r, levelsData.driveUrlLookup, levelsData.validEntities, levelsData.validDocs, templateList);
        } catch (e) { console.error('searchMissingKALFiles audit row ' + r + ': ' + e.message); }
      }
    }
  }

  // Renumber sequentially after insertions
  renumberAllRows_(sheet);

  let summary = 'Added ' + totalMissing + ' file(s): ';
  summary += driveCodes.filter(c => missing[c].length).map(c => c + '- ×' + missing[c].length).join(', ');
  if (searchErrors) summary += ' | ' + searchErrors + ' search error(s)';
  toast(summary, '🔍 Missing Files Found', 6);
}

/** Returns true when the filename follows the KAL pattern CODE-ENTITY_DOCTYPE_Desc… */
function isKALFileName(name) {
  return /^[A-Za-z]{2,4}-[A-Za-z]+_[A-Za-z]+_[A-Za-z0-9]/i.test(name);
}

/** Strips _YYYYMMDD_v{n|FINAL} suffix to return the base name. */
function extractKALBaseName(name) {
  const m = name.match(/^(.+?)(?:_\d{8}_v(?:\d+|FINAL))?\s*$/i);
  return m ? m[1].trim() : name.trim();
}

/**
 * Returns unique drive codes from col A of the Codes sheet,
 * preserving their sheet order.
 */
function getDriveCodesOrdered() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.CODES);
  if (!sheet) throw new Error('Sheet "' + SHEET.CODES + '" not found.');
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
 * Sequentially numbers every row that has a filename, skipping blank rows.
 * Row 1 = first non-blank data row, regardless of sheet row position.
 */
function renumberAllRows_(sheet) {
  const last = sheet.getLastRow();
  if (last < DATA_START) return;
  const n     = last - DATA_START + 1;
  const names = sheet.getRange(DATA_START, COL.FILENAME, n, 1).getValues();
  let seq = 0;
  sheet.getRange(DATA_START, COL.ROW_NUM, n, 1).setValues(
    names.map(row => [row[0] ? ++seq : ''])
  );
}

/**
 * Reads the Codes sheet and returns lookup structures.
 * Uses Sets for O(1) entity/docType validation.
 */
function getLevelsData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.CODES);
  if (!sheet) throw new Error('Sheet "' + SHEET.CODES + '" not found. Check your spreadsheet setup.');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { driveUrlLookup: {}, validEntities: new Set(), validDocs: new Set() };

  const range  = sheet.getRange(1, 1, lastRow, 7);
  const vals   = range.getValues();
  const rich   = range.getRichTextValues();

  const driveUrlLookup = {};
  const validEntities  = new Set();
  const validDocs      = new Set();

  for (let j = 1; j < vals.length; j++) {
    const driveCode = String(vals[j][0]).toUpperCase().trim();
    const hiddenUrl = rich[j][1] ? rich[j][1].getLinkUrl() : null;
    if (driveCode && hiddenUrl) driveUrlLookup[driveCode] = hiddenUrl;
    const entity  = String(vals[j][3]).toUpperCase().trim();
    if (entity)  validEntities.add(entity);
    const docType = String(vals[j][6]).toUpperCase().trim();
    if (docType) validDocs.add(docType);
  }
  return { driveUrlLookup, validEntities, validDocs };
}

/**
 * Audits every data row.
 * Batch-reads col C once, then batch-writes backgrounds and row numbers
 * after the loop (N×11 individual calls → 3 batch calls).
 */
function updateAllInfo() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) return;

  let data, templateList;
  try { data = getLevelsData(); templateList = getTemplateList(); }
  catch (e) { toast('Setup error: ' + e.message, '⚠️ Error', 6); return; }

  const numRows   = lastRow - DATA_START + 1;
  const fileNames = sheet.getRange(DATA_START, COL.FILENAME, numRows, 1).getValues();
  const bgColors  = [];
  let seq = 0, errors = 0;

  for (let i = 0; i < numRows; i++) {
    const r        = DATA_START + i;
    const baseName = String(fileNames[i][0]).trim();
    try {
      const color = processAuditForRow(sheet, r, data.driveUrlLookup, data.validEntities, data.validDocs, templateList, baseName, false);
      bgColors.push(Array(LAST_COL - COL.DESC + 1).fill(color));
    } catch (e) {
      console.error('Row ' + r + ': ' + e.message);
      bgColors.push(Array(LAST_COL - COL.DESC + 1).fill(COLOR.ERROR));
      errors++;
    }
  }

  sheet.getRange(DATA_START, COL.DESC, numRows, LAST_COL - COL.DESC + 1).setBackgrounds(bgColors);
  renumberAllRows_(sheet);

  const msg = errors > 0
    ? 'Done — ' + errors + ' row(s) had errors (View → Logs).'
    : 'All ' + numRows + ' rows audited successfully.';
  toast(msg, '🔄 Audit Complete', 4);
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
    toast('Audit error: ' + e.message, '⚠️ Error', 6);
  }
}

// ── 5. DRIVE SEARCH ENGINE ────────────────────────────────────────────────────

/**
 * Finds the latest versioned Drive file matching baseName.
 * Regex compiled once outside the loop for performance.
 */
function GET_SMART_DETAILS(baseName) {
  const NOT_FOUND = { type: '', version: 'Not Found', folderName: 'Not Found', folderLink: '', fileLink: 'Not Found' };
  if (!baseName) return NOT_FOUND;

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
  } catch (e) { console.error('GET_SMART_DETAILS("' + baseName + '"): ' + e.message); }

  return details;
}

// ── 6. FILE OPERATIONS ────────────────────────────────────────────────────────

function createSelectedFile() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const r     = sheet.getActiveRange().getRow();
  if (r < DATA_START) return;

  const baseName = sheet.getRange(r, COL.FILENAME).getValue().toString().trim();
  if (!baseName) { toast('No filename found in this row.', '🛑 Error', 5); return; }

  let data, templateList;
  try { data = getLevelsData(); templateList = getTemplateList(); }
  catch (e) { toast('Setup error: ' + e.message, '⚠️ Error', 6); return; }

  processAuditForRow(sheet, r, data.driveUrlLookup, data.validEntities, data.validDocs, templateList);
  const nameStatus = sheet.getRange(r, COL.KAL_CHECK).getValue().toString().trim();
  if (nameStatus !== 'OK') { toast('Name issues: ' + nameStatus, '🛑 Cannot Create', 6); return; }

  const dateStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd');
  const info    = GET_SMART_DETAILS(baseName);

  // ── File already exists: copy it as the next version ──────────────────────
  if (info.fileLink !== 'Not Found') {
    if (info.version === 'FINAL') {
      toast('File is already vFINAL — use Promote to vFINAL & Move instead.', 'ℹ️ Info', 5);
      return;
    }
    const nextVer   = (parseInt(info.version, 10) || 1) + 1;
    const newName   = baseName + '_' + dateStr + '_v' + nextVer;
    try {
      const srcFile = DriveApp.getFileById(getIdFromUrl(info.fileLink));
      const parents = srcFile.getParents();
      const folder  = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
      srcFile.makeCopy(newName, folder);
      updateSelectedInfo();
      toast('"' + newName + '" created as new version.', '✅ New Version', 5);
    } catch (e) { toast('Copy failed: ' + e.message, '❌ Error', 6); }
    return;
  }

  // ── File does not exist: create from template ──────────────────────────────
  const templateName = sheet.getRange(r, COL.TEMPLATE).getValue().toString().trim();
  if (!templateName) { toast('Select a template from the dropdown (col K) first.', '🛑 Missing Template', 5); return; }

  try {
    const destId = getIdFromUrl(getUrlFromCell(SHEET.SETTINGS, 'A2'));
    if (!destId) throw new Error('Destination folder URL missing in Settings!A2.');
    const tempFolderId = getIdFromUrl(getUrlFromCell(SHEET.SETTINGS, 'B2'));
    if (!tempFolderId) throw new Error('Template folder URL missing in Settings!B2.');
    const finalFileName = baseName + '_' + dateStr + '_v1';
    const templateFiles = DriveApp.getFolderById(tempFolderId).getFilesByName(templateName);
    if (!templateFiles.hasNext()) { toast('"' + templateName + '" not found in templates folder.', '🛑 Error', 5); return; }
    templateFiles.next().makeCopy(finalFileName, DriveApp.getFolderById(destId));
    updateSelectedInfo();
    toast('"' + finalFileName + '" created successfully!', '✅ Created', 5);
  } catch (e) { toast('Creation failed: ' + e.message, '❌ Error', 6); }
}

/** Opens the Drive folder of the selected row in a new browser tab. */
function openCurrentFolder() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const r     = sheet.getActiveRange().getRow();
  if (r < DATA_START) return;

  const cell = sheet.getRange(r, COL.FOLDER);

  // 1. Try rich-text link (manually set hyperlinks)
  let folderUrl = null;
  try { folderUrl = cell.getRichTextValue().getLinkUrl(); } catch (_) {}

  // 2. Fall back to parsing =HYPERLINK("url","label") formula
  if (!folderUrl) {
    try {
      const formula = cell.getFormula();
      const m = formula.match(/=HYPERLINK\(\s*"([^"]+)"/i);
      if (m) folderUrl = m[1];
    } catch (_) {}
  }

  if (!folderUrl) { toast('No folder link found for this row.', '⚠️ Warning', 5); return; }

  // Show a clickable link — window.open() is blocked by most browsers inside GAS dialogs
  const html = HtmlService.createHtmlOutput(
    '<style>body{font-family:Arial,sans-serif;padding:16px;margin:0}</style>' +
    '<p style="font-size:13px">Click below to open the folder:</p>' +
    '<p><a href="' + folderUrl + '" target="_blank" ' +
    'style="font-size:13px;color:#1155CC;font-weight:bold;">📂 Open Folder</a></p>' +
    '<p style="margin-top:10px"><button onclick="google.script.host.close()" ' +
    'style="font-size:12px;padding:4px 12px;cursor:pointer;">Close</button></p>'
  ).setWidth(280).setHeight(120);
  SpreadsheetApp.getUi().showModalDialog(html, 'Open Folder');
}

/** Moves the latest file version for the selected row to its Destination Drive folder (col J). */
function moveToDestinationDrive() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const r     = sheet.getActiveRange().getRow();
  if (r < DATA_START) return;

  const baseName = sheet.getRange(r, COL.FILENAME).getValue().toString().trim();
  if (!baseName) { toast('No filename in this row.', '⚠️ Warning', 4); return; }

  let destUrl = null;
  try { destUrl = sheet.getRange(r, COL.DEST_DRIVE).getRichTextValue().getLinkUrl(); } catch (_) {}
  const destId = getIdFromUrl(destUrl);
  if (!destId) { toast('No destination folder link in column J.', '🛑 Error', 5); return; }

  const info = GET_SMART_DETAILS(baseName);
  if (info.fileLink === 'Not Found') { toast('File not found in Drive.', '⚠️ Warning', 4); return; }

  try {
    const file = DriveApp.getFileById(getIdFromUrl(info.fileLink));
    file.moveTo(DriveApp.getFolderById(destId));
    updateSelectedInfo();
    toast('"' + file.getName() + '" moved to destination drive.', '🚚 Moved', 5);
  } catch (e) { toast('Move failed: ' + e.message, '❌ Error', 6); console.error('moveToDestinationDrive: ' + e.message); }
}

/** Copies the latest version to vFINAL and archives the previous version. */
function promoteToFinalAndMove() {
  const ui    = SpreadsheetApp.getUi();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const r     = sheet.getActiveRange().getRow();
  if (r < DATA_START) return;

  const baseName = sheet.getRange(r, COL.FILENAME).getValue().toString().trim();
  if (!baseName) { toast('No filename in this row.', '⚠️ Warning', 4); return; }

  const archiveId = getIdFromUrl(getUrlFromCell(SHEET.SETTINGS, 'C2'));
  if (!archiveId) { toast('Archive folder URL missing in Settings!C2.', '🛑 Error', 5); return; }

  const info = GET_SMART_DETAILS(baseName);
  if (info.fileLink === 'Not Found') { toast('No file found in Drive.', '⚠️ Warning', 4); return; }
  if (info.version === 'FINAL')      { toast('Already at vFINAL.', 'ℹ️ Info', 4); return; }

  const response = ui.alert('🏁 Promote to vFINAL',
    'Copy latest version to vFINAL and archive the current version?\n\n' + baseName,
    ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  try {
    const srcFile = DriveApp.getFileById(getIdFromUrl(info.fileLink));
    const dateStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd');
    const finalName = baseName + '_' + dateStr + '_vFINAL';

    // Resolve destination: col J formula → col J rich-text → Settings!A2
    let destUrl = null;
    try {
      const formula = sheet.getRange(r, COL.DEST_DRIVE).getFormula();
      const m = formula.match(/=HYPERLINK\(\s*"([^"]+)"/i);
      if (m) destUrl = m[1];
    } catch (_) {}
    if (!destUrl) {
      try { destUrl = sheet.getRange(r, COL.DEST_DRIVE).getRichTextValue().getLinkUrl(); } catch (_) {}
    }
    if (!destUrl) destUrl = getUrlFromCell(SHEET.SETTINGS, 'A2');
    const destId = getIdFromUrl(destUrl);
    if (!destId) { toast('No destination folder found in col J or Settings!A2.', '🛑 Error', 5); return; }

    srcFile.makeCopy(finalName, DriveApp.getFolderById(destId));
    srcFile.moveTo(DriveApp.getFolderById(archiveId));
    updateSelectedInfo();
    toast('"' + finalName + '" created in destination drive. Original archived.', '🏁 Promoted', 6);
  } catch (e) { toast('Promote failed: ' + e.message, '❌ Error', 6); console.error('promoteToFinalAndMove: ' + e.message); }
}

function removeSelectedFile() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const r     = sheet.getActiveRange().getRow();

  let fileUrl = null;
  try { fileUrl = sheet.getRange(r, COL.LINK).getRichTextValue().getLinkUrl(); } catch (_) {}
  if (!fileUrl) { toast('No file link found in this row.', '⚠️ Warning', 4); return; }

  const fileId = getIdFromUrl(fileUrl);
  if (!fileId) { toast('Could not parse a file ID from the link.', '🛑 Error', 5); return; }

  const response = ui.alert('⚠️ Warning', 'Move THIS version to Trash?', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    updateSelectedInfo();
    toast('Version moved to Trash.', '🗑️ Removed', 4);
  } catch (e) { toast('Could not trash the file: ' + e.message, '❌ Error', 5); }
}

function removeAllVersions() {
  const ui       = SpreadsheetApp.getUi();
  const sheet    = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const r        = sheet.getActiveRange().getRow();
  const baseName = sheet.getRange(r, COL.FILENAME).getValue().toString().trim();
  if (!baseName) return;

  const response = ui.alert('☢️ NUCLEAR WARNING', 'Trash EVERY version of: ' + baseName, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  let trashed = 0, failed = 0;
  try {
    const files = DriveApp.searchFiles("title contains '" + baseName + "'");
    while (files.hasNext()) {
      try { files.next().setTrashed(true); trashed++; }
      catch (e) { console.error('removeAllVersions trash: ' + e.message); failed++; }
    }
  } catch (e) { toast('Drive search failed: ' + e.message, '❌ Error', 5); return; }

  updateSelectedInfo();
  toast('Trashed ' + trashed + ' file(s).' + (failed ? ' (' + failed + ' failed)' : ''), '☢️ Done', 5);
}

// ── 7. MAINTENANCE ────────────────────────────────────────────────────────────

/** Trashes all lower-numbered versions of the selected file, keeping only the highest. */
function keepOnlyLatestVersion() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const r     = sheet.getActiveRange().getRow();
  if (r < DATA_START) return;

  const baseName = sheet.getRange(r, COL.FILENAME).getValue().toString().trim();
  if (!baseName) { toast('No filename in this row.', '⚠️ Warning', 4); return; }

  const response = ui.alert('🧹 Keep Only Latest', 'Trash all older versions of:\n' + baseName, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  const escaped   = baseName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const versionRe = new RegExp(escaped + '.*[vV](\\d+|FINAL)', 'i');
  let latestNum   = -1;
  const allFiles  = [];

  try {
    const iter = DriveApp.searchFiles("title contains '" + baseName + "'");
    while (iter.hasNext()) {
      const file  = iter.next();
      const match = file.getName().match(versionRe);
      if (!match) continue;
      const vNum = match[1].toUpperCase() === 'FINAL' ? 9999 : parseInt(match[1], 10);
      allFiles.push({ file, vNum });
      if (vNum > latestNum) latestNum = vNum;
    }
  } catch (e) { toast('Drive search error: ' + e.message, '❌ Error', 5); return; }

  if (allFiles.length <= 1) {
    toast('Only one version found — nothing to remove.', '🧹 Info', 4);
    return;
  }

  let trashed = 0;
  allFiles.forEach(({ file, vNum }) => {
    if (vNum < latestNum) {
      try { file.setTrashed(true); trashed++; }
      catch (e) { console.error('keepOnlyLatestVersion: ' + e.message); }
    }
  });
  updateSelectedInfo();
  toast('Kept latest. Trashed ' + trashed + ' older version(s).', '🧹 Done', 5);
}

/** Moves all older versions of the selected file to the Archive folder (Settings!C2). */
function archiveOlderVersions() {
  const ui    = SpreadsheetApp.getUi();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const r     = sheet.getActiveRange().getRow();
  if (r < DATA_START) return;

  const baseName = sheet.getRange(r, COL.FILENAME).getValue().toString().trim();
  if (!baseName) { toast('No filename in this row.', '⚠️ Warning', 4); return; }

  const archiveId = getIdFromUrl(getUrlFromCell(SHEET.SETTINGS, 'C2'));
  if (!archiveId) { toast('Archive folder URL missing in Settings!C2.', '🛑 Error', 5); return; }

  const escaped   = baseName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const versionRe = new RegExp(escaped + '.*[vV](\\d+|FINAL)', 'i');
  let latestNum   = -1;
  const allFiles  = [];

  try {
    const iter = DriveApp.searchFiles("title contains '" + baseName + "'");
    while (iter.hasNext()) {
      const file  = iter.next();
      const match = file.getName().match(versionRe);
      if (!match) continue;
      const vNum = match[1].toUpperCase() === 'FINAL' ? 9999 : parseInt(match[1], 10);
      allFiles.push({ file, vNum });
      if (vNum > latestNum) latestNum = vNum;
    }
  } catch (e) { toast('Drive search error: ' + e.message, '❌ Error', 5); return; }

  if (allFiles.length <= 1) {
    toast('Only one version found — nothing older to archive.', '📦 Info', 4);
    return;
  }

  const archiveFolder = DriveApp.getFolderById(archiveId);
  let archived = 0;
  allFiles.forEach(({ file, vNum }) => {
    if (vNum < latestNum) {
      try { file.moveTo(archiveFolder); archived++; }
      catch (e) { console.error('archiveOlderVersions: ' + e.message); }
    }
  });
  updateSelectedInfo();
  toast('Archived ' + archived + ' older version(s).', '📦 Done', 5);
}

/**
 * Re-audits every row whose stored Drive link no longer matches the
 * file's current location (i.e. the file was manually moved in Drive).
 */
function repairBrokenLinks() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) { toast('No data rows to scan.', '🛠️ Repair', 3); return; }

  let data, templateList;
  try { data = getLevelsData(); templateList = getTemplateList(); }
  catch (e) { toast('Setup error: ' + e.message, '⚠️ Error', 6); return; }

  const filenames = sheet.getRange(DATA_START, COL.FILENAME, lastRow - DATA_START + 1, 1).getValues();
  let repaired = 0;

  for (let i = 0; i < filenames.length; i++) {
    const baseName = filenames[i][0].toString().trim();
    if (!baseName) continue;
    const r = DATA_START + i;
    let storedUrl = null;
    try { storedUrl = sheet.getRange(r, COL.LINK).getRichTextValue().getLinkUrl(); } catch (_) {}
    const info = GET_SMART_DETAILS(baseName);
    if (info.fileLink === 'Not Found') continue;
    if (storedUrl !== info.fileLink) {
      processAuditForRow(sheet, r, data.driveUrlLookup, data.validEntities, data.validDocs, templateList, baseName);
      repaired++;
    }
  }
  toast(repaired > 0 ? 'Repaired ' + repaired + ' link(s).' : 'All links are up to date.', '🛠️ Repair Links', 5);
}

/** Clears all row highlight colours on data rows (B–L) and renumbers column A. */
function clearAllDiagnostics() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) return;
  const numRows = lastRow - DATA_START + 1;

  const nullBgs = Array(numRows).fill(null).map(() => Array(LAST_COL - COL.DESC + 1).fill(null));
  sheet.getRange(DATA_START, COL.DESC, numRows, LAST_COL - COL.DESC + 1).setBackgrounds(nullBgs);

  renumberAllRows_(sheet);
  toast('All highlights cleared and rows renumbered.', '🧼 Diagnostics', 3);
}

// ── 8. REPORTS & VIEW ─────────────────────────────────────────────────────────

/**
 * Shows a modal with counts of 🔴 / 🟡 / 🟢 / ⚪ rows.
 */
function generateHealthReport() {
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) { toast('No data rows found.', '📊 Health', 3); return; }

  const numRows   = lastRow - DATA_START + 1;
  const bgs       = sheet.getRange(DATA_START, COL.DESC,     numRows, 1).getBackgrounds();
  const fnames    = sheet.getRange(DATA_START, COL.FILENAME, numRows, 1).getValues();

  let red = 0, yellow = 0, green = 0, ok = 0, empty = 0;
  for (let i = 0; i < numRows; i++) {
    if (!fnames[i][0]) { empty++; continue; }
    const bg = (bgs[i][0] || '').toLowerCase().trim();
    if      (bg === COLOR.RED)    red++;
    else if (bg === COLOR.YELLOW) yellow++;
    else if (bg === COLOR.GREEN)  green++;
    else ok++;
  }
  const total = red + yellow + green + ok;

  const html = HtmlService.createHtmlOutput(`
    <style>
      body{font-family:Arial,sans-serif;padding:14px;margin:0}
      h3{margin:0 0 12px;font-size:15px;color:#1155CC}
      table{width:100%;border-collapse:collapse}
      td,th{padding:7px 10px;text-align:left;border-bottom:1px solid #e0e0e0;font-size:13px}
      th{background:#1155CC;color:#fff}
      .red{background:${COLOR.RED}}
      .yel{background:${COLOR.YELLOW}}
      .grn{background:${COLOR.GREEN}}
      .tot{font-weight:700;border-top:2px solid #aaa}
      .note{font-size:10px;color:#999;margin-top:10px}
    </style>
    <h3>📊 Registry Health Report</h3>
    <table>
      <tr><th>State</th><th>Priority</th><th>#</th></tr>
      <tr class="red"><td>🔴 Structural Errors</td><td>1 — Fix naming</td><td>${red}</td></tr>
      <tr class="yel"><td>🟡 Missing in Drive</td><td>2 — Create file</td><td>${yellow}</td></tr>
      <tr class="grn"><td>🟢 Finalized (vFINAL)</td><td>3 — Complete</td><td>${green}</td></tr>
      <tr><td>⚪ Active / In Progress</td><td>—</td><td>${ok}</td></tr>
      <tr class="tot"><td colspan="2">Total registered files</td><td>${total}</td></tr>
    </table>
    <p class="note">Generated ${new Date().toLocaleString()}</p>
  `).setWidth(390).setHeight(270);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 Health Report');
}

/**
 * Exports the active sheet as a PDF to the destination folder (Settings!A2).
 * Shows a clickable link in a small modal on completion.
 */
function exportRegistryToPDF() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getActiveSheet();
  const dateStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd_HHmm');
  const pdfName = 'KAL_Registry_Export_' + dateStr + '.pdf';

  const exportUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?' + [
    'format=pdf', 'gid=' + sheet.getSheetId(),
    'portrait=false', 'fitw=true', 'size=A3',
    'top_margin=0.25', 'bottom_margin=0.25', 'left_margin=0.25', 'right_margin=0.25',
    'sheetnames=false', 'printtitle=false', 'pagenumbers=false', 'gridlines=false', 'fzr=false'
  ].join('&');

  try {
    const blob = UrlFetchApp.fetch(exportUrl, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } })
      .getBlob().setName(pdfName);
    const destId  = getIdFromUrl(getUrlFromCell(SHEET.SETTINGS, 'A2'));
    const pdfFile = destId ? DriveApp.getFolderById(destId).createFile(blob) : DriveApp.createFile(blob);
    const html = HtmlService.createHtmlOutput(
      '<p style="font-family:Arial;padding:12px;font-size:13px">PDF ready: ' +
      '<a href="' + pdfFile.getUrl() + '" target="_blank"><b>' + pdfName + '</b></a></p>' +
      '<script>setTimeout(()=>google.script.host.close(),9000)</script>'
    ).setWidth(420).setHeight(70);
    SpreadsheetApp.getUi().showModalDialog(html, '📤 Export Complete');
  } catch (e) { toast('Export failed: ' + e.message, '❌ Error', 6); console.error('exportRegistryToPDF: ' + e.message); }
}

/** Toggles visibility of helper columns D, H, J, L. */
function toggleCompactView() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = ss.getActiveSheet();
  const COLS   = [COL.FILETYPE, COL.FOR_WHO, COL.DEST_DRIVE, COL.ABSTRACT]; // D H J L
  const hidden = sheet.isColumnHiddenByUser(COLS[0]);
  COLS.forEach(c => hidden ? sheet.showColumns(c) : sheet.hideColumns(c));
  toast(hidden ? 'Helper columns shown.' : 'Helper columns hidden (D, H, J, L).', '🌓 Compact View', 4);
}

/**
 * Hides rows where the background is 🟢 green or version = FINAL.
 * Running again while rows are hidden restores all rows.
 */
function toggleDoneRows() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) return;

  const numRows = lastRow - DATA_START + 1;
  const bgs     = sheet.getRange(DATA_START, COL.DESC,    numRows, 1).getBackgrounds();
  const vers    = sheet.getRange(DATA_START, COL.VERSION, numRows, 1).getValues();

  const doneRows = [];
  for (let i = 0; i < numRows; i++) {
    const bg  = (bgs[i][0]  || '').toLowerCase().trim();
    const ver = (vers[i][0] || '').toString().toUpperCase().trim();
    if (bg === COLOR.GREEN || ver === 'FINAL') doneRows.push(DATA_START + i);
  }
  if (!doneRows.length) { toast('No finalized rows found.', '👁️ Toggle', 3); return; }

  const shouldHide = !sheet.isRowHiddenByUser(doneRows[0]);
  doneRows.forEach(r => shouldHide ? sheet.hideRows(r) : sheet.showRows(r));
  toast(shouldHide ? doneRows.length + ' finalized row(s) hidden.' : 'All rows visible.', '👁️ Toggle Done', 4);
}

// ── 9. CODES DISPLAY ─────────────────────────────────────────────────────────

function showDriveCodes()   { _showCodesDialog('drives',   '📋 DRIVE Codes');   }
function showEntityCodes()  { _showCodesDialog('entities', '📋 ENTITY Codes');  }
function showDocTypeCodes() { _showCodesDialog('docTypes', '📋 DOCTYPE Codes'); }

/**
 * Generic codes dialog — reads from getAcademyCodes() and renders an HTML table.
 * @param {'drives'|'entities'|'docTypes'} key
 */
function _showCodesDialog(key, title) {
  const data  = getAcademyCodes();
  const items = data[key] || [];

  const rows = items.map(item =>
    '<tr><td><b>' + item.code + '</b></td><td>' + (item.name || '—') + '</td></tr>'
  ).join('');

  const html = HtmlService.createHtmlOutput(`
    <style>
      body{font-family:Arial,sans-serif;padding:12px;margin:0}
      h3{margin:0 0 10px;font-size:14px;color:#1155CC}
      table{width:100%;border-collapse:collapse}
      td,th{padding:6px 10px;border-bottom:1px solid #eee;font-size:12px}
      th{background:#1155CC;color:#fff}
      b{font-size:13px;color:#333}
    </style>
    <h3>${title}</h3>
    <table>
      <tr><th>Code</th><th>Name / Description</th></tr>
      ${rows || '<tr><td colspan="2" style="color:#999">No entries found.</td></tr>'}
    </table>
  `).setWidth(440).setHeight(Math.min(420, 80 + items.length * 30));
  SpreadsheetApp.getUi().showModalDialog(html, title);
}

// ── 10. NOTIFICATIONS ─────────────────────────────────────────────────────────

/**
 * Sends an email to the owner defined in column M of the selected row.
 */
function notifyOwner() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const r     = sheet.getActiveRange().getRow();
  if (r < DATA_START) return;

  const ownerEmail = sheet.getRange(r, COL.OWNER).getValue().toString().trim();
  const baseName   = sheet.getRange(r, COL.FILENAME).getValue().toString().trim();

  if (!ownerEmail)          { toast('No owner email in column M for this row.', '🔔 Notify', 5); return; }
  if (!ownerEmail.includes('@')) { toast('Invalid email in column M: ' + ownerEmail, '🔔 Notify', 5); return; }

  try {
    MailApp.sendEmail({
      to:      ownerEmail,
      subject: '[KAL Registry] Action required: ' + baseName,
      body:    'Dear Owner,\n\nYou are requested to review the following document in the KAL File Registry:\n\n' +
               'File: ' + baseName + '\n\n' +
               'Please log in to the KAL File Registry to take the necessary action.\n\nKAL File System'
    });
    toast('Notification sent to ' + ownerEmail + '.', '🔔 Sent', 5);
  } catch (e) {
    toast('Email failed: ' + e.message, '❌ Error', 6);
    console.error('notifyOwner: ' + e.message);
  }
}

// ── 11. UI ────────────────────────────────────────────────────────────────────

function showUserGuide() {
  try {
    const tmpl = HtmlService.createTemplateFromFile('Sidebar');
    tmpl.logoBase64 = KAL_LOGO_BASE64;
    const html = tmpl.evaluate().setTitle('​').setWidth(350);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    toast('Could not load the User Guide: ' + e.message, '❌ Error', 5);
  }
}

// ── 12. ACADEMY SIDEBAR DATA ──────────────────────────────────────────────────

function getAcademyCodes() {
  const EMPTY = { drives: [], entities: [], docTypes: [], examples: [] };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.CODES);
    if (!sheet) return EMPTY;
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return EMPTY;

    const values   = sheet.getRange(3, 1, lastRow - 2, 10).getValues();
    const drives   = [], entities = [], docTypes = [], examples = [];

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
