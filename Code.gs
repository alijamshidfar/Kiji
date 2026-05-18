/**
 * KAL ACADEMY - Smart File Registry
 *
 * Column layout (Registry sheet):
 *  A(1)  Row number          B(2)  Human-Readable Name
 *  C(3)  Base File Name      D(4)  File Type
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
  ui.createMenu('💠 KAL File Registry')

    // ── Top-level registry actions ──────────────────────────────────────────
    .addItem('🎯 Audit & Sync Selected File',  'updateSelectedInfo')
    .addItem('🔄 Audit & Sync All Files',       'updateAllInfo')
    .addItem('🔍 Search For Missing KAL Files', 'searchMissingKALFiles')
    .addSeparator()

    // ── File Operations sub-menu ────────────────────────────────────────────
    .addSubMenu(ui.createMenu('📁 File Operations')
      .addItem('📋 Show All Versions',      'showFileVersions')
      .addSeparator()
      .addItem('🏗️ Create Selected File',    'createSelectedFile')
      .addItem('📂 Open Current Folder',     'openCurrentFolder')
      .addItem('🚚 Move to Destination Drive','moveToDestinationDrive')
      .addItem('🏁 Promote to vFINAL & Move', 'promoteToFinalAndMove')
      .addItem('🗑️ Remove Selected Version',  'removeSelectedFile')
      .addItem('☢️ Remove All Versions',      'removeAllVersions'))
    .addSeparator()

    // ── Maintenance sub-menu ────────────────────────────────────────────────
    .addSubMenu(ui.createMenu('🧹 Maintenance')
      .addItem('🔁 Rebuild Registry from Drive', 'rebuildRegistryFromDrive')
      .addItem('🗂️ Refresh Non-KAL Section',     'refreshNonKalSection')
      .addSeparator()
      .addItem('🧹 Keep Only Latest Version', 'keepOnlyLatestVersion')
      .addItem('📦 Archive Older Versions',   'archiveOlderVersions')
      .addItem('🛠️ Repair Broken Links',      'repairBrokenLinks')
      .addItem('🧼 Clear All Diagnostics',    'clearAllDiagnostics')
      .addSeparator()
      .addItem('📦 Bulk Archive Selected Rows', 'bulkArchiveSelected'))
    .addSeparator()

    // ── Reports & View sub-menu ─────────────────────────────────────────────
    .addSubMenu(ui.createMenu('📊 Reports & View')
      .addItem('📊 Generate Health Report', 'generateHealthReport')
      .addItem('📊 Summary Dashboard',       'generateSummaryDashboard')
      .addItem('📭 Missing Files Report',    'showMissingFilesReport')
      .addItem('🔁 Detect Duplicates',       'detectDuplicates')
      .addItem('⏰ Flag Stale Files',         'flagStaleFiles')
      .addItem('🔍 Filter Registry',         'showFilterSidebar')
      .addSeparator()
      .addItem('📤 Export Registry to PDF', 'exportRegistryToPDF')
      .addItem('🌓 Toggle Compact View',    'toggleCompactView')
      .addItem('👁️ Toggle Done Rows',        'toggleDoneRows'))
    .addSeparator()

    .addItem('🔔 Notify Owner',        'notifyOwner')
    .addItem('📖 User Guide', 'showUserGuide')
    .addToUi();

  try { updateTemplateDropdown(); } catch (_) { /* non-fatal on open */ }

  // Auto-scan for newly created Drive files and add them to the registry.
  try { searchMissingKALFiles(); } catch (_) { /* non-fatal on open */ }
}

// ── Sheet font default ────────────────────────────────────────────────────────

/** Applies Georgia/10 to every non-header row so newly added rows inherit it. */
function applySheetFont_(sheet) {
  const maxRows = sheet.getMaxRows();
  if (maxRows < DATA_START) return;
  // Start from COL.DESC (col B) — col A formatting is managed by renumberAllRows_
  sheet.getRange(DATA_START, COL.DESC, maxRows - DATA_START + 1, COL.OWNER - COL.DESC + 1)
       .setFontFamily('Georgia')
       .setFontSize(10);
}

/**
 * Simple onEdit trigger — re-applies Georgia/10 to any row the user edits in
 * the Registry sheet, so manually added rows get the correct default instantly.
 */
function onEdit(e) {
  if (!e) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET.REGISTRY) return;
  const r = e.range.getRow();
  if (r < DATA_START) return;

  // Protect the pinned master document: revert any rename attempt on col C.
  if (e.range.getColumn() === COL.FILENAME) {
    const current = e.range.getValue().toString().trim();
    const old     = (e.oldValue || '').toString().trim();
    if (old.startsWith(PINNED_FILE_BASE) && !current.startsWith(PINNED_FILE_BASE)) {
      e.range.setValue(old);
      toast('The master registry document name is protected and cannot be changed.', '🔒 Protected', 5);
      return;
    }
  }

  // Start from COL.DESC (col B) — col A is managed by renumberAllRows_
  sheet.getRange(r, COL.DESC, 1, COL.OWNER - COL.DESC + 1)
       .setFontFamily('Georgia')
       .setFontSize(10);
}

// ── Abstract generation (rule-based, no API key required) ────────────────────

/**
 * Builds a two-sentence abstract from row metadata already available in the
 * registry. No external API, no cost, no setup.
 */
function generateAbstract_(description, docType, entity, drive, version, folder) {
  const typeVerb = {
    POLICY: 'establishes the policy for', PROCEDURE: 'defines the procedure for',
    TEMPLATE: 'provides a standard template for', REPORT: 'presents a report on',
    GUIDE: 'provides guidance on', FORM: 'is the official form for',
    CONTRACT: 'contains the contract terms for', PLAN: 'outlines the plan for',
    TRAINING: 'covers training material on', MANUAL: 'is the manual for',
    CHECKLIST: 'provides a checklist for', AGREEMENT: 'documents the agreement for',
    SOP: 'describes the standard operating procedure for'
  };

  const verb   = typeVerb[(docType || '').toUpperCase()] || 'documents';
  const desc   = (description || '').trim().replace(/\.+$/, '');
  const line1  = desc
    ? 'This document ' + verb + ': ' + desc + '.'
    : 'This KAL ' + (docType || 'document').toLowerCase() + ' ' + verb + ' the topic.';

  const parts = [];
  if (entity && entity !== 'ALL') parts.push('Owned by ' + entity);
  if (drive)                       parts.push(drive + ' drive');
  if (version)                     parts.push('v' + version);
  if (folder)                      parts.push('stored in ' + folder);
  const line2 = parts.length ? parts.join(' · ') + '.' : '';

  return (line1 + (line2 ? '  ' + line2 : '')).trim() || null;
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
    // Preserve separator / header rows (navy or red background) — touch nothing
    const rowBg = sheet.getRange(r, COL.DESC).getBackground().toLowerCase();
    if (rowBg === HEADER_BLUE.toLowerCase() || rowBg === SEPARATOR_RED.toLowerCase()) return rowBg;

    sheet.getRange(r, COL.ROW_NUM, 1, 2).clearContent();
    sheet.getRange(r, COL.FILETYPE, 1, LAST_COL - COL.FILETYPE + 1).clearContent();
    dropdownCell.clearDataValidations();
    if (applyBg) sheet.getRange(r, COL.DESC, 1, LAST_COL - COL.DESC + 1).setBackground(null);
    return null;
  }

  // ── Non-KAL file: does not follow DRIVEPREFIX-ENTITY_DOCTYPE convention ─────
  if (!/^[A-Z]{2,4}-[A-Z]/.test(baseName) || !baseName.includes('_')) {
    sheet.getRange(r, COL.KAL_CHECK).setValue('Non-KAL');
    const nonKalColor = COLOR.NON_KAL;
    if (applyBg) sheet.getRange(r, COL.DESC, 1, LAST_COL - COL.DESC + 1).setBackground(nonKalColor);
    return nonKalColor;
  }

  // ── Smart CamelCase description extractor ─────────────────────────────────
  const nameParts = baseName.split(/[-_]/);
  if (nameParts.length >= 4) {
    const clean = nameParts.slice(3)
      .filter(p => !/^\d{8}$/.test(p) && !/^v\d+$/i.test(p) && !/^vFINAL$/i.test(p))
      .join('')
      .replace(/([a-z\d])([A-Z])/g,    '$1 $2')
      .replace(/([A-Z]+)([A-Z][a-z])/g,'$1 $2')
      .replace(/([a-zA-Z])(\d)/g,      '$1 $2')
      .replace(/(\d)([a-zA-Z])/g,      '$1 $2')
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
  // Rule-based: instant, no API. Only writes when the cell is empty so
  // existing user-written abstracts are never overwritten.
  const abstractCell = sheet.getRange(r, COL.ABSTRACT);
  if (!abstractCell.getValue() && !abstractCell.getFormula()) {
    const desc     = (sheet.getRange(r, COL.DESC).getValue() || '').toString().trim();
    const abstract = generateAbstract_(desc, docTypeCode, entityCode, driveCode, info.version, info.folderName);
    if (abstract) abstractCell.setValue(abstract);
  }

  // ── Priority colour ───────────────────────────────────────────────────────
  let color;
  if (diagnostics.length > 0)          color = COLOR.RED;     // 🔴 structural error
  else if (info.fileLink === 'Not Found') color = COLOR.YELLOW; // 🟡 file missing
  else if (info.version  === 'FINAL')    color = COLOR.GREEN;   // 🟢 finalized
  else                                   color = COLOR.NONE;

  if (applyBg) {
    sheet.getRange(r, COL.DESC, 1, LAST_COL - COL.DESC + 1).setBackground(color);
    renumberAllRows_(sheet);
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
  // Always operate on the Registry sheet — the active sheet may be different
  // when this is called automatically from onOpen().
  const sheet = ss.getSheetByName(SHEET.REGISTRY) || ss.getActiveSheet();

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

  // BFS from Root Folder (Settings!E2) — same proven approach as rebuildCollectNonKalFiles_.
  // driveUrlLookup entries point to Shared Drive roots which getFolderById cannot traverse;
  // the Root Folder is a regular folder that works reliably.
  const missing = {};
  driveCodes.forEach(c => { missing[c] = []; });

  let levelsData;
  try { levelsData = getLevelsData(); }
  catch (e) { toast('Could not load Codes sheet: ' + e.message, '⚠️ Error', 6); return; }

  const rootUrl = getUrlFromCell(SHEET.SETTINGS, 'E2');
  const rootId  = rootUrl ? getIdFromUrl(rootUrl) : null;
  if (!rootId) {
    toast('Settings!E2 (Root Folder) is empty — cannot scan for new files.', '⚠️ Warning', 5);
    return;
  }

  const MAX_DEPTH  = 5;
  const seenBases  = new Set();
  const KAL_RE     = /^[A-Z]{2,4}-[A-Z]/;

  const queue = [{ id: rootId, depth: 0 }];
  while (queue.length > 0) {
    const { id, depth } = queue.shift();
    let folder;
    try { folder = DriveApp.getFolderById(id); } catch (_) { continue; }

    try {
      const files = folder.getFiles();
      while (files.hasNext()) {
        const fileName = files.next().getName();
        if (!KAL_RE.test(fileName) || !fileName.includes('_')) continue;
        if (!isKALFileName(fileName)) continue;
        const base = extractKALBaseName(fileName);
        if (!base) continue;
        const key = base.toUpperCase();
        if (existingNames.has(key) || seenBases.has(key)) continue;
        seenBases.add(key);
        // Assign to the correct drive code bucket
        const codeMatch = base.match(/^([A-Za-z]{2,4})-/);
        const code = codeMatch ? codeMatch[1].toUpperCase() : null;
        if (code && missing[code]) missing[code].push(base);
      }
    } catch (_) {}

    if (depth < MAX_DEPTH) {
      try {
        const subs = folder.getFolders();
        while (subs.hasNext()) queue.push({ id: subs.next().getId(), depth: depth + 1 });
      } catch (_) {}
    }
  }
  driveCodes.forEach(c => missing[c].sort());

  const totalMissing = driveCodes.reduce((s, c) => s + missing[c].length, 0);
  if (totalMissing === 0) {
    toast('Registry is up to date!', '✅ All Good', 5);
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
  let templateList;
  try { templateList = getTemplateList(); }
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

  const summary = 'Added ' + totalMissing + ' file(s): ' +
    driveCodes.filter(c => missing[c].length).map(c => c + '- ×' + missing[c].length).join(', ');
  toast(summary, '🔍 Missing Files Found', 6);
}

/** Returns true when the filename follows the KAL pattern CODE-ENTITY_DOCTYPE_Desc… */
function isKALFileName(name) {
  return /^[A-Za-z]{2,4}-[A-Za-z]+_[A-Za-z]+_[A-Za-z0-9]/i.test(name);
}

/** Strips _YYYYMMDD_v{n|FINAL} suffix to return the base name. */
function extractKALBaseName(name) {
  // Strip optional date (_YYYYMMDD) AND optional version (_vN or _vFINAL) from the end.
  // Both parts are optional and independent so files like _v2 (no date) are handled too.
  const m = name.match(/^(.+?)(?:_\d{8})?(?:_v(?:\d+|FINAL))?\s*$/i);
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
 * Sequentially numbers rows that have a filename AND an existing Drive file.
 * Rows with no filename or "File not found" get a blank in column A.
 */
function renumberAllRows_(sheet) {
  const last = sheet.getLastRow();
  if (last < DATA_START) return;
  const n     = last - DATA_START + 1;
  const names = sheet.getRange(DATA_START, COL.FILENAME, n, 1).getValues();
  // Use col D (file type) rather than col G (HYPERLINK formula) to determine whether a file
  // exists. Col D is written as a literal value via setValues(), so it is always immediately
  // readable. Col G contains a formula whose cached display value can be stale immediately
  // after setFormula(), causing renumberAllRows_ to miss file rows and leave col A blank.
  const types = sheet.getRange(DATA_START, COL.FILETYPE, n, 1).getValues();

  let seq = 0;
  const values  = [];
  const bgs     = [];
  const fColors = [];
  const weights = [];
  const aligns  = [];

  names.forEach((row, i) => {
    const hasFile    = !!row[0];
    const fileExists = hasFile && !!types[i][0].toString().trim();
    const isPinned   = hasFile && row[0].toString().startsWith(PINNED_FILE_BASE);

    if (isPinned) {
      values.push(['★']);
      bgs.push([PINNED_GOLD]);
      fColors.push(['#ffffff']);
      weights.push(['bold']);
      aligns.push(['center']);
    } else {
      // All col A cells → navy; file rows get white bold number, others stay blank
      values.push([fileExists ? ++seq : '']);
      bgs.push([HEADER_BLUE]);
      fColors.push(['#ffffff']);
      weights.push([fileExists ? 'bold' : 'normal']);
      aligns.push(['center']);
    }
  });

  const colA = sheet.getRange(DATA_START, COL.ROW_NUM, n, 1);
  colA.setValues(values)
      .setBackgrounds(bgs)
      .setFontColors(fColors)
      .setFontWeights(weights)
      .setHorizontalAlignments(aligns);
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
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.REGISTRY);
  if (!sheet) { toast('Registry sheet not found.', '❌ Error', 5); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) return;

  let data, templateList;
  try { data = getLevelsData(); templateList = getTemplateList(); }
  catch (e) { toast('Setup error: ' + e.message, '⚠️ Error', 6); return; }

  const numRows   = lastRow - DATA_START + 1;
  const fileNames = sheet.getRange(DATA_START, COL.FILENAME, numRows, 1).getValues();

  // Read existing backgrounds so separator rows (HEADER_BLUE, no filename) are preserved
  const existingBgs = sheet.getRange(DATA_START, COL.DESC, numRows, 1).getBackgrounds();

  const bgColors = [];
  let errors = 0;

  for (let i = 0; i < numRows; i++) {
    const r        = DATA_START + i;
    const baseName = String(fileNames[i][0]).trim();

    // Separator / header row: no filename and background is navy or red — preserve, skip audit
    if (!baseName) {
      const bg = existingBgs[i][0].toLowerCase();
      if (bg === HEADER_BLUE.toLowerCase() || bg === SEPARATOR_RED.toLowerCase()) {
        bgColors.push(Array(LAST_COL - COL.DESC + 1).fill(existingBgs[i][0]));
        continue;
      }
    }

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

  // Ensure 3 blank rows between groups and after the last group FIRST,
  // so that row-height pass below also covers any newly inserted blank rows.
  const rowsInserted = maintainGroupSpacing_(sheet);

  // Set row heights: 21 px for file rows, 20 px for blank separator rows
  const lastDataRow = sheet.getLastRow();
  if (lastDataRow >= DATA_START) {
    const N = lastDataRow - DATA_START + 1;
    sheet.setRowHeights(DATA_START, N, 21); // first pass: everything to 21 px
    const colC = sheet.getRange(DATA_START, COL.FILENAME, N, 1).getValues();
    colC.forEach((r, i) => { if (!r[0]) sheet.setRowHeight(DATA_START + i, 20); });
  }

  // Renumber AFTER all row insertions/deletions so col A always reflects the final layout.
  renumberAllRows_(sheet);
  SpreadsheetApp.flush();

  let msg = errors > 0
    ? 'Done — ' + errors + ' row(s) had errors (View → Logs).'
    : 'All ' + numRows + ' rows audited successfully.';
  if (rowsInserted > 0) msg += ' (' + rowsInserted + ' spacing row(s) added)';
  toast(msg, '🔄 Audit Complete', 4);
}

/**
 * Scans col C for duplicate filenames and deletes later occurrences (bottom-to-top
 * so row-index shifts don't corrupt the loop).  The FIRST occurrence is kept —
 * after consolidation that is the original, fully-audited row.
 *
 * @returns {boolean} true if any rows were deleted
 */
function removeDuplicateFilenames_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) return false;

  const n     = lastRow - DATA_START + 1;
  const names = sheet.getRange(DATA_START, COL.FILENAME, n, 1).getValues();

  const seen     = new Set();
  const toDelete = [];

  for (let i = 0; i < n; i++) {
    const fn = (names[i][0] || '').toString().trim();
    if (!fn) continue;
    // Non-KAL files can legitimately share the same name across different subfolders
    // (e.g. two "Budget.xlsx" files in separate project folders).  Only deduplicate
    // rows that follow the KAL naming convention.
    if (!/^[A-Z]{2,4}-[A-Z]/.test(fn) || !fn.includes('_')) continue;
    const key = fn.toLowerCase();
    if (seen.has(key)) {
      toDelete.push(DATA_START + i); // mark for deletion (later occurrence)
    } else {
      seen.add(key);
    }
  }

  if (!toDelete.length) return false;

  // Delete bottom-to-top so earlier row indices stay valid
  for (let i = toDelete.length - 1; i >= 0; i--) {
    console.log('removeDuplicateFilenames_: deleting duplicate at row ' + toDelete[i]);
    sheet.deleteRows(toDelete[i]);
  }
  return true;
}

/**
 * Reads filenames from col C and returns [{prefix, firstRow, lastRow}] for each
 * consecutive run of files sharing the same drive-code prefix.  Blank rows skipped.
 */
function buildGroups_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) return [];
  const n     = lastRow - DATA_START + 1;
  const names = sheet.getRange(DATA_START, COL.FILENAME, n, 1).getValues();
  const groups = [];
  for (let i = 0; i < n; i++) {
    const fn = (names[i][0] || '').toString().trim();
    if (!fn) continue;
    // Non-KAL files must not be touched by consolidateMisplacedRows_ or the spacing
    // logic — skip them here so they are invisible to all group-management operations.
    if (!/^[A-Z]{2,4}-[A-Z]/.test(fn) || !fn.includes('_')) continue;
    const m      = fn.match(/^([A-Za-z]{2,4})-/);
    const prefix = m ? m[1].toUpperCase() : '??';
    const r      = DATA_START + i;
    if (!groups.length || groups[groups.length - 1].prefix !== prefix) {
      groups.push({ prefix, firstRow: r, lastRow: r });
    } else {
      groups[groups.length - 1].lastRow = r;
    }
  }
  return groups;
}

/**
 * Finds file rows placed in the wrong group section (e.g. an OP row sitting
 * between LP and PC rows) and moves them up to immediately after the first
 * occurrence of that prefix.  Uses sheet.copyTo so values, formulas and
 * formatting all transfer correctly.  Iterates until the sheet is fully
 * consolidated (handles cascading moves).
 *
 * @returns {boolean} true if any rows were moved
 */
function consolidateMisplacedRows_(sheet) {
  let anyMoved = false;

  for (let pass = 0; pass < 10; pass++) {         // safety cap
    const groups = buildGroups_(sheet);

    // ── Check 1: same prefix appears in multiple separate blocks → merge them ──
    const firstIdx  = {};
    let lastMisplaced = null;
    for (let g = 0; g < groups.length; g++) {
      const p = groups[g].prefix;
      if (firstIdx[p] === undefined) {
        firstIdx[p] = g;
      } else {
        lastMisplaced = { g, firstGroupIdx: firstIdx[p] };
      }
    }
    if (lastMisplaced) {
      const src         = groups[lastMisplaced.g];
      const target      = groups[lastMisplaced.firstGroupIdx];
      const insertAfter = target.lastRow;
      const srcFirst    = src.firstRow;
      const numRows     = src.lastRow - srcFirst + 1;

      console.log('consolidateMisplacedRows_: merging duplicate ' + src.prefix +
        ' rows [' + srcFirst + '-' + src.lastRow + '] to after row ' + insertAfter);

      sheet.insertRowsAfter(insertAfter, numRows);
      const shiftedSrc = srcFirst + numRows; // target is above src, so src shifts down
      sheet.getRange(shiftedSrc, 1, numRows, COL.OWNER)
           .copyTo(sheet.getRange(insertAfter + 1, 1, numRows, COL.OWNER));
      sheet.deleteRows(shiftedSrc, numRows);
      anyMoved = true;
      continue;
    }

    // ── Check 2: unknown prefix appears before a known-prefix group → move it down ──
    // Find the last row that belongs to any known (REBUILD_PREFIX_ORDER) group.
    let lastKnownRow = -1;
    for (let g = groups.length - 1; g >= 0; g--) {
      if (REBUILD_PREFIX_ORDER.includes(groups[g].prefix)) {
        lastKnownRow = groups[g].lastRow;
        break;
      }
    }
    if (lastKnownRow > 0) {
      let unknownEarlyIdx = -1;
      for (let g = 0; g < groups.length; g++) {
        if (!REBUILD_PREFIX_ORDER.includes(groups[g].prefix) &&
            groups[g].lastRow < lastKnownRow) {
          unknownEarlyIdx = g;
          break;
        }
      }
      if (unknownEarlyIdx >= 0) {
        const src     = groups[unknownEarlyIdx];
        const numRows = src.lastRow - src.firstRow + 1;

        console.log('consolidateMisplacedRows_: moving unknown prefix ' + src.prefix +
          ' rows [' + src.firstRow + '-' + src.lastRow + '] to after row ' + lastKnownRow);

        // Target (lastKnownRow) is BELOW src, so insertRowsAfter does not shift src.
        sheet.insertRowsAfter(lastKnownRow, numRows);
        sheet.getRange(src.firstRow, 1, numRows, COL.OWNER)
             .copyTo(sheet.getRange(lastKnownRow + 1, 1, numRows, COL.OWNER));
        sheet.deleteRows(src.firstRow, numRows);
        anyMoved = true;
        continue;
      }
    }

    break; // nothing left to fix
  }

  return anyMoved;
}

/**
 * After each audit, guarantees exactly 3 blank rows between drive-code groups
 * and after the last group.  When a user fills a blank separator row with a
 * new filename, the next Sync call inserts the missing blank rows and
 * re-applies the navy col-A styling and red separator border.
 *
 * Also detects rows placed in the wrong section (e.g. an OP row inserted
 * between LP and PC rows) and moves them to the correct group automatically.
 *
 * Processes boundaries bottom-to-top so insertions at one boundary never
 * invalidate the row numbers captured for boundaries higher up the sheet.
 *
 * @returns {number} total rows inserted (0 if already correct)
 */
function maintainGroupSpacing_(sheet) {
  try {
    if (sheet.getLastRow() < DATA_START) return 0;

    // 1. Move any rows that are in the wrong section to their correct group.
    consolidateMisplacedRows_(sheet);

    // 2. Remove duplicate filenames created by the move (e.g. user typed a
    //    filename that already exists in the target group).
    removeDuplicateFilenames_(sheet);

    // 3. Build groups fresh (rows may have moved or been deleted above).
    const groups = buildGroups_(sheet);

    console.log('maintainGroupSpacing_: detected ' + groups.length + ' group(s): ' +
      groups.map(g => g.prefix + '[' + g.firstRow + '-' + g.lastRow + ']').join(', '));

    if (!groups.length) return 0;

    let totalInserted = 0;

    // ── Inter-group spacing (bottom-to-top) ─────────────────────────────────
    for (let g = groups.length - 2; g >= 0; g--) {
      const endRow = groups[g].lastRow;
      const nxtRow = groups[g + 1].firstRow;
      const blanks = nxtRow - endRow - 1; // blank rows currently between the two groups
      console.log('maintainGroupSpacing_: gap between ' + groups[g].prefix +
        ' (end=' + endRow + ') and ' + groups[g+1].prefix +
        ' (start=' + nxtRow + '): ' + blanks + ' blank row(s)');

      if (blanks < 3) {
        // ── Too few blanks → insert missing rows ──────────────────────────
        const needed = 3 - blanks;
        console.log('maintainGroupSpacing_: inserting ' + needed + ' row(s) after row ' + endRow);
        sheet.insertRowsAfter(endRow, needed);
        totalInserted += needed;
        for (let b = 1; b <= needed; b++) {
          // Clear full row: inserted rows inherit ALL formatting from the row above
          // (background, font size, font colour, data validations, etc.).
          const blankRange = sheet.getRange(endRow + b, 2, 1, COL.OWNER - 1);
          blankRange.setBackground(null)
                    .setFontFamily('Georgia').setFontSize(10).setFontWeight('normal').setFontColor(null)
                    .setHorizontalAlignment('left')
                    .clearContent().clearDataValidations();
          sheet.getRange(endRow + b, COL.ROW_NUM).setBackground(HEADER_BLUE).setValue('');
          sheet.setRowHeight(endRow + b, 20);
        }
        // Ensure any pre-existing blank rows in this gap also have correct styling
        for (let b = needed + 1; b <= 3; b++) {
          sheet.getRange(endRow + b, COL.ROW_NUM).setBackground(HEADER_BLUE).setValue('');
          sheet.setRowHeight(endRow + b, 20);
        }

      } else if (blanks > 3) {
        // ── Too many blanks → delete the extra rows ───────────────────────
        // Happens when a user clears a filename that was inside the separator zone.
        // Safe to delete because all rows between the two groups are blank.
        const extra = blanks - 3;
        console.log('maintainGroupSpacing_: removing ' + extra + ' extra blank row(s) after row ' + endRow);
        sheet.deleteRows(endRow + 1, extra);
        // nxtRow has shifted up; update the next group's firstRow so the
        // trailing-blanks check below sees the correct sheet state.
        groups[g + 1].firstRow -= extra;
      }

      // Clear any stale bottom borders in the separator zone (e.g. a border left
      // at position 4 when a file row was cleared, shifting everything up).
      // Then re-stamp exactly one border at position 3 (bottom of 3rd blank row).
      const sepSize = Math.max(blanks, 3);
      for (let sb = 1; sb <= sepSize; sb++) {
        sheet.getRange(endRow + sb, 1, 1, COL.OWNER)
             .setBorder(null, null, false, null, null, null);
      }
      sheet.getRange(endRow + 3, 1, 1, COL.OWNER)
           .setBorder(null, null, true, null, null, null,
                      SEPARATOR_RED, SpreadsheetApp.BorderStyle.SOLID_THICK);
    }

    // ── Trailing blank rows after the last group ─────────────────────────────
    // Re-read last row after any insertions above shifted everything down.
    const curLast = sheet.getLastRow();
    let trailFile = curLast;
    while (trailFile >= DATA_START &&
           !sheet.getRange(trailFile, COL.FILENAME).getValue()) {
      trailFile--;
    }
    if (trailFile >= DATA_START) {
      const trailBlanks = curLast - trailFile;
      console.log('maintainGroupSpacing_: trailing blanks after last file (row ' +
        trailFile + '): ' + trailBlanks);
      if (trailBlanks < 3) {
        for (let b = trailBlanks + 1; b <= 3; b++) {
          sheet.getRange(trailFile + b, COL.ROW_NUM).setBackground(HEADER_BLUE).setValue('');
          sheet.setRowHeight(trailFile + b, 20);
        }
      }
      // Clear stale borders then re-stamp at position 3.
      const trailSepSize = Math.max(trailBlanks, 3);
      for (let sb = 1; sb <= trailSepSize; sb++) {
        sheet.getRange(trailFile + sb, 1, 1, COL.OWNER)
             .setBorder(null, null, false, null, null, null);
      }
      sheet.getRange(trailFile + 3, 1, 1, COL.OWNER)
           .setBorder(null, null, true, null, null, null,
                      SEPARATOR_RED, SpreadsheetApp.BorderStyle.SOLID_THICK);
    }

    const summary = groups.map(g =>
      g.prefix + '[' + g.firstRow + '-' + g.lastRow + ']'
    ).join(', ');
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Groups: ' + (groups.length ? summary : 'none') +
      '\nRows inserted: ' + totalInserted,
      '🔧 Spacing Check', 6
    );
    console.log('maintainGroupSpacing_: done — ' + totalInserted + ' row(s) inserted');
    return totalInserted;

  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'maintainGroupSpacing_ ERROR:\n' + e.message,
      '❌ Spacing Error', 8
    );
    console.error('maintainGroupSpacing_ ERROR: ' + e.message + '\nStack: ' + e.stack);
    return 0;
  }
}

function updateSelectedInfo() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.REGISTRY);
    const r     = sheet.getActiveRange().getRow();
    if (r < DATA_START) return;
    const data         = getLevelsData();
    const templateList = getTemplateList();
    processAuditForRow(sheet, r, data.driveUrlLookup, data.validEntities, data.validDocs, templateList);
    maintainGroupSpacing_(sheet);
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

  // Strip _YYYYMMDD_vN suffix so the search finds all versions of the file
  const base      = extractKALBaseName(baseName);
  const escaped   = base.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const versionRe = new RegExp(escaped + '.*[vV](\\d+|FINAL)', 'i');
  let latestVerNum = -1;
  const details    = Object.assign({}, NOT_FOUND);

  try {
    const files = DriveApp.searchFiles("title contains '" + base + "'");
    while (files.hasNext()) {
      const file  = files.next();
      // Skip files whose base name doesn't match exactly — title contains '' is a
      // substring search and would otherwise pick up unrelated documents (e.g. a
      // Google Doc whose name merely contains the base name as a prefix).
      if (extractKALBaseName(file.getName()).toLowerCase() !== base.toLowerCase()) continue;
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

/**
 * Searches Drive for every version of the file on the active registry row
 * and shows them in a modal dialog, sorted newest-date-first.
 * Each file name and folder name is a clickable hyperlink.
 */
function showFileVersions() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  if (sheet.getName() !== SHEET.REGISTRY) {
    toast('Select a row in the Registry sheet first.', '⚠️ Warning', 4);
    return;
  }
  const r = sheet.getActiveRange().getRow();
  if (r < DATA_START) {
    toast('Select a file row (not the header).', '⚠️ Warning', 4);
    return;
  }

  const baseName = sheet.getRange(r, COL.FILENAME).getValue().toString().trim();
  if (!baseName) { toast('No filename in this row.', '⚠️ Warning', 4); return; }

  toast('Searching Drive for all versions…', '🔍 Versions', 30);

  const base  = extractKALBaseName(baseName);
  const found = [];
  const seenIds = new Set();

  const baseLower = base.toLowerCase();

  function collectFile_(file) {
    if (seenIds.has(file.getId())) return;
    const name = file.getName();
    // Match: exact base name OR base name followed by '_' (date/version suffix).
    // Simpler than extractKALBaseName regex — avoids false negatives on unusual patterns.
    const n = name.toLowerCase();
    if (n !== baseLower && !n.startsWith(baseLower + '_')) return;
    const fileId = file.getId();
    seenIds.add(fileId);

    const dateMatch = name.match(/_(\d{8})/);
    const verMatch  = name.match(/_v(\d+|FINAL)$/i);
    const date      = dateMatch ? dateMatch[1] : '';
    const ver       = verMatch  ? verMatch[1].toUpperCase() : '1';
    const verNum    = ver === 'FINAL' ? 9999 : parseInt(ver, 10);

    const par    = file.getParents();
    const folder = par.hasNext() ? par.next() : null;
    found.push({
      fileId,
      name,
      url:        file.getUrl(),
      date,
      ver,
      verNum,
      folderName: folder ? folder.getName() : '—',
      folderUrl:  folder ? folder.getUrl()  : ''
    });
  }

  const MAX_BFS_DEPTH = 5;

  function scanFolderBFS_(rootFolderId) {
    const queue = [{ id: rootFolderId, depth: 0 }];
    while (queue.length > 0) {
      const { id, depth } = queue.shift();
      let folder;
      try { folder = DriveApp.getFolderById(id); } catch (_) { continue; }
      try {
        const it = folder.getFiles();
        while (it.hasNext()) collectFile_(it.next());
      } catch (_) {}
      if (depth < MAX_BFS_DEPTH) {
        try {
          const subs = folder.getFolders();
          while (subs.hasNext()) queue.push({ id: subs.next().getId(), depth: depth + 1 });
        } catch (_) {}
      }
    }
  }

  try {
    // 1. BFS from the registered drive folder (driveUrlLookup) — primary strategy,
    //    proven to work for Shared Drives where searchFiles is unreliable.
    const driveCodeMatch = base.match(/^([A-Za-z]{2,4})-/);
    const driveCode = driveCodeMatch ? driveCodeMatch[1].toUpperCase() : null;
    let drivedFolderSearched = false;
    if (driveCode) {
      try {
        const levelsData  = getLevelsData();
        const driveUrl    = levelsData.driveUrlLookup[driveCode];
        const driveFolderId = driveUrl ? getIdFromUrl(driveUrl) : null;
        if (driveFolderId) {
          console.log('showFileVersions: BFS from drive folder for code=' + driveCode + ' id=' + driveFolderId);
          scanFolderBFS_(driveFolderId);
          drivedFolderSearched = true;
        }
      } catch (e) { console.warn('showFileVersions: drive folder BFS failed: ' + e.message); }
    }

    // 2. BFS from the archive folder (Settings!C2) — older versions moved by Promote/Archive.
    try {
      const archiveId = getIdFromUrl(getUrlFromCell(SHEET.SETTINGS, 'C2'));
      if (archiveId) {
        console.log('showFileVersions: BFS from archive folder id=' + archiveId);
        scanFolderBFS_(archiveId);
      }
    } catch (e) { console.warn('showFileVersions: archive BFS failed: ' + e.message); }

    // 3. BFS from the current file's own folder via col G link — catches files
    //    whose drive code folder was not reachable or is a different structure.
    try {
      const linkUrl = sheet.getRange(r, COL.LINK).getRichTextValue().getLinkUrl();
      const fileId  = getIdFromUrl(linkUrl);
      if (fileId) {
        const parents = DriveApp.getFileById(fileId).getParents();
        while (parents.hasNext()) scanFolderBFS_(parents.next().getId());
      }
    } catch (_) {}

    // 4. Fallback: searchFiles by prefix (My Drive files not under any registered folder).
    try {
      const prefixMatch  = base.match(/^([A-Za-z]{2,4}-)/);
      const searchPrefix = prefixMatch ? prefixMatch[1].toUpperCase() : base;
      const iter = DriveApp.searchFiles("title contains '" + searchPrefix + "' and trashed = false");
      while (iter.hasNext()) collectFile_(iter.next());
    } catch (_) {}

    console.log('showFileVersions: base="' + base + '" found=' + found.length);
  } catch (e) {
    toast('Drive search failed: ' + e.message, '❌ Error', 6);
    return;
  }

  if (!found.length) {
    SpreadsheetApp.getUi().showModalDialog(
      HtmlService.createHtmlOutput(
        '<p style="font-family:Arial;padding:16px;font-size:13px">No versions found in Drive for:<br><b>' +
        base + '</b></p>'
      ).setWidth(400).setHeight(100),
      '📋 File Versions'
    );
    return;
  }

  // Newest date first; within same date, highest version first
  found.sort((a, b) => {
    if (b.date !== a.date) return b.date.localeCompare(a.date);
    return b.verNum - a.verNum;
  });

  const latestVerIdx = found.reduce((best, v, i) => v.verNum > found[best].verNum ? i : best, 0);

  // ── Feature 6: Version gap detection ─────────────────────────────────────
  // Build sorted-by-verNum list (excluding FINAL=9999) to detect gaps
  const verNums = found
    .map((v, i) => ({ verNum: v.verNum, idx: i }))
    .filter(v => v.verNum !== 9999)
    .sort((a, b) => a.verNum - b.verNum);

  const gapAfterIdx  = new Set(); // original found[] index: gap AFTER this entry
  const gapBeforeIdx = new Set(); // original found[] index: gap BEFORE this entry
  for (let g = 0; g < verNums.length - 1; g++) {
    if (verNums[g + 1].verNum - verNums[g].verNum > 1) {
      const missingVer = verNums[g].verNum + 1;
      // Mark the entries adjacent to the gap
      gapAfterIdx.add(verNums[g].idx);
      gapBeforeIdx.add(verNums[g + 1].idx);
    }
  }

  // ── Feature 9: Load changelog notes from Changelog sheet ─────────────────
  const changelogMap = {};
  try {
    const clSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Changelog');
    if (clSheet && clSheet.getLastRow() >= 2) {
      const clData = clSheet.getRange(2, 1, clSheet.getLastRow() - 1, 2).getValues();
      clData.forEach(row => { if (row[0]) changelogMap[row[0]] = row[1] || ''; });
    }
  } catch (_) {}

  const rows = found.map((v, i) => {
    const isLatestDate = i === 0;
    const isLatestVer  = i === latestVerIdx;

    // Gap warning icons
    let gapWarning = '';
    if (gapAfterIdx.has(i)) {
      const nextVer = v.verNum + 1;
      gapWarning += ` <span title="Gap: v${nextVer} is missing" style="cursor:help">⚠️</span>`;
    }
    if (gapBeforeIdx.has(i)) {
      const prevVer = v.verNum - 1;
      gapWarning = `<span title="Gap: v${prevVer} is missing" style="cursor:help">⚠️</span> ` + gapWarning;
    }

    const badges = (isLatestDate ? ' <span class="badge-date">Latest Date</span>' : '') +
                   (isLatestVer  ? ' <span class="badge-ver">Latest Version</span>' : '');
    const nameCell   = `<a href="${v.url}" target="_blank">${v.name}</a>${badges}${gapWarning}`;
    const folderCell = v.folderUrl
      ? `<a href="${v.folderUrl}" target="_blank">${v.folderName}</a>`
      : v.folderName;
    const rowClass = (isLatestDate || isLatestVer) ? ' class="row-marked"' : '';

    // Feature 1: Promote button on Latest Version row
    const promoteBtn = isLatestVer && v.ver !== 'FINAL'
      ? `<button class="promote-btn" data-id="${v.fileId}" data-name="${v.name.replace(/"/g,'&quot;')}" title="Promote to vFINAL">🏁</button>`
      : '';

    // Feature 9: Notes (changelog) cell — contenteditable
    const existingNote = (changelogMap[v.fileId] || '').replace(/</g,'&lt;').replace(/>/g,'&gt;');

    return `<tr data-id="${v.fileId}"${rowClass}>
      <td>${nameCell}</td>
      <td style="white-space:nowrap;text-align:center">${v.date || '—'}</td>
      <td style="white-space:nowrap;text-align:center">v${v.ver}</td>
      <td>${folderCell}</td>
      <td class="notes-cell" contenteditable="true" data-id="${v.fileId}" style="min-width:120px;font-size:11px;color:#555;outline:none;cursor:text">${existingNote}</td>
      <td style="text-align:center;white-space:nowrap;width:60px">
        ${promoteBtn}
        <button class="del-btn" data-id="${v.fileId}" data-name="${v.name.replace(/"/g,'&quot;')}" title="Move to trash">
          <svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor">
            <path d="M5.5 5.5A.5.5 0 0 1 6 6v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5m2.5 0a.5.5 0 0 1 .5.5v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5m3 .5a.5.5 0 0 0-1 0v6a.5.5 0 0 0 1 0z"/>
            <path d="M14.5 3a1 1 0 0 1-1 1H13v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V4h-.5a1 1 0 0 1-1-1V2a1 1 0 0 1 1-1H6a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1h3.5a1 1 0 0 1 1 1zM4.118 4 4 4.059V13a1 1 0 0 0 1 1h6a1 1 0 0 0 1-1V4.059L11.882 4zM2.5 3h11V2h-11z"/>
          </svg>
        </button>
      </td>
    </tr>`;
  }).join('');

  const html = HtmlService.createHtmlOutput(`
    <style>
      *{box-sizing:border-box}
      body{font-family:Arial,sans-serif;padding:14px 16px;margin:0;font-size:13px}
      h3{margin:0 0 3px;font-size:15px;color:#111184}
      .sub{margin:0 0 12px;font-size:11px;color:#888}
      table{width:100%;border-collapse:collapse}
      th,td{padding:7px 11px;text-align:left;border-bottom:1px solid #e8e8e8;vertical-align:middle}
      th{background:#111184;color:#fff;font-size:12px;font-weight:600;white-space:nowrap}
      tr:last-child td{border-bottom:none}
      tr:hover td{background:#f4f6ff}
      a{color:#1155CC;text-decoration:none}
      a:hover{text-decoration:underline}
      .del-btn{background:none;border:none;cursor:pointer;color:#999;padding:2px 4px;border-radius:3px;line-height:1;display:inline-flex;align-items:center}
      .del-btn:hover{color:#c0392b;background:#fdecea}
      .del-btn:disabled{opacity:.35;cursor:default}
      .promote-btn{background:none;border:1px solid #1a7f37;color:#1a7f37;cursor:pointer;padding:2px 6px;border-radius:3px;font-size:12px;margin-right:4px}
      .promote-btn:hover{background:#d9ead3}
      .promote-btn:disabled{opacity:.35;cursor:default}
      tr.deleted td{opacity:.4;text-decoration:line-through}
      tr.row-marked td{background:#f0f4ff}
      tr.row-marked:hover td{background:#e4ecff}
      .badge-date,.badge-ver{display:inline-block;margin-left:6px;padding:1px 7px;font-size:10px;font-weight:700;letter-spacing:.4px;border-radius:10px;vertical-align:middle;text-transform:uppercase}
      .badge-date{background:#111184;color:#fff}
      .badge-ver{background:#1a7f37;color:#fff}
      .notes-cell:focus{background:#fffde7;border:1px solid #f9ab00;border-radius:3px}
      .notes-cell:empty:before{content:"Add note…";color:#bbb;font-style:italic;pointer-events:none}
    </style>
    <h3>📋 All Versions</h3>
    <p class="sub" id="sub-line">${base}&nbsp;&nbsp;·&nbsp;&nbsp;${found.length} version(s) found in Drive</p>
    <table id="ver-table">
      <tr>
        <th>File Name</th>
        <th>Date</th>
        <th>Version</th>
        <th>Drive Location</th>
        <th>Notes</th>
        <th></th>
      </tr>
      ${rows}
    </table>
    <script>
      var remaining = ${found.length};

      // ── Delete handler ────────────────────────────────────────────────────
      document.getElementById('ver-table').addEventListener('click', function(e) {
        var btn = e.target.closest('.del-btn');
        if (!btn) return;
        var fileId   = btn.dataset.id;
        var fileName = btn.dataset.name;
        if (!confirm('Move "' + fileName + '" to trash?')) return;
        btn.disabled = true;
        google.script.run
          .withSuccessHandler(function() {
            var row = document.querySelector('tr[data-id="' + fileId + '"]');
            if (row) row.classList.add('deleted');
            remaining--;
            document.getElementById('sub-line').textContent =
              '${base}  ·  ' + remaining + ' version(s) found in Drive';
          })
          .withFailureHandler(function(err) {
            btn.disabled = false;
            alert('Could not delete: ' + err.message);
          })
          .deleteVersionFile_(fileId);
      });

      // ── Promote handler (Feature 1) ───────────────────────────────────────
      document.getElementById('ver-table').addEventListener('click', function(e) {
        var btn = e.target.closest('.promote-btn');
        if (!btn) return;
        var fileId   = btn.dataset.id;
        var fileName = btn.dataset.name;
        if (!confirm('Promote "' + fileName + '" to vFINAL?')) return;
        btn.disabled = true;
        btn.textContent = '…';
        google.script.run
          .withSuccessHandler(function() {
            google.script.host.close();
          })
          .withFailureHandler(function(err) {
            btn.disabled = false;
            btn.textContent = '🏁';
            alert('Promote failed: ' + err.message);
          })
          .promoteVersionFromDialog_(fileId);
      });

      // ── Notes blur-save handler (Feature 9) ──────────────────────────────
      document.getElementById('ver-table').addEventListener('blur', function(e) {
        var cell = e.target;
        if (!cell.classList.contains('notes-cell')) return;
        var fileId = cell.dataset.id;
        var note   = cell.innerText.trim();
        google.script.run
          .withFailureHandler(function(err) {
            console.warn('Note save failed: ' + err.message);
          })
          .saveVersionChangelog_(fileId, note);
      }, true);
    </script>
  `).setWidth(980).setHeight(Math.min(160 + found.length * 38, 540));

  SpreadsheetApp.getUi().showModalDialog(html, '📋 Versions — ' + base);
}

/** Moves a Drive file to trash. Called from the Show All Versions modal. */
function deleteVersionFile_(fileId) {
  if (!fileId) throw new Error('No file ID provided.');
  DriveApp.getFileById(fileId).setTrashed(true);
}

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

    const newFile = srcFile.makeCopy(finalName, DriveApp.getFolderById(destId));
    srcFile.moveTo(DriveApp.getFolderById(archiveId));
    updateSelectedInfo();
    toast('"' + finalName + '" created in destination drive. Original archived.', '🏁 Promoted', 6);

    // ── Feature 8: Email notification on vFINAL promotion ─────────────────
    const ownerEmail = sheet.getRange(r, COL.OWNER).getValue().toString().trim();
    if (ownerEmail && ownerEmail.includes('@')) {
      try {
        const promotionDate = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm');
        MailApp.sendEmail({
          to:      ownerEmail,
          subject: '[KAL] ' + finalName + ' promoted to vFINAL',
          body:    'Hello,\n\nThe following file has been promoted to vFINAL:\n\n' +
                   'File Name: ' + finalName + '\n' +
                   'Drive Link: ' + newFile.getUrl() + '\n' +
                   'Date: ' + promotionDate + '\n\n' +
                   'The previous version has been moved to the archive folder.\n\n' +
                   'KAL File Registry'
        });
        console.log('promoteToFinalAndMove: notification sent to ' + ownerEmail);
      } catch (mailErr) {
        console.warn('promoteToFinalAndMove: email notification failed: ' + mailErr.message);
      }
    }
  } catch (e) { toast('Promote failed: ' + e.message, '❌ Error', 6); console.error('promoteToFinalAndMove: ' + e.message); }
}

function removeSelectedFile() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const r     = sheet.getActiveRange().getRow();
  if (r < DATA_START) { toast('Select a data row first.', '⚠️ Warning', 4); return; }

  const selectedName = sheet.getRange(r, COL.FILENAME).getValue().toString().trim();
  if (selectedName.startsWith(PINNED_FILE_BASE)) {
    toast('The master registry document is protected and cannot be removed.', '🔒 Protected', 5);
    return;
  }

  let fileUrl = null;
  try { fileUrl = sheet.getRange(r, COL.LINK).getRichTextValue().getLinkUrl(); } catch (_) {}
  if (!fileUrl) { toast('No file link found in this row.', '⚠️ Warning', 4); return; }

  const fileId = getIdFromUrl(fileUrl);
  if (!fileId) { toast('Could not parse a file ID from the link.', '🛑 Error', 5); return; }

  const response = ui.alert('⚠️ Warning', 'Move THIS version to Trash?', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    sheet.deleteRow(r);
    maintainGroupSpacing_(sheet);
    toast('Version moved to Trash and removed from registry.', '🗑️ Removed', 4);
  } catch (e) { toast('Could not trash the file: ' + e.message, '❌ Error', 5); }
}

function removeAllVersions() {
  const ui       = SpreadsheetApp.getUi();
  const sheet    = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const r        = sheet.getActiveRange().getRow();
  if (r < DATA_START) return;
  const baseName = sheet.getRange(r, COL.FILENAME).getValue().toString().trim();
  if (!baseName) return;
  if (baseName.startsWith(PINNED_FILE_BASE)) {
    toast('The master registry document is protected and cannot be removed.', '🔒 Protected', 5);
    return;
  }

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

  sheet.deleteRow(r);
  maintainGroupSpacing_(sheet);
  toast('Trashed ' + trashed + ' file(s) and removed from registry.' + (failed ? ' (' + failed + ' failed)' : ''), '☢️ Done', 5);
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
    'format=pdf',
    'id='          + ss.getId(),
    'gid='         + sheet.getSheetId(),
    'portrait=false', 'fitw=true', 'size=A3',
    'top_margin=0.25', 'bottom_margin=0.25', 'left_margin=0.25', 'right_margin=0.25',
    'sheetnames=false', 'printtitle=false', 'pagenumbers=false', 'gridlines=false', 'fzr=false'
  ].join('&');

  try {
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    const code = response.getResponseCode();
    if (code !== 200) throw new Error('Google export returned HTTP ' + code + '. Check that the sheet is not empty and you have access.');

    const blob    = response.getBlob().setName(pdfName).setContentType('application/pdf');
    const destId  = getIdFromUrl(getUrlFromCell(SHEET.SETTINGS, 'A2'));
    const pdfFile = destId ? DriveApp.getFolderById(destId).createFile(blob) : DriveApp.createFile(blob);

    const html = HtmlService.createHtmlOutput(
      '<p style="font-family:Arial;padding:12px;font-size:13px">PDF ready: ' +
      '<a href="' + pdfFile.getUrl() + '" target="_blank"><b>' + pdfName + '</b></a></p>' +
      '<script>setTimeout(()=>google.script.host.close(),9000)</script>'
    ).setWidth(420).setHeight(70);
    SpreadsheetApp.getUi().showModalDialog(html, '📤 Export Complete');
  } catch (e) {
    toast('Export failed: ' + e.message, '❌ Error', 6);
    console.error('exportRegistryToPDF: ' + e.message);
  }
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

// ── 9. NOTIFICATIONS ─────────────────────────────────────────────────────────

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
    const richB    = sheet.getRange(3, 2, lastRow - 2, 1).getRichTextValues(); // col B folder links
    const drives   = [], entities = [], docTypes = [], examples = [];

    values.forEach((row, i) => {
      if (row[0]) {
        const folderLink = richB[i][0] ? richB[i][0].getLinkUrl() : null;
        const folderName = String(row[1] || '').trim();
        drives.push({ code: row[0], name: row[2], folderName, folderLink });
      }
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

// ── Registry Rebuild from Drive ───────────────────────────────────────────────

/**
 * Rebuilds the registry by scanning Google Drive for KAL-convention files.
 *
 * Groups files by drive-code prefix in REBUILD_PREFIX_ORDER (OP, KAL, PC …),
 * then any other prefixes found, sorted alphabetically.
 * Between groups: 3 blank rows + 1 full blue separator row.
 * Last row is also a blue separator.
 * Row 1 and col A are frozen; header row is formatted to match the sheet design.
 *
 * ⚠️  This clears all existing data rows. Confirm before running.
 */
function rebuildRegistryFromDrive() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.REGISTRY);
  if (!sheet) { toast('Registry sheet not found.', '❌ Error', 5); return; }

  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    '🔁 Rebuild Registry from Drive',
    'This will CLEAR all existing rows and repopulate from Google Drive.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  toast('Scanning Drive — this may take a moment…', '🔍 Rebuilding', 30);

  // 1. Format header row
  rebuildFormatHeader_(sheet);

  // 2. Clear ALL rows from DATA_START to the sheet's maximum row count.
  //    Using getMaxRows() (not getLastRow()) ensures border-only rows — which have
  //    no cell content and are therefore invisible to getLastRow() — are also wiped,
  //    preventing ghost red lines from surviving across rebuilds.
  const maxRows = sheet.getMaxRows();
  if (maxRows >= DATA_START) {
    sheet.getRange(DATA_START, 1, maxRows - DATA_START + 1, COL.OWNER).clear();
  }

  // 3. Collect Drive files grouped by drive-code prefix
  const groups = rebuildCollectGroups_();

  // Collect all KAL file IDs so the non-KAL scan can skip them
  const kalFileIds = new Set();
  Object.values(groups).forEach(arr => arr.forEach(f => kalFileIds.add(f.getId())));

  // 4. Render order: defined prefixes first, then others sorted
  const defined = REBUILD_PREFIX_ORDER.filter(p => groups[p] && groups[p].length);
  const others  = Object.keys(groups)
                        .filter(p => !REBUILD_PREFIX_ORDER.includes(p))
                        .sort();
  const renderOrder = [...defined, ...others];

  if (!renderOrder.length) {
    toast('No KAL-convention files found in Drive.', 'ℹ️ Info', 5);
    return;
  }

  // 5. Write file rows; red border line between groups (no colored row)
  let r = DATA_START;
  let lastFileRow = -1;
  const blankRows = []; // track blank separator rows to give them a thin height

  // Pin the master document to the very top before all groups.
  let pinnedFile = null;
  for (const prefix of Object.keys(groups)) {
    const idx = groups[prefix].findIndex(f => f.getName().startsWith(PINNED_FILE_BASE));
    if (idx !== -1) {
      pinnedFile = groups[prefix].splice(idx, 1)[0];
      if (!groups[prefix].length) delete groups[prefix];
      break;
    }
  }
  if (pinnedFile) {
    rebuildWriteFileRow_(sheet, r, pinnedFile);
    lastFileRow = r++;
    // No blank rows after pinned — master doc sits flush above the first group
  }

  // Rebuild renderOrder in case a group became empty after pinned extraction
  const definedUp  = REBUILD_PREFIX_ORDER.filter(p => groups[p] && groups[p].length);
  const othersUp   = Object.keys(groups)
                           .filter(p => !REBUILD_PREFIX_ORDER.includes(p))
                           .sort();
  const finalOrder = [...definedUp, ...othersUp];

  finalOrder.forEach((prefix, idx) => {
    if (idx > 0) {
      for (let b = 0; b < 3; b++) blankRows.push(r + b);
      r += 3; // 3 blank rows between groups
      // Red BOTTOM border on the last blank row (r-1) — placed on a thin row
      // that never gets content written to it, so the border always stays visible.
      // Visually this line sits right above the first file row of the new group.
      sheet.getRange(r - 1, 1, 1, COL.OWNER)
           .setBorder(null, null, true, null, null, null,
                      SEPARATOR_RED, SpreadsheetApp.BorderStyle.SOLID_THICK);
    }
    (groups[prefix] || []).forEach(file => {
      rebuildWriteFileRow_(sheet, r, file);
      lastFileRow = r++;
    });
  });
  // Non-KAL files section: files in Settings!E2 Root Folder that don't follow KAL convention
  const nonKalFiles = rebuildCollectNonKalFiles_(kalFileIds);
  if (nonKalFiles.length > 0 && lastFileRow >= DATA_START) {
    for (let b = 0; b < 3; b++) blankRows.push(r + b);
    r += 3;
    sheet.getRange(r - 1, 1, 1, COL.OWNER)
         .setBorder(null, null, true, null, null, null,
                    SEPARATOR_RED, SpreadsheetApp.BorderStyle.SOLID_THICK);
    nonKalFiles.forEach(file => {
      rebuildWriteNonKalRow_(sheet, r, file);
      lastFileRow = r++;
    });
  }

  // 3 trailing blank rows with navy col A + red bottom border to close the registry
  if (lastFileRow >= DATA_START) {
    for (let b = 1; b <= 3; b++) {
      blankRows.push(lastFileRow + b);
      sheet.getRange(lastFileRow + b, COL.ROW_NUM)
           .setBackground(HEADER_BLUE).setValue('');
    }
    sheet.getRange(lastFileRow + 3, 1, 1, COL.OWNER)
         .setBorder(null, null, true, null, null, null,
                    SEPARATOR_RED, SpreadsheetApp.BorderStyle.SOLID_THICK);
  }

  // 6. Freeze, set specific column widths
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  [70, 260, 360, 65, 65, 110, 50, 50, 188, 110, 130, 420, 80]
    .forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  renumberAllRows_(sheet);

  // Clip Abstract column to prevent long formula text from expanding rows.
  // Row heights are set by updateAllInfo() below (21 px file rows, 14 px blank rows).
  const lastWritten = sheet.getLastRow();
  if (lastWritten >= DATA_START) {
    sheet.getRange(DATA_START, COL.ABSTRACT, lastWritten - DATA_START + 1, 1)
         .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  const total = renderOrder.reduce((n, p) => n + (groups[p] || []).length, 0);
  toast('Rebuild complete — ' + total + ' files. Running full audit…', '✅ Rebuilt', 4);

  // Apply Georgia/10 to all data rows so newly added rows inherit the same default.
  applySheetFont_(sheet);

  // Auto-run full audit so KAL check, folder links and row colours are filled immediately.
  // updateAllInfo calls maintainGroupSpacing_ which may call insertRowsAfter — inserting
  // rows causes Google Sheets to reset floating images, so the logo must be (re)inserted
  // AFTER updateAllInfo finishes.
  updateAllInfo();

  // Insert logo last — after all row operations — so it appears immediately without a
  // page refresh.  rebuildInsertLogo_ removes any stale image before inserting a fresh one.
  rebuildInsertLogo_(sheet);

  // Activate cell A1 to force the Google Sheets UI to re-render the full sheet,
  // making col A numbers and the logo visible without a manual page refresh.
  try { sheet.getRange(1, 1).activate(); } catch (_) {}

  // =AI() formulas are written per-row by processAuditForRow (only on empty cells).
  // No separate refresh pass needed after rebuild.
}

/** Searches Drive for all KAL-convention files; returns {PREFIX: [DriveFile]} map. */
function rebuildCollectGroups_() {
  const groups = {};
  const seen   = new Set();
  const searchPrefixes = [...new Set([...REBUILD_PREFIX_ORDER, 'LP', 'CA', 'SA', 'RA', 'GA', 'MA'])];

  searchPrefixes.forEach(prefix => {
    try {
      const iter = DriveApp.searchFiles(
        `title contains '${prefix}-' and trashed = false`
      );
      while (iter.hasNext()) {
        const f = iter.next();
        if (seen.has(f.getId())) continue;
        const n = f.getName();
        // Must match KAL pattern: UPPERCASE-UPPERCASE_... (underscore required)
        const m = n.match(/^([A-Z]{2,4})-[A-Z]/);
        if (!m || !n.includes('_')) continue;
        seen.add(f.getId());
        const key = m[1];
        if (!groups[key]) groups[key] = [];
        groups[key].push(f);
      }
    } catch (_) {}
  });

  // Sort: entity code → doc type → base filename (alphabetical) → version desc (newest first)
  Object.values(groups).forEach(arr =>
    arr.sort((a, b) => {
      const na = a.getName(), nb = b.getName();
      const pa = na.split('_'),  pb = nb.split('_');
      const ea = (pa[0] || '').split('-')[1] || '';
      const eb = (pb[0] || '').split('-')[1] || '';
      if (ea !== eb) return ea.localeCompare(eb);
      const da = pa[1] || '', db = pb[1] || '';
      if (da !== db) return da.localeCompare(db);
      const ba = extractKALBaseName(na), bb = extractKALBaseName(nb);
      if (ba !== bb) return ba.localeCompare(bb);   // same entity+doctype → alphabetical
      return rebuildExtractVer_(nb) - rebuildExtractVer_(na); // same file → newest first
    })
  );

  // Deduplicate: the sort put the newest version first, so keep only the first
  // occurrence of each base name — older versions (v1, v2…) are discarded.
  Object.keys(groups).forEach(key => {
    const seen = new Set();
    groups[key] = groups[key].filter(f => {
      const base = extractKALBaseName(f.getName());
      if (seen.has(base)) return false;
      seen.add(base);
      return true;
    });
  });

  return groups;
}

/** Extracts numeric version from a KAL filename for sort comparison. */
function rebuildExtractVer_(name) {
  if (/vFINAL$/i.test(name)) return 9999;
  const m = name.match(/_v(\d+)(?:_\d{8})?$/i);
  return m ? parseInt(m[1], 10) : 1;
}

/** Writes one Drive file's data into a registry row. */
function rebuildWriteFileRow_(sheet, r, driveFile) {
  const name   = driveFile.getName();
  const url    = driveFile.getUrl();
  const mime   = driveFile.getMimeType();
  const par    = driveFile.getParents();
  const folder = par.hasNext() ? par.next() : null;

  // Parse filename: DRIVECODE-ENTITY_DOCTYPE_NameParts[_YYYYMMDD][_vN|_vFINAL]
  const parts   = name.split('_');
  const prefix  = parts[0] || '';               // e.g. "OP-KAL"

  // Human-readable: parts after docType, filtering out date (8 digits) and version tokens
  const rawDesc = parts.slice(2)
    .filter(p => !/^\d{8}$/.test(p) && !/^v\d+$/i.test(p) && !/^vFINAL$/i.test(p))
    .join(' ')
    .trim();
  const humanDesc = rawDesc
    .replace(/([a-z\d])([A-Z])/g,    '$1 $2')
    .replace(/([A-Z]+)([A-Z][a-z])/g,'$1 $2')
    .replace(/([a-zA-Z])(\d)/g,      '$1 $2')
    .replace(/(\d)([a-zA-Z])/g,      '$1 $2')
    .trim();

  // Version
  const verMatch = name.match(/_v(\d+|FINAL)$/i);
  const version  = verMatch ? verMatch[1] : '1';

  // For Who = entity code: part after '-' in the drive-prefix segment
  const entMatch = prefix.match(/-(.+)/);
  const forWho   = entMatch ? entMatch[1] : 'ALL';

  // Folder link
  const folderName = folder ? folder.getName() : '';
  const folderUrl  = folder ? folder.getUrl()  : '';

  // Write plain values and reset background to white (row may follow a blue separator)
  const rowRange = sheet.getRange(r, 1, 1, COL.OWNER);
  rowRange.setBackground(null)
          .setFontColor(null)
          .setFontWeight('normal')
          .setHorizontalAlignment('left')
          .setValues([[
            '',          // A: row number
            humanDesc,              // B: human-readable description
            extractKALBaseName(name), // C: filename without date/version suffix
            formatMimeType(mime), // D: file type
            version,     // E: current version
            '',          // F: current folder (hyperlink set below)
            '',          // G: link (hyperlink set below)
            forWho,      // H: for who
            '',          // I: KAL name check
            '',          // J: destination drive
            '',          // K: preferred template
            '',          // L: abstract
            ''           // M: owner
          ]]);

  // Column A: navy background, white bold, centered row number
  sheet.getRange(r, COL.ROW_NUM)
       .setBackground(HEADER_BLUE)
       .setFontColor('#ffffff')
       .setFontWeight('bold')
       .setHorizontalAlignment('center');

  // Center Current Version (col E)
  sheet.getRange(r, COL.VERSION).setHorizontalAlignment('center');

  // Hyperlinks
  if (folderUrl) {
    const safeFolder = folderName.replace(/"/g, "'");
    sheet.getRange(r, COL.FOLDER)
         .setFormula(`=HYPERLINK("${folderUrl}","${safeFolder}")`);
  }
  sheet.getRange(r, COL.LINK)
       .setFormula(`=HYPERLINK("${url}","Link")`);
}

/**
 * Refreshes only the non-KAL file section at the bottom of the registry
 * without running a full Rebuild Registry from Drive.  All KAL rows are
 * left untouched; the non-KAL block (everything below the last KAL group's
 * three separator blank rows) is cleared and re-populated from the Root
 * Folder defined in Settings!E2.
 */
function refreshNonKalSection() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.REGISTRY);
  if (!sheet) { toast('Registry sheet not found.', '❌ Error', 5); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) {
    toast('Registry is empty — run Rebuild Registry from Drive first.', '⚠️ Warning', 5);
    return;
  }

  toast('Scanning Drive for non-KAL files…', '🔍 Refreshing', 60);

  const n       = lastRow - DATA_START + 1;
  const KAL_RE  = /^[A-Z]{2,4}-[A-Z]/;

  // 1. Scan col C to find the last KAL row and collect existing KAL file IDs
  //    (so the Drive scan can skip files already tracked in the registry).
  const fileNames = sheet.getRange(DATA_START, COL.FILENAME, n, 1).getValues();
  const linkRich  = sheet.getRange(DATA_START, COL.LINK,     n, 1).getRichTextValues();

  const kalFileIds = new Set();
  let lastKalRow   = -1;

  fileNames.forEach((row, i) => {
    const fn = (row[0] || '').toString().trim();
    if (!fn || !KAL_RE.test(fn) || !fn.includes('_')) return; // blank or non-KAL — skip
    lastKalRow = DATA_START + i;
    try {
      const url = linkRich[i][0] ? linkRich[i][0].getLinkUrl() : null;
      const id  = getIdFromUrl(url);
      if (id) kalFileIds.add(id);
    } catch (_) {}
  });

  if (lastKalRow < DATA_START) {
    toast('No KAL rows found — run Rebuild Registry from Drive first.', '⚠️ Warning', 5);
    return;
  }

  // 2. Clear the non-KAL block (lastKalRow+4 onwards).
  //    Rows lastKalRow+1..+3 are the separator blanks — they are preserved and
  //    re-stamped below, so any stale content there is corrected without a full clear.
  const maxRows   = sheet.getMaxRows();
  const clearFrom = lastKalRow + 4;
  if (maxRows >= clearFrom) {
    sheet.getRange(clearFrom, 1, maxRows - clearFrom + 1, COL.OWNER).clear();
  }

  // 3. Re-stamp the three separator blank rows and their red bottom border.
  for (let b = 1; b <= 3; b++) {
    sheet.getRange(lastKalRow + b, COL.ROW_NUM).setBackground(HEADER_BLUE).setValue('');
  }
  sheet.getRange(lastKalRow + 3, 1, 1, COL.OWNER)
       .setBorder(null, null, true, null, null, null,
                  SEPARATOR_RED, SpreadsheetApp.BorderStyle.SOLID_THICK);

  // 4. Collect non-KAL files from Drive (excludes KAL IDs and "Help" prefix).
  const nonKalFiles = rebuildCollectNonKalFiles_(kalFileIds);

  // 5. Write non-KAL rows and close with three trailing blank rows + border.
  if (nonKalFiles.length > 0) {
    let r = lastKalRow + 4;
    let lastFileRow = lastKalRow;
    nonKalFiles.forEach(file => {
      rebuildWriteNonKalRow_(sheet, r, file);
      lastFileRow = r++;
    });
    for (let b = 1; b <= 3; b++) {
      sheet.getRange(lastFileRow + b, COL.ROW_NUM).setBackground(HEADER_BLUE).setValue('');
    }
    sheet.getRange(lastFileRow + 3, 1, 1, COL.OWNER)
         .setBorder(null, null, true, null, null, null,
                    SEPARATOR_RED, SpreadsheetApp.BorderStyle.SOLID_THICK);
  }
  // If no non-KAL files the separator border at lastKalRow+3 closes the registry.

  // 6. Apply font, re-audit non-KAL rows (sets col I to "Non-KAL"), renumber, flush.
  applySheetFont_(sheet);
  renumberAllRows_(sheet);
  SpreadsheetApp.flush();

  const msg = nonKalFiles.length === 0
    ? 'No non-KAL files found in Root Folder.'
    : nonKalFiles.length + ' non-KAL file(s) listed.';
  toast(msg, '✅ Non-KAL Refreshed', 5);
}

/**
 * Recursively searches the Root Folder (Settings!E2) and returns all files
 * that do NOT follow KAL naming convention and are not already in the KAL
 * registry.  Limited to MAX_DEPTH folder levels to avoid execution timeouts.
 */
function rebuildCollectNonKalFiles_(kalFileIds) {
  const rootUrl = getUrlFromCell(SHEET.SETTINGS, 'E2');
  if (!rootUrl) {
    console.log('Settings!E2 (Root Folder) is empty — skipping non-KAL scan');
    return [];
  }
  const rootId = getIdFromUrl(rootUrl);
  if (!rootId) { console.warn('Could not extract folder ID from Settings!E2'); return []; }

  let rootFolder;
  try { rootFolder = DriveApp.getFolderById(rootId); }
  catch (e) { console.warn('Root Folder not accessible: ' + e.message); return []; }

  const KAL_RE  = /^[A-Z]{2,4}-[A-Z]/;
  const MAX_DEPTH = 5;
  const nonKal  = [];
  const queue   = [{ folder: rootFolder, depth: 0 }];

  while (queue.length > 0) {
    const { folder, depth } = queue.shift();
    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (kalFileIds.has(f.getId())) continue;
      const name = f.getName();
      if (name.startsWith('Help')) continue;
      if (!KAL_RE.test(name) || !name.includes('_')) nonKal.push(f);
    }
    if (depth < MAX_DEPTH) {
      const subs = folder.getFolders();
      while (subs.hasNext()) queue.push({ folder: subs.next(), depth: depth + 1 });
    }
  }

  nonKal.sort((a, b) => a.getName().localeCompare(b.getName()));
  console.log('Non-KAL scan: found ' + nonKal.length + ' file(s) in Root Folder');
  return nonKal;
}

/** Writes one non-KAL file's data into a registry row. */
function rebuildWriteNonKalRow_(sheet, r, driveFile) {
  const name   = driveFile.getName();
  const url    = driveFile.getUrl();
  const mime   = driveFile.getMimeType();
  const par    = driveFile.getParents();
  const folder = par.hasNext() ? par.next() : null;

  const folderName = folder ? folder.getName() : '';
  const folderUrl  = folder ? folder.getUrl()  : '';

  const rowRange = sheet.getRange(r, 1, 1, COL.OWNER);
  rowRange.setBackground(null)
          .setFontColor(null)
          .setFontWeight('normal')
          .setHorizontalAlignment('left')
          .setValues([[
            '',                   // A: row number (set by renumberAllRows_)
            name,                 // B: human-readable name
            name,                 // C: base file name
            formatMimeType(mime), // D: file type
            '',                   // E: version
            '',                   // F: folder (hyperlink below)
            '',                   // G: link (hyperlink below)
            '',                   // H: for who
            'Non-KAL',            // I: KAL check
            '',                   // J: destination drive
            '',                   // K: preferred template
            '',                   // L: abstract
            ''                    // M: owner
          ]]);

  sheet.getRange(r, COL.ROW_NUM)
       .setBackground(HEADER_BLUE)
       .setFontColor('#ffffff')
       .setFontWeight('bold')
       .setHorizontalAlignment('center');

  sheet.getRange(r, COL.VERSION).setHorizontalAlignment('center');

  if (folderUrl) {
    const safeFolder = folderName.replace(/"/g, "'");
    sheet.getRange(r, COL.FOLDER).setFormula(`=HYPERLINK("${folderUrl}","${safeFolder}")`);
  }
  sheet.getRange(r, COL.LINK).setFormula(`=HYPERLINK("${url}","Link")`);
}

/** Fills a full row with the separator red (blank text — visual group divider). */
function rebuildWriteSeparator_(sheet, r) {
  sheet.getRange(r, 1, 1, COL.OWNER)
       .setValues([Array(COL.OWNER).fill('')])
       .setBackground(SEPARATOR_RED);
}

/**
 * Applies the exact blue header design to row 1.
 * Column A retains its background colour; the Kiji logo floating image is unaffected.
 */
function rebuildFormatHeader_(sheet) {
  const headers = [
    '',                                  // A: logo (floating image — value stays blank)
    'Human-Readable\nName',              // B
    'Base File Name',                    // C
    'File Type',                         // D
    'Current\nVersion',                  // E
    'Current\nFolder',                   // F
    'Link',                              // G
    'For\nWho',                          // H
    'KAL\nName Conversion\nCheck',       // I
    'Destination\nDrive',                // J
    'Preferred\nKAL Template',           // K
    'Abstract',                          // L
    'Owner'                              // M
  ];

  sheet.getRange(1, 1, 1, headers.length)
       .setValues([headers])
       .setBackground(HEADER_BLUE)
       .setFontFamily('Georgia')
       .setFontColor('#ffffff')
       .setFontWeight('bold')
       .setFontSize(10)
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle')
       .setWrap(true);

  sheet.setRowHeight(1, 60);
}

/**
 * Inserts the Kiji logo into A1.  Called AFTER all row operations so that
 * insertRowsAfter calls in maintainGroupSpacing_ do not cause the floating
 * image to vanish until the next page refresh (a known Google Sheets quirk).
 */
function rebuildInsertLogo_(sheet) {
  // Remove any existing logo images in row 1 so we never stack duplicates.
  try {
    sheet.getImages().forEach(img => {
      if (img.getAnchorCell().getRow() === 1) img.remove();
    });
  } catch (_) {}

  // Priority: KAL_LOGO_PNG_BASE64 constant → Settings tab D2 (Drive link or direct URL)
  const blob = (KAL_LOGO_PNG_BASE64
      ? Utilities.newBlob(Utilities.base64Decode(KAL_LOGO_PNG_BASE64), 'image/png', 'kiji-logo.png')
      : null) || rebuildGetLogoBlobFromSettings_();
  if (!blob) {
    console.warn('Logo: no blob — set KAL_LOGO_PNG_BASE64 or add a PNG Drive link in Settings!D2');
    return;
  }
  const ct = blob.getContentType() || '';
  if (ct.indexOf('svg') !== -1) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Logo file is SVG — Google Sheets only supports PNG/JPG.\n' +
      'Upload a PNG version to Drive and update Settings!D2.',
      '⚠️ Logo', 10);
    console.warn('Logo skipped: SVG not supported by insertImage (contentType=' + ct + ')');
    return;
  }
  try {
    // Insert at 50×50 display size, centred in the 70×60 header cell.
    // High-res source (800×800) downsampled to 50×50 renders crisply.
    const img = sheet.insertImage(blob, 1, 1, 10, 5);
    if (img) { img.setWidth(50).setHeight(50); }
    Utilities.sleep(500); // allow Sheets to process the image before flushing
    SpreadsheetApp.flush();
    console.log('Logo inserted successfully (contentType=' + ct + ')');
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Logo insert failed: ' + e.message, '⚠️ Logo', 8);
    console.warn('Logo insertImage failed: ' + e.message);
  }
}

/** Reads logo blob from Settings!D2 (Drive share link or direct image URL). */
function rebuildGetLogoBlobFromSettings_() {
  try {
    const url = getUrlFromCell(SHEET.SETTINGS, 'D2');
    if (!url) { console.warn('Logo: no URL found in Settings!D2'); return null; }

    const driveMatch = url.match(/\/d\/([a-zA-Z0-9_-]{10,})/);
    if (driveMatch) {
      const fileId = driveMatch[1];
      // Fetch a scaled thumbnail (sz=s120 → ~120 px) to stay well under the
      // 2 MB / 1 M-pixel insertImage limit. Returns a JPEG regardless of source format.
      const thumbUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=s120';
      const resp = UrlFetchApp.fetch(thumbUrl, { muteHttpExceptions: true });
      if (resp.getResponseCode() === 200) {
        console.log('Logo: thumbnail fetched for Drive file ' + fileId);
        return resp.getBlob().setName('logo.jpg');
      }
      console.warn('Logo: thumbnail fetch returned ' + resp.getResponseCode() + ', trying full blob');
      return DriveApp.getFileById(fileId).getBlob();
    }

    // Fallback: treat as direct image URL
    console.log('Logo: fetching direct URL');
    return UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getBlob();
  } catch (e) {
    console.warn('Logo from Settings!D2 failed: ' + e.message);
    return null;
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// ── NEW FEATURES (1–10) ───────────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

// ── Feature 1: Promote from Versions dialog ───────────────────────────────────

/**
 * Called from the Show All Versions modal when the user clicks the Promote
 * button on the Latest Version row.  Mirrors the core logic of
 * promoteToFinalAndMove but operates by file ID (not active row).
 */
function promoteVersionFromDialog_(fileId) {
  if (!fileId) throw new Error('No file ID provided.');

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const archiveId = getIdFromUrl(getUrlFromCell(SHEET.SETTINGS, 'C2'));
  if (!archiveId) throw new Error('Archive folder URL missing in Settings!C2.');

  const srcFile   = DriveApp.getFileById(fileId);
  const name      = srcFile.getName();
  const baseName  = extractKALBaseName(name);
  const dateStr   = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd');
  const finalName = baseName + '_' + dateStr + '_vFINAL';

  // Resolve destination folder: Settings!A2 as default
  const destUrl = getUrlFromCell(SHEET.SETTINGS, 'A2');
  const destId  = getIdFromUrl(destUrl);
  if (!destId) throw new Error('No destination folder found in Settings!A2.');

  const newFile = srcFile.makeCopy(finalName, DriveApp.getFolderById(destId));
  srcFile.moveTo(DriveApp.getFolderById(archiveId));

  // Re-audit the matching registry row if possible
  try { updateAllInfo(); } catch (_) {}
  toast('"' + finalName + '" promoted to vFINAL. Original archived.', '🏁 Promoted', 6);

  // Email notification if owner is set on the matching row
  try {
    const sheet   = ss.getSheetByName(SHEET.REGISTRY);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= DATA_START) {
        const names = sheet.getRange(DATA_START, COL.FILENAME, lastRow - DATA_START + 1, 1).getValues();
        for (let i = 0; i < names.length; i++) {
          if (extractKALBaseName(names[i][0].toString()) === baseName) {
            const ownerEmail = sheet.getRange(DATA_START + i, COL.OWNER).getValue().toString().trim();
            if (ownerEmail && ownerEmail.includes('@')) {
              const promotionDate = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm');
              MailApp.sendEmail({
                to:      ownerEmail,
                subject: '[KAL] ' + finalName + ' promoted to vFINAL',
                body:    'Hello,\n\nThe following file has been promoted to vFINAL:\n\n' +
                         'File Name: ' + finalName + '\n' +
                         'Drive Link: ' + newFile.getUrl() + '\n' +
                         'Date: ' + promotionDate + '\n\n' +
                         'The previous version has been moved to the archive folder.\n\n' +
                         'KAL File Registry'
              });
            }
            break;
          }
        }
      }
    }
  } catch (mailErr) {
    console.warn('promoteVersionFromDialog_: email notification failed: ' + mailErr.message);
  }
}

// ── Feature 2: Duplicate Detection ───────────────────────────────────────────

/**
 * Scans Registry col C for rows sharing the same extractKALBaseName().
 * Highlights duplicates with orange (#f9cb9c) on col B–L.
 */
function detectDuplicates() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.REGISTRY);
  if (!sheet) { toast('Registry sheet not found.', '❌ Error', 5); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) { toast('No data rows found.', '⚠️ Warning', 4); return; }

  const n     = lastRow - DATA_START + 1;
  const names = sheet.getRange(DATA_START, COL.FILENAME, n, 1).getValues();

  // First pass: count occurrences of each base name
  const baseCount = {};
  names.forEach(row => {
    const fn = (row[0] || '').toString().trim();
    if (!fn) return;
    const base = extractKALBaseName(fn).toLowerCase();
    baseCount[base] = (baseCount[base] || 0) + 1;
  });

  // Second pass: highlight rows with base names that appear more than once
  const DUPE_COLOR = '#f9cb9c';
  const bgs = [];
  let dupeCount = 0;
  const dupeBases = new Set();

  names.forEach(row => {
    const fn = (row[0] || '').toString().trim();
    if (!fn) {
      bgs.push(null); // will not be applied to blank rows
      return;
    }
    const base = extractKALBaseName(fn).toLowerCase();
    if (baseCount[base] > 1) {
      bgs.push(DUPE_COLOR);
      dupeBases.add(base);
    } else {
      bgs.push(null);
    }
  });

  dupeCount = dupeBases.size;

  // Apply backgrounds
  const bgMatrix = bgs.map(bg => Array(LAST_COL - COL.DESC + 1).fill(bg));
  sheet.getRange(DATA_START, COL.DESC, n, LAST_COL - COL.DESC + 1).setBackgrounds(bgMatrix);

  toast('Found ' + dupeCount + ' duplicate base name(s). Highlighted in orange.', '🔁 Duplicates', 6);
}

// ── Feature 3: Stale File Warning ─────────────────────────────────────────────

/**
 * For each Registry row with a Drive file link, checks getLastUpdated().
 * Rows not modified in > threshold days get a yellow-orange background on col B
 * and " ⚠️ Stale" appended to col I.
 * Settings!F2 = stale threshold in days (default 180).
 */
function flagStaleFiles() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.REGISTRY);
  if (!sheet) { toast('Registry sheet not found.', '❌ Error', 5); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) { toast('No data rows found.', '⚠️ Warning', 4); return; }

  // Read stale threshold from Settings!F2
  let threshold = 180;
  try {
    const settingsSheet = ss.getSheetByName(SHEET.SETTINGS);
    if (settingsSheet) {
      const val = settingsSheet.getRange('F2').getValue();
      if (val && !isNaN(parseInt(val, 10))) threshold = parseInt(val, 10);
    }
  } catch (_) {}

  const n        = lastRow - DATA_START + 1;
  const names    = sheet.getRange(DATA_START, COL.FILENAME, n, 1).getValues();
  const linkRich = sheet.getRange(DATA_START, COL.LINK, n, 1).getRichTextValues();
  const kalChecks = sheet.getRange(DATA_START, COL.KAL_CHECK, n, 1).getValues();

  const now         = new Date();
  const STALE_COLOR = '#ffe599';
  let flagged = 0;

  for (let i = 0; i < n; i++) {
    const fn = (names[i][0] || '').toString().trim();
    if (!fn) continue;

    let fileUrl = null;
    try { fileUrl = linkRich[i][0] ? linkRich[i][0].getLinkUrl() : null; } catch (_) {}
    if (!fileUrl) continue;

    const fileId = getIdFromUrl(fileUrl);
    if (!fileId) continue;

    try {
      const file        = DriveApp.getFileById(fileId);
      const lastUpdated = file.getLastUpdated();
      const daysSince   = (now - lastUpdated) / (1000 * 60 * 60 * 24);

      if (daysSince > threshold) {
        const r = DATA_START + i;
        sheet.getRange(r, COL.DESC).setBackground(STALE_COLOR);
        const currentCheck = (kalChecks[i][0] || '').toString().trim();
        if (!currentCheck.includes('⚠️ Stale')) {
          sheet.getRange(r, COL.KAL_CHECK).setValue(currentCheck ? currentCheck + ' ⚠️ Stale' : '⚠️ Stale');
        }
        flagged++;
      }
    } catch (e) {
      console.warn('flagStaleFiles row ' + (DATA_START + i) + ': ' + e.message);
    }
  }

  toast('Flagged ' + flagged + ' stale file(s) (>' + threshold + ' days).', '⏰ Stale Files', 6);
}

// ── Feature 4: Summary Dashboard ─────────────────────────────────────────────

/**
 * Shows a modal dialog with a summary dashboard: totals by drive code,
 * entity, doc type, and KAL status.
 */
function generateSummaryDashboard() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName(SHEET.REGISTRY);
  if (!regSheet) { toast('Registry sheet not found.', '❌ Error', 5); return; }

  const lastRow = regSheet.getLastRow();
  if (lastRow < DATA_START) { toast('No data rows found.', '⚠️ Warning', 4); return; }

  const n         = lastRow - DATA_START + 1;
  const fileNames = regSheet.getRange(DATA_START, COL.FILENAME,  n, 1).getValues();
  const versions  = regSheet.getRange(DATA_START, COL.VERSION,   n, 1).getValues();
  const kalChecks = regSheet.getRange(DATA_START, COL.KAL_CHECK, n, 1).getValues();
  const linkVals  = regSheet.getRange(DATA_START, COL.LINK,      n, 1).getValues();

  const driveCounts  = {};
  const entityCounts = {};
  const docCounts    = {};
  const statusCounts = { ok: 0, error: 0, nonKal: 0, notFound: 0 };
  let vFinalCount = 0, inProgress = 0, totalFiles = 0;

  for (let i = 0; i < n; i++) {
    const fn = (fileNames[i][0] || '').toString().trim();
    if (!fn) continue;
    totalFiles++;

    const driveCode  = (fn.match(/^([A-Za-z]{2,4})-/) || [])[1] || '?';
    const parts      = fn.split(/[-_]/);
    const entityCode = parts.length >= 2 ? parts[1].toUpperCase() : '?';
    const docCode    = parts.length >= 3 ? parts[2].toUpperCase() : '?';

    driveCounts[driveCode]   = (driveCounts[driveCode]   || 0) + 1;
    entityCounts[entityCode] = (entityCounts[entityCode] || 0) + 1;
    docCounts[docCode]       = (docCounts[docCode]       || 0) + 1;

    const kalStatus = (kalChecks[i][0] || '').toString().trim();
    const linkVal   = (linkVals[i][0]  || '').toString().trim();
    if      (kalStatus === 'Non-KAL') statusCounts.nonKal++;
    else if (kalStatus === 'File not found' || linkVal === 'File not found') statusCounts.notFound++;
    else if (kalStatus && kalStatus !== 'OK') statusCounts.error++;
    else statusCounts.ok++;

    const ver = (versions[i][0] || '').toString().trim().toUpperCase();
    if (ver === 'FINAL') vFinalCount++;
    else inProgress++;
  }

  function tableRows_(obj) {
    return Object.entries(obj).sort((a, b) => b[1] - a[1])
      .map(([k, v]) => `<tr><td>${k}</td><td class="num">${v}</td></tr>`).join('');
  }

  const overviewRows = [
    ['Total Files',            totalFiles,              ''],
    ['✅ vFINAL',              vFinalCount,             'color:#1a7f37;font-weight:600'],
    ['🔄 In Progress',         inProgress,              ''],
    ['✔️ KAL Check OK',        statusCounts.ok,         'color:#1a7f37'],
    ['⚠️ Errors',             statusCounts.error,      'color:#c0392b'],
    ['📭 Missing from Drive',  statusCounts.notFound,   'color:#b45309'],
    ['📁 Non-KAL',             statusCounts.nonKal,     'color:#555'],
  ].map(([label, val, style]) =>
    `<tr><td style="${style}">${label}</td><td class="num" style="${style}">${val}</td></tr>`
  ).join('');

  const html = `
    <style>
      *{box-sizing:border-box}
      body{font-family:Arial,sans-serif;font-size:13px;padding:14px 16px;margin:0;color:#202124}
      h3{margin:0 0 2px;font-size:15px;color:#111184}
      .sub{margin:0 0 14px;font-size:11px;color:#888}
      .grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}
      .card{border:1px solid #e0e0e0;border-radius:6px;overflow:hidden}
      .card-title{background:#111184;color:#fff;font-size:11px;font-weight:700;
                  letter-spacing:.5px;text-transform:uppercase;padding:6px 10px}
      table{width:100%;border-collapse:collapse}
      td{padding:5px 10px;border-bottom:1px solid #f0f0f0;font-size:12px}
      tr:last-child td{border-bottom:none}
      tr:hover td{background:#f4f6ff}
      .num{text-align:right;font-variant-numeric:tabular-nums;font-weight:600;width:50px}
    </style>
    <h3>📊 Registry Summary</h3>
    <p class="sub">Generated ${new Date().toLocaleString()} &nbsp;·&nbsp; ${totalFiles} total file(s)</p>
    <div class="grid">
      <div class="card">
        <div class="card-title">Overview</div>
        <table>${overviewRows}</table>
      </div>
      <div class="card">
        <div class="card-title">By Drive Code</div>
        <table>${tableRows_(driveCounts)}</table>
      </div>
      <div class="card">
        <div class="card-title">By Entity</div>
        <table>${tableRows_(entityCounts)}</table>
      </div>
      <div class="card">
        <div class="card-title">By Doc Type</div>
        <table>${tableRows_(docCounts)}</table>
      </div>
    </div>
  `;

  const cardRows = Math.max(
    Object.keys(driveCounts).length,
    Object.keys(entityCounts).length,
    Object.keys(docCounts).length,
    7
  );
  const dialogHeight = Math.min(120 + cardRows * 28, 600);

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(720).setHeight(dialogHeight),
    '📊 KAL Registry Summary'
  );
}

// ── Feature 5: Missing Files Report ──────────────────────────────────────────

/**
 * Shows a modal dialog listing all Registry rows where col G = "File not found".
 */
function showMissingFilesReport() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.REGISTRY);
  if (!sheet) { toast('Registry sheet not found.', '❌ Error', 5); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) { toast('No data rows found.', '⚠️ Warning', 4); return; }

  const n         = lastRow - DATA_START + 1;
  const fileNames = sheet.getRange(DATA_START, COL.FILENAME, n, 1).getValues();
  const linkVals  = sheet.getRange(DATA_START, COL.LINK, n, 1).getValues();
  const kalChecks = sheet.getRange(DATA_START, COL.KAL_CHECK, n, 1).getValues();
  const rowNums   = sheet.getRange(DATA_START, COL.ROW_NUM, n, 1).getValues();

  const missing = [];
  for (let i = 0; i < n; i++) {
    const linkText = (linkVals[i][0] || '').toString().trim();
    if (linkText === 'File not found') {
      missing.push({
        sheetRow: DATA_START + i,
        rowNum:   (rowNums[i][0]   || '').toString() || String(DATA_START + i),
        fileName: (fileNames[i][0] || '').toString().trim(),
        kalCheck: (kalChecks[i][0] || '').toString().trim()
      });
    }
  }

  if (!missing.length) {
    SpreadsheetApp.getUi().showModalDialog(
      HtmlService.createHtmlOutput(
        '<p style="font-family:Arial;padding:16px;font-size:13px;color:#1a7f37">✅ No missing files found — all rows have Drive links.</p>'
      ).setWidth(380).setHeight(80),
      '📭 Missing Files Report'
    );
    return;
  }

  const rows = missing.map(m =>
    `<tr>
      <td style="text-align:center;font-weight:bold">${m.rowNum}</td>
      <td style="word-break:break-word">${m.fileName.replace(/</g,'&lt;')}</td>
      <td style="color:#888;font-size:11px">${m.kalCheck.replace(/</g,'&lt;')}</td>
    </tr>`
  ).join('');

  const html = HtmlService.createHtmlOutput(`
    <style>
      body{font-family:Arial,sans-serif;padding:14px 16px;margin:0;font-size:13px}
      h3{margin:0 0 8px;font-size:15px;color:#111184}
      .sub{margin:0 0 12px;font-size:11px;color:#888}
      table{width:100%;border-collapse:collapse}
      th,td{padding:7px 10px;text-align:left;border-bottom:1px solid #e8e8e8;vertical-align:middle}
      th{background:#111184;color:#fff;font-size:11px;font-weight:600}
      tr:last-child td{border-bottom:none}
      tr:hover td{background:#fff2cc}
    </style>
    <h3>📭 Missing Files Report</h3>
    <p class="sub">${missing.length} row(s) where Drive link = "File not found"</p>
    <table>
      <tr><th>#</th><th>File Name</th><th>KAL Check</th></tr>
      ${rows}
    </table>
    <p style="margin-top:10px;font-size:11px;color:#888">Use <em>Create Selected File</em> or add the file to Drive and re-audit to resolve.</p>
  `).setWidth(700).setHeight(Math.min(200 + missing.length * 36, 500));

  SpreadsheetApp.getUi().showModalDialog(html, '📭 Missing Files — ' + missing.length + ' found');
}

// ── Feature 7: Bulk Archive Selected Rows ────────────────────────────────────

/**
 * Archives Drive files for all selected rows to the archive folder (Settings!C2).
 */
function bulkArchiveSelected() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  if (sheet.getName() !== SHEET.REGISTRY) {
    toast('Switch to the Registry sheet first.', '⚠️ Warning', 4);
    return;
  }

  const archiveId = getIdFromUrl(getUrlFromCell(SHEET.SETTINGS, 'C2'));
  if (!archiveId) { toast('Archive folder URL missing in Settings!C2.', '🛑 Error', 5); return; }

  const range = sheet.getActiveRange();
  if (!range) { toast('Select one or more rows first.', '⚠️ Warning', 4); return; }

  const startRow = range.getRow();
  const numRows  = range.getNumRows();
  if (startRow < DATA_START) { toast('Select data rows (not the header).', '⚠️ Warning', 4); return; }

  const archiveFolder = DriveApp.getFolderById(archiveId);
  let archived = 0, skipped = 0;

  for (let i = 0; i < numRows; i++) {
    const r = startRow + i;
    const fn = sheet.getRange(r, COL.FILENAME).getValue().toString().trim();
    if (!fn) { skipped++; continue; }

    let fileUrl = null;
    try { fileUrl = sheet.getRange(r, COL.LINK).getRichTextValue().getLinkUrl(); } catch (_) {}
    if (!fileUrl) { skipped++; continue; }

    const fileId = getIdFromUrl(fileUrl);
    if (!fileId) { skipped++; continue; }

    try {
      toast('Archiving row ' + r + '…', '📦 Bulk Archive', 5);
      DriveApp.getFileById(fileId).moveTo(archiveFolder);
      archived++;
    } catch (e) {
      console.error('bulkArchiveSelected row ' + r + ': ' + e.message);
      skipped++;
    }
  }

  toast('Archived ' + archived + ' file(s).' + (skipped ? ' ' + skipped + ' row(s) skipped.' : ''), '📦 Bulk Archive', 6);
}

// ── Feature 9: Changelog Sheet helpers ───────────────────────────────────────

/**
 * Ensures a hidden "Changelog" sheet exists and returns it.
 */
function getOrCreateChangelogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Changelog');
  if (!sheet) {
    sheet = ss.insertSheet('Changelog');
    sheet.getRange(1, 1, 1, 2).setValues([['File ID', 'Note']]);
    sheet.hideSheet();
  }
  return sheet;
}

/**
 * Saves a version note to the Changelog sheet, keyed by Drive file ID.
 * Called from the Versions dialog (Feature 9).
 */
function saveVersionChangelog_(fileId, note) {
  if (!fileId) return;
  const sheet   = getOrCreateChangelogSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === fileId) {
        sheet.getRange(2 + i, 2).setValue(note || '');
        return;
      }
    }
  }
  // Not found: append a new row
  sheet.appendRow([fileId, note || '']);
}

// ── Feature 10: Filter/Sort Sidebar ──────────────────────────────────────────

/**
 * Shows a sidebar with dropdowns to filter the Registry by Drive Code,
 * Entity, Doc Type, and KAL Status.
 */
function showFilterSidebar() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('FilterSidebar')
      .setTitle('🔍 Filter Registry')
      .setWidth(300);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    toast('Could not load filter sidebar: ' + e.message, '❌ Error', 5);
  }
}

/**
 * Returns unique values for each filter dimension from the Registry.
 * Called by FilterSidebar.html on load.
 */
function getFilterOptions() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.REGISTRY);
  if (!sheet) return { driveCodes: [], entities: [], docTypes: [], statuses: [] };

  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) return { driveCodes: [], entities: [], docTypes: [], statuses: [] };

  const n         = lastRow - DATA_START + 1;
  const fileNames = sheet.getRange(DATA_START, COL.FILENAME, n, 1).getValues();
  const kalChecks = sheet.getRange(DATA_START, COL.KAL_CHECK, n, 1).getValues();

  const driveCodes = new Set();
  const entities   = new Set();
  const docTypes   = new Set();
  const statuses   = new Set();

  for (let i = 0; i < n; i++) {
    const fn = (fileNames[i][0] || '').toString().trim();
    if (!fn) continue;
    const parts = fn.split(/[-_]/);
    const m = fn.match(/^([A-Za-z]{2,4})-/);
    if (m) driveCodes.add(m[1].toUpperCase());
    if (parts.length >= 2) entities.add(parts[1].toUpperCase());
    if (parts.length >= 3) docTypes.add(parts[2].toUpperCase());
    const status = (kalChecks[i][0] || '').toString().trim();
    if (status) statuses.add(status.includes('⚠️ Stale') ? '⚠️ Stale' : status === 'OK' ? 'OK' : status === 'Non-KAL' ? 'Non-KAL' : 'Error');
  }

  return {
    driveCodes: [...driveCodes].sort(),
    entities:   [...entities].sort(),
    docTypes:   [...docTypes].sort(),
    statuses:   [...statuses].sort()
  };
}

/**
 * Hides rows not matching the selected filters.
 * Called by FilterSidebar.html on "Apply Filter".
 */
function applyRegistryFilter(filters) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.REGISTRY);
  if (!sheet) { toast('Registry sheet not found.', '❌ Error', 5); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) return;

  const n         = lastRow - DATA_START + 1;
  const fileNames = sheet.getRange(DATA_START, COL.FILENAME, n, 1).getValues();
  const kalChecks = sheet.getRange(DATA_START, COL.KAL_CHECK, n, 1).getValues();

  // First show all rows, then selectively hide
  sheet.showRows(DATA_START, n);

  const { driveCode, entity, docType, kalStatus } = filters;

  for (let i = 0; i < n; i++) {
    const fn = (fileNames[i][0] || '').toString().trim();
    if (!fn) continue; // leave blank/separator rows visible

    const parts   = fn.split(/[-_]/);
    const rowDrive  = ((fn.match(/^([A-Za-z]{2,4})-/) || [])[1] || '').toUpperCase();
    const rowEntity = (parts[1] || '').toUpperCase();
    const rowDoc    = (parts[2] || '').toUpperCase();
    const rowStatus = (kalChecks[i][0] || '').toString().trim();
    const rowStatusNorm = rowStatus.includes('⚠️ Stale') ? '⚠️ Stale' : rowStatus === 'OK' ? 'OK' : rowStatus === 'Non-KAL' ? 'Non-KAL' : rowStatus ? 'Error' : '';

    let hide = false;
    if (driveCode && driveCode !== 'ALL' && rowDrive  !== driveCode) hide = true;
    if (entity    && entity    !== 'ALL' && rowEntity !== entity)    hide = true;
    if (docType   && docType   !== 'ALL' && rowDoc    !== docType)   hide = true;
    if (kalStatus && kalStatus !== 'ALL' && rowStatusNorm !== kalStatus) hide = true;

    if (hide) sheet.hideRows(DATA_START + i);
  }

  toast('Filter applied.', '🔍 Filter', 3);
}

/**
 * Shows all rows in the Registry (clears any active filter).
 */
function clearRegistryFilter() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.REGISTRY);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow >= DATA_START) sheet.showRows(DATA_START, lastRow - DATA_START + 1);
  toast('Filter cleared — all rows visible.', '🔍 Filter', 3);
}
