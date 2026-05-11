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
 *  L(12) Abstract               [rightmost / user-entered]
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💠 KAL File System')
    .addItem('🎯 Audit & Sync Selected File', 'updateSelectedInfo')
    .addItem('🔄 Audit & Sync All Files', 'updateAllInfo')
    .addSeparator()
    .addItem('🏗️ Create Selected File', 'createSelectedFile')
    .addItem('🗑️ Remove Selected Version', 'removeSelectedFile')
    .addItem('☢️ Remove All Versions', 'removeAllVersions')
    .addSeparator()
    .addItem('📖 Show The User Guide', 'showUserGuide')
    .addToUi();
  updateTemplateDropdown();
}

/** HELPER: Extracts hidden URLs from cells (Rich Text) */
function getUrlFromCell(sheetName, cellAddress) {
  try {
    const cell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(cellAddress);
    const richText = cell.getRichTextValue();
    let url = richText.getLinkUrl();
    if (!url) url = cell.getValue().toString().trim();
    return url;
  } catch (e) { return null; }
}

function getIdFromUrl(url) {
  if (!url) return null;
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

/** 1. DROPDOWN FETCHER */
function getTemplateList() {
  const folderUrl = getUrlFromCell("Settings", "B2");
  const folderId = getIdFromUrl(folderUrl);
  if (!folderId) return [];
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const names = [];
    while (files.hasNext()) { names.push(files.next().getName()); }
    return names;
  } catch (e) { return []; }
}

/** 2. CORE AUDIT ENGINE (With CamelCase Smart Splitter) */
function processAuditForRow(sheet, r, driveUrlLookup, validEntities, validDocs, templateList) {
  let baseName = sheet.getRange(r, 3).getValue().toString().trim();
  let dropdownCell = sheet.getRange(r, 11); // Col K: Preferred KAL Template

  if (!baseName) {
    sheet.getRange(r, 1, 1, 2).clearContent();
    sheet.getRange(r, 4, 1, 8).clearContent();    // cols D–K (File Type … Preferred KAL Template)
    sheet.getRange(r, 2, 1, 10).setBackground(null); // cols B–K
    dropdownCell.clearContent().clearDataValidations();
    return;
  }

  // --- SMART DESCRIPTION EXTRACTOR ---
  let nameParts = baseName.split(/[-_]/);

  if (nameParts.length >= 4) {
    let descriptionRaw = nameParts.slice(3).join("");
    let cleanDescription = descriptionRaw
      .replace(/([a-z])([A-Z])/g, '$1 $2')       // "FileRegister" -> "File Register"
      .replace(/([A-Z])([A-Z][a-z])/g, '$1 $2'); // "PPFile" -> "PP File"
    sheet.getRange(r, 2).setValue(cleanDescription.trim());
  }

  let driveCodeRaw   = (nameParts[0] || "").trim();
  let entityCodeRaw  = (nameParts[1] || "").trim();
  let docTypeCodeRaw = (nameParts[2] || "").trim();

  let diagnostics = [];
  if (!driveUrlLookup[driveCodeRaw.toUpperCase()]) {
    diagnostics.push("Invalid Drive Code");
  } else if (driveCodeRaw !== driveCodeRaw.toUpperCase()) {
    diagnostics.push("Drive Code must be UPPERCASE");
  }

  if (validEntities.indexOf(entityCodeRaw.toUpperCase()) === -1) {
    diagnostics.push("Unregistered Entity");
  } else if (entityCodeRaw !== entityCodeRaw.toUpperCase()) {
    diagnostics.push("Entity Code must be UPPERCASE");
  }

  if (validDocs.indexOf(docTypeCodeRaw.toUpperCase()) === -1) {
    diagnostics.push("Invalid DocType");
  } else if (docTypeCodeRaw !== docTypeCodeRaw.toUpperCase()) {
    diagnostics.push("DocType must be UPPERCASE");
  }

  let status = diagnostics.length > 0 ? diagnostics.join(" | ") : "OK";
  sheet.getRange(r, 9).setValue(status); // Col I: KAL Name Conversion Check

  // Col J (10): Destination Drive
  if (driveUrlLookup[driveCodeRaw.toUpperCase()]) {
    sheet.getRange(r, 10).setFormula('=HYPERLINK("' + driveUrlLookup[driveCodeRaw.toUpperCase()] + '", "' + driveCodeRaw.toUpperCase() + ' Drive")');
  } else {
    sheet.getRange(r, 10).clearContent();
  }

  let info = GET_SMART_DETAILS(baseName);
  sheet.getRange(r, 4, 1, 5).clearContent();

  if (info.fileLink !== "Not Found") {
    sheet.getRange(r, 4).setValue(info.type);
    sheet.getRange(r, 5).setValue(info.version);
    sheet.getRange(r, 6).setFormula('=HYPERLINK("' + info.folderLink + '", "' + info.folderName + '")');
    sheet.getRange(r, 7).setFormula('=HYPERLINK("' + info.fileLink + '", "Link")');
    sheet.getRange(r, 8).setValue(entityCodeRaw.toUpperCase());
    dropdownCell.clearContent().clearDataValidations();
  } else {
    sheet.getRange(r, 7).setValue("File not found");
    if (templateList && templateList.length > 0) {
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(templateList).setAllowInvalid(false).build();
      dropdownCell.setDataValidation(rule);
    }
  }

  let color = status !== "OK" ? "#f4cccc" : (info.version === "FINAL" ? "#d9ead3" : null);
  sheet.getRange(r, 2, 1, 11).setBackground(color); // cols B–L (includes Abstract at col 12)
}

/** 3. SYNC LOGIC */
function getLevelsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const levelsSheet = ss.getSheetByName("Levels");
  const lastRow = levelsSheet.getLastRow();
  const lvRange = levelsSheet.getRange(1, 1, lastRow, 7);
  const lvValues = lvRange.getValues();
  const lvRichText = lvRange.getRichTextValues();
  const driveUrlLookup = {}; const validEntities = []; const validDocs = [];
  for (let j = 1; j < lvValues.length; j++) {
    let driveCode = String(lvValues[j][0]).toUpperCase().trim();
    let hiddenUrl = lvRichText[j][1] ? lvRichText[j][1].getLinkUrl() : null;
    if (driveCode && hiddenUrl) driveUrlLookup[driveCode] = hiddenUrl;
    let entity = String(lvValues[j][3]).toUpperCase().trim();
    if (entity) validEntities.push(entity);
    let docType = String(lvValues[j][6]).toUpperCase().trim();
    if (docType) validDocs.push(docType);
  }
  return { driveUrlLookup, validEntities, validDocs };
}

function updateAllInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = getLevelsData();
  const templateList = getTemplateList();
  for (let r = 2; r <= sheet.getLastRow(); r++) {
    processAuditForRow(sheet, r, data.driveUrlLookup, data.validEntities, data.validDocs, templateList);
    if (sheet.getRange(r, 3).getValue()) sheet.getRange(r, 1).setValue(r - 1);
  }
}

function updateSelectedInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const r = sheet.getActiveRange().getRow();
  if (r < 2) return;
  const data = getLevelsData();
  const templateList = getTemplateList();
  processAuditForRow(sheet, r, data.driveUrlLookup, data.validEntities, data.validDocs, templateList);
}

/** 4. DRIVE SEARCH ENGINE */
function GET_SMART_DETAILS(baseName) {
  const files = DriveApp.searchFiles("title contains '" + baseName + "'");
  let latestVerNum = -1;
  let details = { type: "", version: "Not Found", folderName: "Not Found", folderLink: "", fileLink: "Not Found" };
  while (files.hasNext()) {
    let file = files.next();
    let name = file.getName();
    let regex = new RegExp(baseName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ".*[vV](\\d+|FINAL)", "i");
    let match = name.match(regex);
    if (match) {
      let vSuffix = match[1].toUpperCase();
      let vNum = (vSuffix === "FINAL") ? 9999 : parseInt(vSuffix, 10);
      if (vNum > latestVerNum) {
        latestVerNum = vNum;
        details.version = vSuffix;
        details.fileLink = file.getUrl();
        details.type = formatMimeType(file.getMimeType());
        let parents = file.getParents();
        if (parents.hasNext()) {
          let p = parents.next();
          details.folderName = p.getName();
          details.folderLink = p.getUrl();
        }
      }
    }
  }
  return details;
}

/** 5. FILE MANAGEMENT */
function createSelectedFile() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const r = sheet.getActiveRange().getRow();
  if (r < 2) return;

  let baseName     = sheet.getRange(r, 3).getValue().toString().trim();
  let templateName = sheet.getRange(r, 11).getValue().toString().trim(); // Col K: Preferred KAL Template

  const data = getLevelsData();
  const templateList = getTemplateList();
  processAuditForRow(sheet, r, data.driveUrlLookup, data.validEntities, data.validDocs, templateList);
  let nameStatus = sheet.getRange(r, 9).getValue().toString().trim();

  if (!baseName || nameStatus !== "OK" || !templateName) {
    SpreadsheetApp.getUi().alert("🛑 Check name status and template selection.");
    return;
  }

  if (GET_SMART_DETAILS(baseName).fileLink !== "Not Found") {
    SpreadsheetApp.getUi().alert("🛑 File already exists.");
    return;
  }

  try {
    const destUrl = getUrlFromCell("Settings", "A2");
    const dateString = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyyMMdd");
    const finalFileName = baseName + "_" + dateString + "_v1";

    const tempFolder = DriveApp.getFolderById(getIdFromUrl(getUrlFromCell("Settings", "B2")));
    const templateFiles = tempFolder.getFilesByName(templateName);
    if (templateFiles.hasNext()) {
      templateFiles.next().makeCopy(finalFileName, DriveApp.getFolderById(getIdFromUrl(destUrl)));
      updateSelectedInfo();
      SpreadsheetApp.getUi().alert("✅ Success!");
    }
  } catch (e) { SpreadsheetApp.getUi().alert(e.message); }
}

function removeSelectedFile() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const r = sheet.getActiveRange().getRow();
  const fileUrl = sheet.getRange(r, 7).getRichTextValue().getLinkUrl();
  if (!fileUrl) return;

  const response = SpreadsheetApp.getUi().alert("⚠️ Warning", "Move THIS version to Trash?", SpreadsheetApp.getUi().ButtonSet.YES_NO);
  if (response == SpreadsheetApp.getUi().Button.YES) {
    DriveApp.getFileById(getIdFromUrl(fileUrl)).setTrashed(true);
    updateSelectedInfo();
  }
}

function removeAllVersions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const r = sheet.getActiveRange().getRow();
  const baseName = sheet.getRange(r, 3).getValue().toString().trim();
  if (!baseName) return;

  const response = SpreadsheetApp.getUi().alert("☢️ NUCLEAR WARNING", "Trash EVERY version of: " + baseName, SpreadsheetApp.getUi().ButtonSet.YES_NO);
  if (response == SpreadsheetApp.getUi().Button.YES) {
    const files = DriveApp.searchFiles("title contains '" + baseName + "'");
    let count = 0;
    while (files.hasNext()) {
      files.next().setTrashed(true);
      count++;
    }
    updateSelectedInfo();
    SpreadsheetApp.getUi().alert("Trashed " + count + " file(s).");
  }
}

function formatMimeType(mime) {
  const types = {
    "application/vnd.google-apps.document":     "G-Doc",
    "application/vnd.google-apps.spreadsheet":  "G-Sheet",
    "application/vnd.google-apps.presentation": "G-Slide",
    "application/pdf": "PDF"
  };
  return types[mime] || "File";
}

function showUserGuide() {
  const template = HtmlService.createTemplateFromFile('Sidebar');
  template.logoBase64 = KAL_LOGO_BASE64;
  const html = template.evaluate()
    .setTitle('KAL File System')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Fetches all registered codes and smart description examples for the Academy sidebar.
 */
function getAcademyCodes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const levelsSheet = ss.getSheetByName("Levels");
  const lastRow = levelsSheet.getLastRow();

  if (lastRow < 3) return { drives: [], entities: [], docTypes: [], examples: [] };

  // Adjusted range to include Column I (index 8) and Column J (index 9)
  const values = levelsSheet.getRange(3, 1, lastRow - 2, 10).getValues();

  let driveList   = [];
  let entityList  = [];
  let docTypeList = [];
  let exampleList = [];

  values.forEach(row => {
    if (row[0]) driveList.push({ code: row[0], name: row[2] });
    if (row[3]) entityList.push({ code: row[3], name: row[4] });
    if (row[6]) docTypeList.push({ code: row[6], name: row[7] });
    // Captures CamelCase examples from Column I and J
    if (row[8]) exampleList.push({ input: row[8], output: row[9] });
  });

  return {
    drives:   driveList,
    entities: entityList,
    docTypes: docTypeList,
    examples: exampleList
  };
}
