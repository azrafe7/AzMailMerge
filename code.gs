const MERGE_SETTINGS_SHEET = 'merge_settings';
const TEMPLATE_SETTINGS_TEST_COL = 'TEST';

function NO_OP() {}

// create menu
function onOpen() {
  TOASTER.log(TOASTER.WARN, "Loading menu & scripts. Please wait...", 10);

  const { rowMap, templateSettings } = loadTemplateSettings(true);

  let menu = G.ui.createMenu('AzMM')
    .addItem("Merge", "AzMM.merge");

  menu.addSeparator()
    .addItem("Show Template Settings", 'AzMM.showTemplateSettings');

  menu.addToUi();
  
  TOASTER.log(TOASTER.INFO, "Menu & scripts loaded!", 5);
}

function uiAvailable() {
  try {
    SpreadsheetApp.getUi();
    return true;
  } catch (error) {
    return false;
  }
}

function getSheetColumnsMap(sheetOrData) {
  let columnsMap = new Map();
  const headers = Array.isArray(sheetOrData) ? sheetOrData[0] : sheetOrData.getRange("1:1").getValues()[0];
  headers.forEach((name) => columnsMap.set(name, headers.indexOf(name)));
  return columnsMap;
}

/**
 * Converts row/col numbers to A1 notation
 * @param {number} row - Row number (1-indexed)
 * @param {number} col - Column number (1-indexed)
 * @returns {Array<string>} A1 notation (e.g., ['A', '1'], ['B', '5'])
 */
function toA1Notation(row, col) {
  let colStr = '';
  let tempCol = col;
  
  while (tempCol > 0) {
    tempCol--;
    colStr = String.fromCharCode(65 + (tempCol % 26)) + colStr;
    tempCol = Math.floor(tempCol / 26);
  }
  
  return [colStr, row];
}

/**
 * Parses A1 notation to row/col numbers
 * @param {string} a1 - A1 notation (e.g., 'A1', 'B5', 'AA10')
 * @returns {Object} Object with row and col properties
 */
function parseA1Notation(a1) {
  const match = a1.match(/^\$?([A-Z]+)?\$?(\d+)?$/);
  if (!match) throw new Error(`Invalid A1 notation: ${a1}`);
  
  const colStr = match[1] || null;
  const row = parseInt(match[2]) || null;
  
  let col = null;
  for (let i = 0; colStr != null && i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }
  
  return {row, col};
}


const MERGE_DATA_FIRST_COL = 'F';

/**
 * Reads data and splits at MERGE_DATA_FIRST_COL
 * Returns { metaColumnsMap, metaData, mergeColumnsMap, mergeData, mergeDisplayData }
 */
function getMergeData() {
  const sheet = G.ss.getSheetByName(MERGE_SETTINGS_SHEET);
  const dataRange = sheet.getDataRange();
  let allData = dataRange.getValues();
  let allDisplayData = dataRange.getDisplayValues();
  
  const MERGE_DATA_FIRST_COL_IDX = parseA1Notation(MERGE_DATA_FIRST_COL).col - 1;
  let metaData = [];
  let mergeData = [];
  let mergeDisplayData = [];

  for (let rowIndex = 0; rowIndex < allData.length; rowIndex++) {
    metaData.push(allData[rowIndex].slice(0, MERGE_DATA_FIRST_COL_IDX));
    mergeData.push(allData[rowIndex].slice(MERGE_DATA_FIRST_COL_IDX));
    mergeDisplayData.push(allDisplayData[rowIndex].slice(MERGE_DATA_FIRST_COL_IDX));
  }

  const metaColumnsMap = getSheetColumnsMap(metaData);
  metaData.shift();

  const mergeColumnsMap = getSheetColumnsMap(mergeData);
  mergeData.shift();
  mergeDisplayData.shift();

  return { metaColumnsMap, metaData, mergeColumnsMap, mergeData, mergeDisplayData };
}

function getFileNameFromId(id) {
  const file = DriveApp.getFileById(id);
  const name = file.getName();
  return name;
}

function transpose(matrix) {
  return matrix[0].map((_, colIndex) => matrix.map(row => row[colIndex]));
}

function toBool(value) {
  const strValue = String(value);
  return strValue === 'true' || strValue === 'yes' || strValue === '1' ? true : false;
}

const REPLACE_STATUS = Object.freeze({
  OK: "OK",
  NOT_FOUND: 'NOT_FOUND',
  KEPT_RAW: 'RAW',
  UNKNOWN: "UNKNOWN"
});

/**
 * Parses `string` into tokens, marking {{placeholders}} as special ones (and extracts offset and key).
 * Returns an array of tokens like [{raw:"text ", start:0, key:null}, {raw:"{{placeholder}}", start:5, key:"placeholder"}].
 */
function tokenize(string) {
  const tokens = [];

  let currPos = 0;
  const matches = Array.from(string.matchAll(/{{([^{}}]+|{[^}]+})}}/g)); // allow placeholders to be in object notation (wrapped in {})
  const numMatches = matches.length;
  const stringLength = string.length;

  for (let i=0; i<numMatches; i++) {
    const currMatch = matches[i];
    const nextMatch = i < numMatches-1 ? matches[i + 1] : null;
    const key = currMatch[0].replace(/^{{|}}$/g, "");
    if (currPos < currMatch.index) {
      tokens.push({
        raw: string.slice(currPos, currMatch.index),
        start: currPos,
        key: null
      });
    }
    tokens.push({
      raw: currMatch[0],
      start: currMatch.index,
      key: key
    });
    currPos = currMatch.index + currMatch[0].length;

    if (nextMatch && currMatch.index + currMatch[0].length > nextMatch.index) {
      throw "Next match within prev match!?";
    }
  }

  if (currPos < stringLength - 1) {
    tokens.push({
      raw: string.slice(currPos),
      start: currPos,
      key: null
    });
  }

  return tokens;
}

function interpolateTemplateString(templateString, dataRow, columnsMap, cachedMap) {
  if (cachedMap == null) cachedMap = new Map();
  if (columnsMap == null) columnsMap = new Map();
  let tokens = tokenize(templateString);
  let interpolated = tokens.map((token) => {
    const key = token.key;
    token = Object.assign(token, {
      replacedWith: "",
      status: REPLACE_STATUS.UNKNOWN
    });
    
    if (key == null) {
      token.replacedWith = token.raw;
      token.status = REPLACE_STATUS.KEPT_RAW;
      return token.replacedWith;
    }

    const cachedMapValue = cachedMap.get(key);
    const inCachedMap = cachedMapValue != null;   
    let replacedWith = cachedMapValue != null ? cachedMapValue : "";
    const dataRowIndex = columnsMap.get(key);
    const inDataRow = Array.isArray(dataRow) && dataRowIndex != null && dataRowIndex >= 0; 
    if (inDataRow) {
      replacedWith = dataRow[dataRowIndex];
    }
    token.replacedWith = replacedWith;
    token.status = inCachedMap || inDataRow ? REPLACE_STATUS.OK : REPLACE_STATUS.NOT_FOUND;
    return replacedWith;
  });

  return { interpolated: interpolated.join(""), tokens };
}

function getCachedMap() {
  const mapping = {
    NOW: (() => new Date().toLocaleString())()
  };
  const map = new Map(Object.entries(mapping));
  return map;
}

function fillTemplateSettingsTestCol() {
  let sheet = G.ss.getSheetByName(TEMPLATE_SETTINGS_SHEET);
  const { rowMap, templateSettings } = loadTemplateSettings(true);
  let columnsMap = getSheetColumnsMap(sheet);
  const testColIdx = columnsMap.get(TEMPLATE_SETTINGS_TEST_COL) + 1;
  const testRange = sheet.getRange(2, testColIdx, Object.keys(templateSettings, 2).length);
  let testValues = testRange.getValues();
  Logger.log(testValues);

  let { metaColumnsMap, metaData, mergeColumnsMap, mergeData, mergeDisplayData } = getMergeData();
  const dataRow = mergeData.length > 0 ? mergeData[0] : null;
  const cachedMap = getCachedMap();

  const filledTemplateSettings = {};
  for (let [k, v] of Object.entries(templateSettings)) {
    let testValueRow = testValues[rowMap.get(k)];
    if (k.endsWith(ID_MATCHER) && String(v) != '') testValueRow[0] = getFileNameFromId(v);
    else if (k.endsWith(FORMAT_MATCHER) && String(v) != '') testValueRow[0] = interpolateTemplateString(v, dataRow, mergeColumnsMap, cachedMap).interpolated;
    else testValueRow[0] = v;
    filledTemplateSettings[k] = testValueRow[0];
  }

  testRange.setValues(testValues);
  Logger.log((testValues));
  return filledTemplateSettings;
}

function merge() {
  const { rowMap, templateSettings } = loadTemplateSettings(true);
  const filledTemplateSettings = fillTemplateSettingsTestCol();
  const cachedMap = getCachedMap();
  let { metaColumnsMap, metaData, mergeColumnsMap, mergeData, mergeDisplayData } = getMergeData();

  if (templateSettings[TSETTING_DOC_TEMPLATE_ID] === '') {
    const message = `Missing ${TSETTING_DOC_FOLDER_ID}!`;
    G.ui.alert(TOASTER.ERROR, message, G.ui.ButtonSet.OK);
    Logger.log(TOASTER.ERROR + " " + message);
    return;
  }

  if (mergeData.length == 0) {
    const message = `No data to merge. Check '${MERGE_SETTINGS_SHEET}' sheet!`;
    G.ui.alert(TOASTER.ERROR, message, G.ui.ButtonSet.OK);
    Logger.log(TOASTER.ERROR + " " + message);
    return;
  }

  {
    const file = DriveApp.getFileById(templateSettings[TSETTING_DOC_TEMPLATE_ID]);
    const originalFolderId = file.getParents().next().getId();
    const templateFolderId = templateSettings[TSETTING_DOC_FOLDER_ID] !== '' ? originalFolderId : templateSettings[TSETTING_DOC_FOLDER_ID];
    const pdfFolderId = templateSettings[TSETTING_PDF_FOLDER_ID] !== '' ? originalFolderId : templateSettings[TSETTING_PDF_FOLDER_ID];
    const templateFolder = DriveApp.getFolderById(templateFolderId);
    const pdfFolder = DriveApp.getFolderById(pdfFolderId);
    const copy = file.makeCopy("copy 00", templateFolder);

    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();

    for (let rowIndex=0; rowIndex < 1; rowIndex++) {
      const dataRow = mergeData[rowIndex];
      Logger.log(`Row ${rowIndex}`);

      const matches = findPlaceholders(body);
      // iterate matches backwards
      for (let i = matches.length - 1; i >= 0; i--) {
        const r = matches[i];
        const textElement = r.getElement().asText();
        const text = textElement.getText();
        const start = r.getStartOffset();
        const endInclusive = r.getEndOffsetInclusive();
        const matched = text.slice(start, endInclusive + 1);

        const textToReplace = matched;
        const replacement = interpolateTemplateString(textToReplace, dataRow, mergeColumnsMap, cachedMap);

        // save formatting attributes (into a clone)
        const attrs = Object.assign({}, textElement.getAttributes(start));
        Logger.log(JSON.stringify([matched, start, endInclusive]));
        Logger.log(JSON.stringify(["BEFORE:", textElement.getAttributes(start)]));

        // delete text and insert replacement
        textElement.deleteText(start, endInclusive);
        textElement.insertText(start, replacement.interpolated);
        const newEndInclusive = start + replacement.interpolated.length - 1;
        
        // restore attributes (only non null values, to prevent not inheriting attributes from parent, and messing things up!)
        for (const [attr, value] of Object.entries(attrs)) {
          if (value != null) {
            textElement.setAttributes(start, newEndInclusive, {
              [attr]: value
            });
          }
        }

        Logger.log(JSON.stringify(["AFTER:", textElement.getAttributes(start)]));
      }
    }

    doc.saveAndClose();
  } /*catch (error) {
    G.ui.alert(TOASTER.ERROR, error, G.ui.ButtonSet.OK);
    throw error;
  }*/
}

function findPlaceholders(body) {
  return Array.from(findAllText(body, "{{([^{}}]+|{[^}]+})}}"));
}

/**
 * Generator that finds all the entries when searching.
 * @param {DocumentApp.Body} body Body of the file.
 * @param {string} text Text to find.
 * @returns {Iterator.<DocumentApp.RangeElement>}
 */
function* findAllText(body, text) {
  let entry = body.findText(text);
  while(entry != null) {
    yield entry;
    entry = body.findText(text, entry);
  }
}

function getChart(sheetName) {
  sheetName = "Chart 1";
  const sheet = G.ss.getSheetByName(sheetName);
  const charts = sheet.getCharts();
  if (charts.length > 0) {
    return charts[0].getAs
  } else {
    return null;
  }
}

// exported functions
var AzMM = {
  onOpen: onOpen,
  showTemplateSettings: showTemplateSettings,
  merge: merge
}
