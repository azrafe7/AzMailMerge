const MERGE_SETTINGS_SHEET = 'merge_settings';
const TEMPLATE_SETTINGS_TEST_COL = 'TEST';

// create menu
function onOpen() {
  TOASTER.log(TOASTER.WARN, "Loading menu & scripts. Please wait...", 10);

  const templateSettings = loadTemplateSettings(true);

  let menu = G.ui.createMenu('AzMM')
    .addItem("Merge", "AzMM.merge");

  menu.addSeparator()
    .addItem("Show Template Settings", 'AzMM.showTemplateSettings');

  menu.addToUi();
  
  TOASTER.log(TOASTER.INFO, "Menu & scripts loaded!", 5);
}

function NO_OP() {}

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
  const matches = Array.from(string.matchAll(/{{[^{}]+}}/g));
  const numMatches = matches.length;
  const stringLength = string.length;

  for (let i=0; i<numMatches; i++) {
    const currMatch = matches[i];
    const nextMatch = i < numMatches-1 ? matches[i + 1] : null;
    const key = currMatch[0].replace(/{{|}}/g, "");
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
    replacedWith = inDataRow ? dataRow[dataRowIndex] : "";
    token.replacedWith = replacedWith;
    token.status = inCachedMap || inDataRow ? REPLACE_STATUS.OK : REPLACE_STATUS.NOT_FOUND;
    return replacedWith;
  });

  return { interpolated: interpolated.join(""), tokens };
}

function fillTemplateSettingsTestCol() {
  let sheet = G.ss.getSheetByName(TEMPLATE_SETTINGS_SHEET);
  const { rowMap, templateSettings } = loadTemplateSettings(true);
  let columnsMap = getSheetColumnsMap(sheet);
  const testColIdx = columnsMap.get(TEMPLATE_SETTINGS_TEST_COL) + 1;
  const testRange = sheet.getRange(2, testColIdx, Object.keys(templateSettings, 2).length);
  let testValues = testRange.getValues();
  Logger.log(testValues);

  for (let [k, v] of Object.entries(templateSettings)) {
    if (k.endsWith(ID_MATCHER) && String(v) != '') testValues[rowMap.get(k)][0] = getFileNameFromId(v);
    else if (k.endsWith(FORMAT_MATCHER) && String(v) != '') testValues[rowMap.get(k)][0] = v;
    else testValues[rowMap.get(k)][0] = v;
  }

  testRange.setValues(testValues);
  Logger.log((testValues));
}

// exported functions
var AzMM = {
  onOpen: onOpen,
  showTemplateSettings: showTemplateSettings,
  merge: NO_OP
}
