const TEMPLATE_SETTINGS_SHEET = 'template_settings';

const ID_MATCHER = ' ID';
const FORMAT_MATCHER = ' FORMAT';

const TSETTINGS_NAME_COL = "SETTING";
const TSETTINGS_VALUE_COL = "VALUE";

const TSETTING_DOC_TEMPLATE_ID = "DOC TEMPLATE ID";
const TSETTING_DOC_FOLDER_ID = "DOC FOLDER ID";
const TSETTING_DOC_NAME_FORMAT = "DOC NAME FORMAT";
const TSETTING_KEEP_DOC = "KEEP DOC";
const TSETTING_CREATE_PDF = "CREATE PDF";
const TSETTING_PDF_FOLDER_ID = "DOC FOLDER ID";
const TSETTING_PDF_NAME_FORMAT = "PDF NAME FORMAT";

const DEFAULT_TSETTINGS_PAIRS = [
  [TSETTING_DOC_TEMPLATE_ID, ''],
  [TSETTING_DOC_FOLDER_ID, ''],
  [TSETTING_DOC_NAME_FORMAT, ''],
  [TSETTING_KEEP_DOC, 'yes'],
  [TSETTING_CREATE_PDF, 'no'],
  [TSETTING_PDF_FOLDER_ID, ''],
  [TSETTING_PDF_NAME_FORMAT, ''],
];

const DEFAULT_TSETTINGS = {};
DEFAULT_TSETTINGS_PAIRS.forEach(([n, v]) => DEFAULT_TSETTINGS[n] = v);

function getTemplateSettingsMessage() {
  const { _, templateSettings } = loadTemplateSettings(false);
  let message = "[templateSettings]\r\n\r\n";
  
  for (let [k,v] of Object.entries(templateSettings)) {
    if (k == "") continue;
    message += `  ${k}: ${v}\r\n`;
  }
  message = message.replace(/ /g, '\xa0'); // non breaking spaces

  return message;
}

function showTemplateSettings() {
  const message = getTemplateSettingsMessage();

  G.ui.alert(TOASTER.INFO, message, G.ui.ButtonSet.OK);
}

/**
 * Returns { rowMap, templateSettings }
 */
function loadTemplateSettings(quiet=false) {
  let sheet = SpreadsheetApp.getActive().getSheetByName(TEMPLATE_SETTINGS_SHEET);

  let TEMPLATE_SETTINGS = {};
  TEMPLATE_SETTINGS = Object.assign({}, DEFAULT_TSETTINGS, TEMPLATE_SETTINGS);
  let rowMap = null;

  if (!sheet) {
    if (!quiet) G.ui.alert(TOASTER.WARN, `Sheet '${TEMPLATE_SETTINGS_SHEET}' not found.\r\nUsing defaults.`, G.ui.ButtonSet.OK);
  } else {

    let data = sheet.getDataRange().getValues();
    const columnsMap = getSheetColumnsMap(data);
    const headers = data.shift();
    rowMap = getSheetColumnsMap(transpose(data));
    let idx = 0;
    for (let row of data) {
      const setting = row[columnsMap.get(TSETTINGS_NAME_COL)];
      const value = row[columnsMap.get(TSETTINGS_VALUE_COL)];
      if (String(value) != '') TEMPLATE_SETTINGS[setting] = value;
    }
  }

  Logger.log(["TEMPLATE_SETTINGS:", TEMPLATE_SETTINGS]);
  return { rowMap, templateSettings: TEMPLATE_SETTINGS };
}

