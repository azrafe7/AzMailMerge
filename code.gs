const MERGE_SETTINGS_SHEET = 'merge_settings';
const LOG_SHEET = 'log';
const TEMPLATE_SETTINGS_TEST_COL = 'TEST';

const DOC_PLACEHOLDERS_PATTERN = "{{([^{}}]+|{[^}]+})}}";

const MERGE_DATA_FIRST_COL = 'F';

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

function fillTemplateSettingsTestCol() {
  let sheet = G.ss.getSheetByName(TEMPLATE_SETTINGS_SHEET);
  const { rowMap, templateSettings } = loadTemplateSettings(true);
  let columnsMap = getSheetColumnsMap(sheet);
  const testColIdx = columnsMap.get(TEMPLATE_SETTINGS_TEST_COL) + 1;
  const testRange = sheet.getRange(2, testColIdx, Object.keys(templateSettings, 2).length);
  let testValues = testRange.getValues();
  Logger.log(testValues);

  let { metaColumnsMap, metaData, mergeColumnsMap, mergeData, mergeDisplayData } = getMergeData();
  const dataRow = mergeDisplayData.length > 0 ? mergeDisplayData[0] : null;

  const textRenderer = new TextRenderer();

  const context = {
    dataRow,
    columnsMap: mergeColumnsMap,
  };

  const interpolator = new Interpolator({
    context,
    functions: getFunctionsMap(),
    commands: getCommandsMap(),
  });

  const filledTemplateSettings = {};
  for (let [k, v] of Object.entries(templateSettings)) {
    let testValueRow = testValues[rowMap.get(k)];
    if (k.endsWith(ID_MATCHER)) {
      if (String(v) === '') {
        testValueRow[0] = getFileNameFromId(templateSettings[TSETTING_DOC_TEMPLATE_ID]);
      } else {
        testValueRow[0] = getFileNameFromId(v);
      }
    } else if (k.endsWith(FORMAT_MATCHER)) {
      if (String(v) === '') {
        testValueRow[0] = filledTemplateSettings[TSETTING_DOC_TEMPLATE_ID];
      } else {
        testValueRow[0] = textRenderer.render(interpolator.interpolate(v).items);
      }
    } else testValueRow[0] = v;
    filledTemplateSettings[k] = testValueRow[0];
  }

  testRange.setValues(testValues);
  Logger.log((testValues));
  return filledTemplateSettings;
}

function merge() {
  const { rowMap, templateSettings } = loadTemplateSettings(true);
  const filledTemplateSettings = fillTemplateSettingsTestCol();
  const functionsMap = getFunctionsMap();
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
    Logger.log(`Opening template '${templateSettings[TSETTING_DOC_TEMPLATE_ID]}'`);
    const file = DriveApp.getFileById(templateSettings[TSETTING_DOC_TEMPLATE_ID]);
    const fileName = filledTemplateSettings[TSETTING_DOC_TEMPLATE_ID];
    Logger.log(`Retrieving folders`);
    const originalFolderId = file.getParents().next().getId();
    const templateFolderId = templateSettings[TSETTING_DOC_FOLDER_ID] !== '' ? templateSettings[TSETTING_DOC_FOLDER_ID] : originalFolderId;
    const pdfFolderId = templateSettings[TSETTING_PDF_FOLDER_ID] !== '' ? templateSettings[TSETTING_PDF_FOLDER_ID] : originalFolderId;
    const templateFolder = DriveApp.getFolderById(templateFolderId);
    const pdfFolder = pdfFolderId === templateFolderId ? templateFolder : DriveApp.getFolderById(pdfFolderId);
  
    const functions = getFunctionsMap();
    const commands = getCommandsMap();
    const docRenderer = new DocRenderer();

    let mergeDoc = {
      file: null,
      fileName: null,
      doc: null,
      body: null,
      firstPara: null,
      header: null,
      footer: null,
    };
    let mergeAll = toBool(filledTemplateSettings[TSETTING_MERGE_ALL]);
    if (mergeAll) {
      mergeDoc.fileName = fileName + " _MERGE_ALL";
      const copy = file.makeCopy(mergeDoc.fileName, templateFolder);
      mergeDoc.doc = DocumentApp.openById(copy.getId());
      mergeDoc.body = mergeDoc.doc.getBody();
      mergeDoc.body.clear();

      // store first empty paragraph (will be removed later)
      mergeDoc.firstPara = mergeDoc.body.getChild(0);

      // clear and store header/footer (for later use)
      const tab = mergeDoc.doc.getTabs()[0].asDocumentTab();
      mergeDoc.header = tab.getHeader();
      mergeDoc.header?.clear();
      mergeDoc.footer = tab.getFooter();
      mergeDoc.footer?.clear();
    }

    let docs = [];

    let isFirstDoc = true;
    for (let rowIndex=0; rowIndex < 2; rowIndex++) {
      const copyName = String(templateSettings[TSETTING_DOC_NAME_FORMAT]) === '' ? fileName + '_' + String(rowIndex).padStart(2, "0") : filledTemplateSettings[TSETTING_DOC_NAME_FORMAT];
      const copy = file.makeCopy(copyName, templateFolder);
      const pdfName = String(templateSettings[TSETTING_PDF_NAME_FORMAT]) === '' ? fileName + '_' + String(rowIndex).padStart(2, "0") : filledTemplateSettings[TSETTING_PDF_NAME_FORMAT];

      const doc = DocumentApp.openById(copy.getId());
      const body = doc.getBody();

      const dataRow = mergeDisplayData[rowIndex];

      Logger.log(`Processing row ${rowIndex}`);

      const matches = findPlaceholders(body, DOC_PLACEHOLDERS_PATTERN);
      // iterate matches backwards
      for (let i = matches.length - 1; i >= 0; i--) {
        const r = matches[i];

        const context = {
          dataRow,
          rowIndex,
          templateName: fileName,
          columnsMap: mergeColumnsMap,
          rangeElement: r,
          document: copy,
          body,
        };

        const interpolator = new Interpolator({
          context,
          functions,
          commands,
        });

        const matchElement = docRenderer.getMatchFromRangeElement(r, context);
        const items = interpolator.interpolate(matchElement.matched).items;
        docRenderer.render(items, matchElement);

        //Logger.log(JSON.stringify(["AFTER:", textElement.getAttributes(start)]));
      }

      docs.push(doc);
      
      if (mergeAll) {
        appendDocTo(doc, mergeDoc.doc, { header:isFirstDoc, footer:isFirstDoc });
        if (isFirstDoc) {
          mergeDoc.firstPara.removeFromParent();
          if (mergeDoc.header) mergeDoc.header.getChild(0).removeFromParent();
          if (mergeDoc.footer) mergeDoc.footer.getChild(0).removeFromParent();
        }
        mergeDoc.body.appendPageBreak();
      }
      
      doc.saveAndClose();
      Logger.log(`Filled doc saved as '${copyName}' (in '${filledTemplateSettings[TSETTING_DOC_FOLDER_ID]}' folder)`);

      // convert to pdf
      let pdfFile = null;
      if (toBool(filledTemplateSettings[TSETTING_CREATE_PDF])) {
        Logger.log(`Converting doc to pdf as '${pdfName}' (in '${filledTemplateSettings[TSETTING_PDF_FOLDER_ID]}' folder)`);
        try {
          pdfFile = convertDocToPdf(pdfName, doc, pdfFolder);
        } catch (error) {
          G.ui.alert(TOASTER.ERROR, error, G.ui.ButtonSet.OK);
          throw error;
        }
      }

      // delete doc copy
      if (!toBool(filledTemplateSettings[TSETTING_KEEP_DOC])) {
        Logger.log(`Deleting doc copy '${copyName}'`);
        copy.setTrashed(true);
      }

      isFirstDoc = false;
    }

    if (mergeAll) {
      mergeDoc.doc.saveAndClose();
      Logger.log(`Merged all into '${mergeDoc.fileName}'`);
      if (toBool(filledTemplateSettings[TSETTING_CREATE_PDF])) {
        try {
          pdfFile = convertDocToPdf(mergeDoc.fileName, mergeDoc.doc, templateFolder);
        } catch (error) {
          G.ui.alert(TOASTER.ERROR, error, G.ui.ButtonSet.OK);
          throw error;
        }
      }
    }

  } /*catch (error) {
    G.ui.alert(TOASTER.ERROR, error, G.ui.ButtonSet.OK);
    throw error;
  }*/
}

// exported functions
var AzMM = {
  onOpen: onOpen,
  showTemplateSettings: showTemplateSettings,
  merge: merge
}

