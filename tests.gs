function test_getFileNameFromId() {
  SpreadsheetApp.getUi().alert(getFileNameFromId('1BauByY7pvAU5LoL_zsaCRWQl8jyklTe_'));
}

function test_A1Notation() {
  const toA1Tests = [
    [1, 3],
    [null, 3],
    [3, null],
    [10, 52]
  ];
  for (let test of toA1Tests) {
    Logger.log(test + " -> " + toA1Notation(...test) + " " + toA1Notation(...test).join(''));
  }

  const fromA1Tests = [
    'C4', '$C$4', '$AZ', '21', '$21'
  ];
  for (let test of fromA1Tests) {
    Logger.log(test + " -> " + JSON.stringify(parseA1Notation(test)));
  }
}


function test_Interpolator() {
  const INDENT = 0;
  const templates = [
    "before{{name}} - {{NOW}}after",
    "before{{cod_fisc}} - {{NOW}}after",
    "before{{{x:32}}} - {{NOW}}after {{chart+}}",
    "before{{{x:32}}} - {{NOW}}after {{ROW_INDEX}} {{ROW_NUMBER}}",
    'before{{{x:32}}} - {{NOW}}after {{{"type":"NUMBER", "value":12, "format":"%03.2f"}}} {{ROW_NUMBER}}',
    '{{{"type":"CHART", "value":12, "format":"%03.2f"}}}',
  ];
  let { metaColumnsMap, metaData, mergeColumnsMap, mergeData, mergeDisplayData } = getMergeData();
  const dataRow = mergeData.length > 0 ? mergeData[0] : null;
  const context = { dataRow, rowIndex: 12, columnsMap:mergeColumnsMap };
  for (let template of templates) {
    const interpolator = new Interpolator({
      context,
      functions: getFunctionsMap(),
      commands: getCommandsMap(),
    });

    const result = interpolator.interpolate(template);
    Logger.log("TEST  : " + JSON.stringify(template));
    Logger.log("RESULT: " + JSON.stringify(result, null, INDENT));
  }
}

function test_Utilities() {
  const num = Number("1");
  Logger.log(Utilities.formatString("%2.3f", num));
  Logger.log(Utilities.formatString("%03.2f", num));
  Logger.log(Utilities.formatString("%03d", num));
}

function test_getChart() {
  const sheetName = "data";
  const sheet = G.ss.getSheetByName(sheetName);
  const charts = sheet.getCharts();
  Logger.log(`Charts on '${sheetName}': ${charts.length}`);
  const blob = charts[0].getAs('image/png');
}

function test_getTableData() {
  const sheetName = "data";
  const namedRange = "EMP_DATA";
  const sheet = G.ss.getSheetByName(sheetName);
  let range = G.ss.getRangeByName(namedRange);
  if (range) {
    Logger.log(JSON.stringify(range.getValues()));
  }
  const rangeStr = "data!A1:B7";
  range = G.ss.getRange(rangeStr);
  if (range) {
    Logger.log(JSON.stringify(range.getValues()));
  }
}

function test_getRichTextValues() {
  const sheetName = "data";
  const rangeStr = "data!C2:C3";
  range = G.ss.getRange(rangeStr);
  if (range) {
    Logger.log(JSON.stringify(getRangeRuns(range)));
  }
}

function test_docStructure() {
  const {rowMap, templateSettings} = loadTemplateSettings(true);
  const doc = DocumentApp.openById(templateSettings[TSETTING_DOC_TEMPLATE_ID]);
  const tabs = doc.getTabs();
  Logger.log(tabs);
  Logger.log(JSON.stringify(tabs.map((t) => t.getId())));
  const footer = tabs[0].asDocumentTab().getFooter();
  Logger.log(footer.getText());
  Logger.log(footer.getType());
}

function getSectionStructure(section) {
  let tree = {};
  let parentsMap = {};
  let id = 0;
  let stack = [[{elementId:id, element:section}]];
  let step = 0;
  let level = 0;
  const MAX_STEPS = 500;
  let maxStepsReached = false;
  while (stack.length > 0 && !maxStepsReached) {
    const stackElements = stack.pop();
    for (let stackElement of stackElements) {
      const { elementId, element } = stackElement;
      const numChildren = element.getNumChildren ? element.getNumChildren() : 0;
      const type = element.getType ? element.getType() : null;
      const text = type === DocumentApp.ElementType.TEXT ? element.getText() : null;

      let node = { elementId, type, level:null, children: [], text };
      //Logger.log(JSON.stringify(node));

      if (step == 0) {
        tree = node;
        node.level = 0;
      } else {
        const parent = parentsMap[node.elementId];
        parent.children.push(node);
        node.level = parent.level + 1;
        level = Math.max(level, node.level);
      }
      delete node.elementId;
      if (node.text == null) delete node.text;

      let children = [];
      for (let childIndex = 0; childIndex < numChildren; childIndex++) {
        id++;
        const child = element.getChild(childIndex);
        children.push({ elementId:id, element:child });
        parentsMap[id] = node;
      }
      if (children.length > 0) {
        stack.push(children);
      }
    }
    step++;
    maxStepsReached = step >= MAX_STEPS;
  }

  Logger.log(`numElements: ${id + 1} steps:${step} maxLevel:${level}`);
  return { tree, maxStepsReached };
}

/**
 * Returns `element` as properly typed.
 * `typedAs` can be:
 *  - null: type will be what's returned by `element.getType()`
 *  - one of DocumentApp.ElementType
 *  - another element: type will be what's returned by `typedAs.getType()`
 */
function getAsTyped(element, typedAs=null) {
  const type = typedAs == null ? element.getType() : (typeof typedAs.getType === "function" ? typedAs.getType() : typedAs);
  let typed = null;

  switch (type) {
    case DocumentApp.ElementType.BODY_SECTION: {
      typed = element.asBody();
      break;
    }
    case DocumentApp.ElementType.COMMENT_SECTION: {
      typed = element.asCommentSection();
      break;
    }
    case DocumentApp.ElementType.DATE: {D
      typed = element.asate();
      break;
    }
    case DocumentApp.ElementType.EQUATION: {
      typed = element.asEquation();
      break;
    }
    case DocumentApp.ElementType.EQUATION_FUNCTION: {
      typed = element.asEquationFunction();
      break;
    }
    case DocumentApp.ElementType.EQUATION_FUNCTION_ARGUMENT_SEPARATOR: {
      typed = element.asEquationFunctionArgumentSeparator();
      break;
    }
    case DocumentApp.ElementType.EQUATION_SYMBOL: {
      typed = element.asEquationSymbol();
      break;
    }
    case DocumentApp.ElementType.RICH_LINK: {
      typed = element.asRichLink();
      break;
    }
    case DocumentApp.ElementType.FOOTER_SECTION: {
      typed = element.asFooterSection();
      break;
    }
    case DocumentApp.ElementType.FOOTNOTE: {
      typed = element.asFootnote();
      break;
    }
    case DocumentApp.ElementType.FOOTNOTE_SECTION: {
      typed = element.asFootnoteSection();
      break;
    }
    case DocumentApp.ElementType.HEADER_SECTION: {
      typed = element.asHeaderSection();
      break;
    }
    case DocumentApp.ElementType.HORIZONTAL_RULE: {
      typed = element.asHorizontalRule();
      break;
    }
    case DocumentApp.ElementType.INLINE_DRAWING: {
      typed = element.asInlineDrawing();
      break;
    }
    case DocumentApp.ElementType.INLINE_IMAGE: {
      typed = element.asInlineImage();
      break;
    }
    case DocumentApp.ElementType.LIST_ITEM: {
      typed = element.asListItem();
      break;
    }
    case DocumentApp.ElementType.PAGE_BREAK: {
      typed = element.asPageBreak();
      break;
    }
    case DocumentApp.ElementType.PARAGRAPH: {
      typed = element.asParagraph();
      break;
    }
    case DocumentApp.ElementType.PERSON: {
      typed = element.asPerson();
      break;
    }
    case DocumentApp.ElementType.TABLE: {
      typed = element.asTable();
      break;
    }
    case DocumentApp.ElementType.TABLE_CELL: {
      typed = element.asTableCell();
      break;
    }
    case DocumentApp.ElementType.TABLE_OF_CONTENTS: {
      typed = element.asTableOfContents();
      break;
    }
    case DocumentApp.ElementType.TABLE_ROW: {
      typed = element.asTableRow();
      break;
    }
    case DocumentApp.ElementType.TEXT: {
      typed = element.asText();
      break;
    }
    case DocumentApp.ElementType.UNSUPPORTED:
    default:
      throw new Error("UNSUPPORTED ELEMENT TYPE");
  }

  return typed;
}

function getDocStructure() {
  //const DOC_ID = "1CA7WCoXleWHWqT2dEyjt3DDxTwZKlcWli_kz9Y767-Q"; // with inline image inside paragraph
  //const DOC_ID = "1Xtyi4fxIh3EDZ-e97Bd-Mop5lLgQbiG4JfFGSu5Pel4"; // without
  //const DOC_ID = "1VaS4kPTCm6kAon5pI7RY-1H6Rd19g6Fj0hm0m1BepLg"; // test
  const DOC_ID = "19VfFKqV5fihG5RkLcSTAK8QXV_uT2qFC5mLcN2T2Pdg";
  
  const doc = DocumentApp.openById(DOC_ID);
  const body = doc.getBody();

  const { tree, maxStepsReached } = getSectionStructure(body);

  //Logger.log(JSON.stringify(tree));
  const sheet = G.ss.getSheetByName("log");
  sheet.appendRow([maxStepsReached, JSON.stringify(tree)]);
  if (maxStepsReached) Logger.log("MAX_STEPS REACHED!");
}

