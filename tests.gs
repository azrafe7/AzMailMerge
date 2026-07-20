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

function test_footerHeader() {
  const {rowMap, templateSettings} = loadTemplateSettings(true);
  const doc = DocumentApp.openById(templateSettings[TSETTING_DOC_TEMPLATE_ID]);
  const tabs = doc.getTabs();
  Logger.log(tabs);
  Logger.log(JSON.stringify(tabs.map((t) => t.getId())));
  const footer = tabs[0].asDocumentTab().getFooter();
  Logger.log(footer.getText());
  Logger.log(footer.getType());
}

function test_getSectionStructure() {
  //const DOC_ID = "1CA7WCoXleWHWqT2dEyjt3DDxTwZKlcWli_kz9Y767-Q"; // with inline image inside paragraph
  //const DOC_ID = "1Xtyi4fxIh3EDZ-e97Bd-Mop5lLgQbiG4JfFGSu5Pel4"; // without
  //const DOC_ID = "1VaS4kPTCm6kAon5pI7RY-1H6Rd19g6Fj0hm0m1BepLg"; // test
  const DOC_ID = "1t5zeo633NvEjWByJQIT_8hea9XE_L9PDd5zH2yW5ZLI";
  
  const doc = DocumentApp.openById(DOC_ID);
  const body = doc.getBody();

  const { tree, maxStepsReached } = getSectionStructure(body);

  //Logger.log(JSON.stringify(tree));
  logToSheet([maxStepsReached, JSON.stringify(tree)]);
  if (maxStepsReached) Logger.log("MAX_STEPS REACHED!");
}

// for slides
function getColorAsRgb(color, colorScheme) {
  const colorType = color.getColorType();
  let concreteColor = null;
  if (colorType === SlidesApp.ColorType.RGB) {
    concreteColor = color.asRgbColor();
  } else if (colorType === SlidesApp.ColorType.THEME) {
    const themeColor = color.asThemeColor().getThemeColorType();
    concreteColor = colorScheme.getConcreteColor(themeColor).asRgbColor();
  } else {
    throw new Error(`Unsupported color type ${colorType}`);
  }

  return concreteColor;
}

// for slides
function getTextRangeAttrs(tr, colorScheme) {
  let docRuns = [];
  for (let run of tr.getRuns()) {
    docRun = {};
    const text = run.asString();
    let attrs = {};
    
    // defaults
    attrs["STRIKETHROUGH"] = null;
    attrs["ITALIC"] = null;
    attrs["FOREGROUND_COLOR"] = null;
    attrs["BACKGROUND_COLOR"] = null;
    attrs["LINK_URL"] = null;
    attrs["UNDERLINE"] = null,
    attrs["FONT_SIZE"] = null;
    attrs["FONT_FAMILY"] = null;
    attrs["FONT_WEIGHT"] = null;
    attrs["BOLD"] = null;

    let start = run.getStartIndex();
    let end = run.getEndIndex();
    if (start < end) {
      let textStyle = run.getTextStyle();
      let fontFamily = textStyle.getFontFamily();
      let fontSize = textStyle.getFontSize();
      let fontWeight = textStyle.getFontWeight();
      let fontColor = getColorAsRgb(textStyle.getForegroundColor(), colorScheme)?.asHexString();
      let background = textStyle.isBackgroundTransparent ? null : getColorAsRgb(textStyle.getBackgroundColor(), colorScheme)?.asHexString();
      let bold = textStyle.isBold();
      let italic = textStyle.isItalic();
      let strikeThrough = textStyle.isStrikethrough();
      let underLine = textStyle.isUnderline();

      attrs["STRIKETHROUGH"] = strikeThrough;
      attrs["ITALIC"] = italic;
      attrs["FOREGROUND_COLOR"] = fontColor;
      attrs["BACKGROUND_COLOR"] = background;
      attrs["LINK_URL"] = null;
      attrs["UNDERLINE"] = underLine,
      attrs["FONT_SIZE"] = fontSize;
      attrs["FONT_FAMILY"] = fontFamily;
      //attrs["FONT_WEIGHT"] = fontWeight;
      attrs["BOLD"] = bold;
    }

    docRun.text = text;
    docRun.start = start;
    docRun.end = end;
    docRun["attrs"] = attrs;
    docRuns.push(docRun);
  }

  Logger.log(JSON.stringify(docRuns));
  return docRuns;
}

// for slides
function setTextRangeAttrs(tr, attrs, onlyNonNull=true) {  
  const fnMap = {
    "STRIKETHROUGH": "setStrikethrough",
    "ITALIC": "setItalic",
    "FOREGROUND_COLOR": "setForegroundColor",
    "BACKGROUND_COLOR": "setBackgroundColor",
    "LINK_URL": "setLinkUrl",
    "UNDERLINE": "setUnderline",
    "FONT_SIZE": "setFontSize",
    "FONT_FAMILY": "setFontFamily",
    //"FONT_WEIGHT": "setFontWeight",
    "BOLD": "setBold",
  };
  const textStyle = tr.getTextStyle();
  // set attributes (optionally only non null values, to prevent not inheriting attributes from parent, and messing things up!)
  for (const [attr, value] of Object.entries(attrs)) {
    if (value != null && onlyNonNull) {
      //Logger.log(fnMap[attr])
      //Logger.log(textStyle[fnMap[attr]])
      textStyle[fnMap[attr]](value);
    }
  }
}

function test_slidesPlaceholdersAttrs() {
  const DOC_ID = "1ByO7zNcjpEoCOX8th2d0n9LGH6uzyv6sBt0TWZ3o12g";
  const presentation = SlidesApp.openById(DOC_ID);

  const matches = findPresentationPlaceholders(presentation, SLIDES_PLACEHOLDERS_PATTERN);
  Logger.log(JSON.stringify(matches));

  for (let match of matches) {
    Logger.log(JSON.stringify(match));
    const attrs = getTextRangeAttrs(match.textRange, match.page.getColorScheme());
    Logger.log(JSON.stringify(attrs));
  }
}

function test_slidesPlaceholders() {
  const DOC_ID = "1ByO7zNcjpEoCOX8th2d0n9LGH6uzyv6sBt0TWZ3o12g";
  const presentation = SlidesApp.openById(DOC_ID);

  const { tree, maxStepsReached } = getPageElementStructure(presentation.getSlides()[0].getPageElements());

  //Logger.log(JSON.stringify(tree));
  logToSheet([maxStepsReached, JSON.stringify(tree)]);
  return;


  const pattern = "\\{\\{([^\\{}}]+|\\{[^}]+})}}";
  const matches = findPresentationPlaceholders(presentation, pattern);
  Logger.log(JSON.stringify(matches));

  const slides = presentation.getSlides();
  for (let slideIdx=0; slideIdx<slides.length; slideIdx++) {
    const slide = slides[slideIdx];
    Logger.log(`SLIDE ${slideIdx}`);
    for (let shape of slide.getShapes()) {
      const shapeTextElement = shape.getText();
      Logger.log(`TEXT: '${shapeTextElement.asString()}'`);
    }
  }
}

