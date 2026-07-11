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

function test_interpolateTemplateString() {
  const INDENT = 0;
  let functionsMap = getFunctionsMap();
  const template1 = "before{{name}} - {{NOW}}after";
  const template2 = "before{{{x:32}}} - {{NOW}}after {{chart+}}";
  const template3 = "before{{{x:32}}} - {{NOW}}after {{ROW_INDEX}} {{ROW_NUMBER}}";
  const template4 = 'before{{{x:32}}} - {{NOW}}after {{{"type":"NUMBER", "value":12, "format":"%03.2f"}}} {{ROW_NUMBER}}';
  const testStrings = [template1, template2, template3, template4];
  let { metaColumnsMap, metaData, mergeColumnsMap, mergeData, mergeDisplayData } = getMergeData();
  const dataRow = mergeData.length > 0 ? mergeData[0] : null;
  const context = { dataRow, columnsMap:mergeColumnsMap };
  for (let template of testStrings) {
    const tokens = tokenize(template);
    const res = interpolateTemplateString(template, context, functionsMap);
    Logger.log("TEST  : " + JSON.stringify(template));
    Logger.log("TOKENS: " + JSON.stringify(tokens, null, INDENT));
    Logger.log("INTERP: " + JSON.stringify(res, null, INDENT));
    Logger.log("RESULT: " + res.interpolated);
  }
}

function test_Utilities() {
  const num = Number("1");
  Logger.log(Utilities.formatString("%2.3f", num));
  Logger.log(Utilities.formatString("%03.2f", num));
  Logger.log(Utilities.formatString("%03d", num));
}
