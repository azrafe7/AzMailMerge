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
