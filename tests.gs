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
  const INDENT = 2;
  const testStrings = [];
  let cachedMap = null;
  const template1 = "before{{name}} - {{NOW}}after";
  const template2 = "before{{{x:32}}} - {{NOW}}after {{chart+}}";
  let { metaColumnsMap, metaData, mergeColumnsMap, mergeData, mergeDisplayData } = getMergeData();
  const dataRow = mergeData.length > 0 ? mergeData[0] : null;
  for (let template of [template1, template2]) {
    const tokens = tokenize(template);
    const res = interpolateTemplateString(template, dataRow, mergeColumnsMap, cachedMap);
    Logger.log("TEST  : " + JSON.stringify(template));
    Logger.log("TOKENS: " + JSON.stringify(tokenize(template), null, INDENT));
    Logger.log("INTERP: " + JSON.stringify(res, null, INDENT));
    Logger.log("RESULT: " + res.interpolated);
  }
}

