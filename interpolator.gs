function getFunctionsMap() {
  return new Map([
    ["NOW", () => new Date().toLocaleString()],
    ["ROW_INDEX", ({ rowIndex }) => (rowIndex ?? 0) + 1],
  ]);
}

function getCommandsMap() {
  return new Map([
    ["NUMBER", ({ args }) =>
      Utilities.formatString(args.format, Number(args.value))
    ],

    ["IMAGE", ({ args }) => ({
      type: "image",
      fileId: args.fileId,
    })],

    ["CHART", ({ args }) => ({
      type: "chart",
      options: args,
    })],
  ]);
}

class Interpolator {
  constructor({
    context = {},
    functions = new Map(),
    commands = new Map(),
  } = {}) {
    this.context = context;
    this.functions = functions;
    this.commands = commands;
    this.columnsMap = context.columnsMap ? context.columnsMap : new Map();
  }

  interpolate(template) {
    const processedTokens = this.tokenize(template)
      .map(token => this.resolve(token));

    return {
      interpolated: processedTokens.map(t => t.replacedWith).join(""),
      processedTokens,
    };
  }

  tokenize(string) {
    const tokens = [];
    let pos = 0;

    for (const match of string.matchAll(TOKENIZER_REGEXP)) {
      if (match.index > pos) {
        tokens.push({
          raw: string.slice(pos, match.index),
          key: null,
          start: pos,
        });
      }

      tokens.push({
        raw: match[0],
        key: match[0].replace(/^{{|}}$/g, ""),
        start: match.index,
      });

      pos = match.index + match[0].length;
    }

    if (pos < string.length) {
      tokens.push({
        raw: string.slice(pos),
        key: null,
        start: pos,
      });
    }

    return tokens;
  }

  resolve(token) {
    if (!token.key) {
      return {
        ...token,
        replacedWith: token.raw,
        status: REPLACE_STATUS.KEPT_RAW,
      };
    }

    const node = this.parse(token);

    switch (node.kind) {
      case "field":
        return this.resolveField(token, node);

      case "function":
        return this.resolveFunction(token, node);

      case "command":
        return this.resolveCommand(token, node);

      default:
        return {
          ...token,
          replacedWith: token.raw,
          status: REPLACE_STATUS.NOT_FOUND,
        };
    }
  }

  parse(token) {
    const key = token.key;

    // {{{ ...json... }}}
    if (key.startsWith("{")) {
      try {
        const obj = JSON.parse(key);

        return {
          kind: "command",
          name: obj.type,
          args: obj,
        };
      } catch {
        return { kind: "invalid" };
      }
    }

    if (this.functions.has(key)) {
      return {
        kind: "function",
        name: key,
      };
    }

    return {
      kind: "field",
      name: key,
    };
  }

  resolveField(token, node) {
    const idx = this.columnsMap.get(node.name);

    if (idx == null)
      return {
        ...token,
        replacedWith: token.raw,
        status: REPLACE_STATUS.NOT_FOUND,
      };

    const value = this.context.dataRow[idx];

    const result = this.interpolate(String(value));

    return {
      ...token,
      replacedWith: result.interpolated,
      status: REPLACE_STATUS.OK,
    };
  }

  resolveFunction(token, node) {
    const fn = this.functions.get(node.name);

    return {
      ...token,
      replacedWith: String(fn(this.context)),
      status: REPLACE_STATUS.OK,
    };
  }

  resolveCommand(token, node) {
    const fn = this.commands.get(node.name);

    if (!fn)
      return {
        ...token,
        replacedWith: token.raw,
        status: REPLACE_STATUS.NOT_FOUND,
      };

    return {
      ...token,
      replacedWith: fn({
        ...this.context,
        args: node.args,
      }),
      status: REPLACE_STATUS.OK,
    };
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

