function textItem(value) {
  return {
    kind: "text",
    value: String(value)
  };
}

function pageBreakItem() {
  return {
    kind: "pagebreak",
  };
}

function splitItem() {
  return {
    kind: "split",
  };
}

function getFunctionsMap() {
  return new Map([
    ["NOW", () => [textItem(new Date().toLocaleString())]],
    ["ROW_INDEX", ({ rowIndex }) => [textItem((rowIndex ?? 0) + 1)]],
    ["PADDED_ROW_INDEX", ({ rowIndex }) => [textItem(String((rowIndex ?? 0) + 1).padStart(2, "0"))]],
    ["PAGEBREAK", () => [pageBreakItem()]],
    ["SPLIT", () => [splitItem()]],
    ["TEMPLATE_NAME", ({ templateName }) => [textItem(templateName)]],
  ]);
}

function getCommandsMap() {
  return new Map([
    ["NUMBER", ({ args }) => [
      textItem(
        Utilities.formatString(args.format, Number(args.value))
      )
    ]],

    ["IMAGE", ({ args }) => [
      {
        kind: "image",
        fileId: args.fileId,
        url: args.url,
        width: args.width,
        height: args.height,
      }
    ]],

    ["CHART", ({ args }) => [
      {
        kind: "chart",
        src: args.src,
        width: args.width,
        height: args.height,
      }
    ]],

    ["TABLE", ({ args }) => [
      {
        kind: "table",
        src: args.src,
        format: args.format,
      }
    ]],

    ["LINK", ({ args }) => [
      {
        kind: "link",
        value: args.value,
        url: args.url,
      }
    ]],
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
    this.columnsMap = context.columnsMap ?? new Map();
  }

  interpolate(template) {
    const items = [];

    for (const token of this.tokenize(template)) {
      items.push(...this.resolve(token));
    }

    return {
      items
      //items: this.mergeAdjacentText(items)
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
        key: match[0].replace(DETOKENIZER_REGEXP, ""),
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
    if (!token.key)
      return [textItem(token.raw)];

    const node = this.parse(token);

    switch (node.kind) {
      case "field":
        return this.resolveField(node);

      case "function":
        return this.resolveFunction(node);

      case "command":
        return this.resolveCommand(node);

      default:
        return [textItem(token.raw)];
    }
  }

  parse(token) {
    const key = token.key;

    if (key.startsWith("{")) {
      try {
        const obj = JSON.parse(key);

        return {
          kind: "command",
          name: obj.type,
          args: obj,
        };
      } catch {
        return {
          kind: "invalid"
        };
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

  resolveField(node) {
    const idx = this.columnsMap.get(node.name);

    if (idx == null)
      return [textItem(`{{${node.name}}}`)];

    return this.interpolate(
      String(this.context.dataRow[idx])
    ).items;
  }

  resolveFunction(node) {
    const fn = this.functions.get(node.name);

    return fn(this.context);
  }

  resolveCommand(node) {
    const fn = this.commands.get(node.name);

    if (!fn)
      return [textItem(`{{${JSON.stringify(node.args)}}}`)];

    return fn({
      ...this.context,
      args: node.args
    });
  }

  mergeAdjacentText(items) {
    const merged = [];

    for (const item of items) {
      const last = merged[merged.length - 1];

      if (
        last &&
        last.kind === "text" &&
        item.kind === "text"
      ) {
        last.value += item.value;
      } else {
        merged.push(item);
      }
    }

    return merged;
  }
}

