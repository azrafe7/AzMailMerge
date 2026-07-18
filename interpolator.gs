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
        linkUrl: args.linkUrl,
        width: args.width,
        height: args.height,
      }
    ]],

    ["CHART", ({ args }) => [
      {
        kind: "chart",
        src: args.src,
        linkUrl: args.linkUrl,
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

    const tokenizer = new TemplateTokenizer(template);
    const tokens = tokenizer.tokenize();
    
    for (const token of tokens) {
      items.push(...this.resolve(token));
    }

    return { items };
  }

  resolve(token) {
    if (!token.key)
      return [textItem(token.raw)];

    // If the token contains templates, interpolate it first
    let key = token.key;
    if (key.includes("{{")) {
      // Interpolate the key content
      const interpolated = this.interpolate(key);
      // Merge text items back into a string
      key = interpolated.items
        .filter(item => item.kind === "text")
        .map(item => item.value)
        .join("");
    }

    const node = this.parse({ ...token, key });

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

    // Recursively interpolate any string values in args that contain templates
    const interpolatedArgs = this.interpolateArgs(node.args);

    return fn({
      ...this.context,
      args: interpolatedArgs
    });
  }

  interpolateArgs(args) {
    const interpolated = {};
    
    for (const [key, value] of Object.entries(args)) {
      if (typeof value === "string" && value.includes("{{")) {
        // Recursively interpolate this string value
        const result = this.interpolate(value);
        // If it results in plain text, use the combined text
        interpolated[key] = result.items
          .filter(item => item.kind === "text")
          .map(item => item.value)
          .join("");
      } else if (typeof value === "object" && value !== null && !Array.isArray(value)) {
        // Recursively handle nested objects
        interpolated[key] = this.interpolateArgs(value);
      } else {
        interpolated[key] = value;
      }
    }
    
    return interpolated;
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

class TemplateTokenizer {
  constructor(string) {
    this.string = string;
    this.pos = 0;
  }

  tokenize() {
    const tokens = [];
    
    while (this.pos < this.string.length) {
      const textToken = this.readText();
      if (textToken) {
        tokens.push(textToken);
      }
      
      if (this.pos < this.string.length) {
        const templateToken = this.readTemplate();
        if (templateToken) {
          tokens.push(templateToken);
        } else {
          // Malformed, just consume {{
          tokens.push({
            raw: "{{",
            key: null,
            start: this.pos,
          });
          this.pos += 2;
        }
      }
    }
    
    return tokens;
  }

  readText() {
    const start = this.pos;
    while (this.pos < this.string.length) {
      if (this.string[this.pos] === "{" && this.string[this.pos + 1] === "{") {
        break;
      }
      this.pos++;
    }

    if (this.pos > start) {
      return {
        raw: this.string.slice(start, this.pos),
        key: null,
        start,
      };
    }
    return null;
  }

  readTemplate() {
    const start = this.pos;
    
    if (this.string[this.pos] !== "{" || this.string[this.pos + 1] !== "{") {
      return null;
    }

    this.pos += 2; // consume {{

    // Check if it's a JSON command
    this.skipWhitespace();
    const isJson = this.string[this.pos] === "{";

    let content;
    if (isJson) {
      content = this.readJsonCommand();
    } else {
      content = this.readPlaceholder();
    }

    if (!content) {
      this.pos = start; // Reset on error
      return null;
    }

    // Consume closing }}
    this.skipWhitespace();
    if (this.string[this.pos] !== "}" || this.string[this.pos + 1] !== "}") {
      this.pos = start;
      return null;
    }

    const endPos = this.pos + 2;
    const raw = this.string.slice(start, endPos);
    
    this.pos = endPos;

    return {
      raw,
      key: content,
      start,
    };
  }

  readJsonCommand() {
    const start = this.pos;
    let depth = 0;

    while (this.pos < this.string.length) {
      const char = this.string[this.pos];

      if (char === '"') {
        // Skip string
        this.pos++;
        this.skipQuotedString();
        continue;
      }

      if (char === "{") {
        depth++;
      } else if (char === "}") {
        depth--;
        if (depth === 0) {
          // Found closing }
          const content = this.string.slice(start, this.pos + 1);
          this.pos++;
          return content;
        }
      }

      this.pos++;
    }

    return null;
  }

  readPlaceholder() {
    const start = this.pos;

    while (this.pos < this.string.length) {
      if (this.string[this.pos] === "}" && this.string[this.pos + 1] === "}") {
        const content = this.string.slice(start, this.pos);
        return content.trim();
      }
      this.pos++;
    }

    return null;
  }

  skipQuotedString() {
    // this.pos should be after the opening "
    while (this.pos < this.string.length) {
      const char = this.string[this.pos];

      if (char === "\\") {
        this.pos += 2; // Skip escaped char
        continue;
      }

      if (char === '"') {
        this.pos++;
        break;
      }

      this.pos++;
    }
  }

  skipWhitespace() {
    while (this.pos < this.string.length && /\s/.test(this.string[this.pos])) {
      this.pos++;
    }
  }
}

