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

function renderToText(items) {
  let result = [];

  for (const item of items) {
    switch (item.kind) {

      case "text": {
        result.push(item.value);
        break;
      }

      case "image": {
        break;
      }

      case "chart": {
        break;
      }

      case "table": {
        break;
      }

      case "pagebreak": {
        result.push("\n");
        break;
      }

      case "link": {
        result.push(`[${item.value}](${item.url})`);
        break;
      }

      default:
        
    }
  }

  return result.join("");
}

function getMatchFromRangeElement(rangeElement, context) {
  const textElement = rangeElement.getElement().asText();
  const text = textElement.getText();
  const start = rangeElement.getStartOffset();
  const endInclusive = rangeElement.getEndOffsetInclusive();
  const matched = text.slice(start, endInclusive + 1);

  return {
    context,
    rangeElement,
    textElement,
    text,
    start,
    endInclusive,
    matched
  }
}

function getMatchAttributes(matchElement) {
  // get formatting attributes (into a clone)
  const attrs = Object.assign({}, matchElement.textElement.getAttributes(matchElement.start));
  //Logger.log(JSON.stringify([matched, start, endInclusive]));
  //Logger.log(JSON.stringify(["BEFORE:", matchElement.textElement.getAttributes(matchElement.start)]));
  return attrs;
}

function setTextAttributes(textElement, start, endInclusive, attrs, onlyNonNull=true) {  
  // set attributes (optionally only non null values, to prevent not inheriting attributes from parent, and messing things up!)
  if (endInclusive > start) {
    for (const [attr, value] of Object.entries(attrs)) {
      if (value != null && onlyNonNull) {
        textElement.setAttributes(start, endInclusive, {
          [attr]: value
        });
      }
    }
  }
}

function setMatchAttributes(matchElement, start, endInclusive, attrs) {  
  setTextAttributes(matchElement.textElement, start, endInclusive, attrs);
}

function replaceMatchText(matchElement, text) {
  // delete text and insert replacement
  if (matchElement.endInclusive > matchElement.start) { // as deleteText(10, 10) will still erase one char
    matchElement.textElement.deleteText(matchElement.start, matchElement.endInclusive);
  }
  matchElement.textElement.insertText(matchElement.start, text);
}

function splitElement(element, start, endInclusive) {
  const before = element.copy();
  const after = element.copy();
  before.editAsText().deleteText(start, element.getText().length - 1);
  after.editAsText().deleteText(0, endInclusive);
  const parent = element.getParent();
  const childIndex = parent.getChildIndex(element);
  const type = element.getType();
  let insertFn = () => {};
  switch (type) {
    case DocumentApp.ElementType.PARAGRAPH:
      insertFn = parent.insertParagraph;
      break;
    case DocumentApp.ElementType.LIST_ITEM:
      insertFn = parent.insertListItem;
      break;
    default:
      throw new Error(`[split] Unsupported type ${type}`);
  }
  insertFn(childIndex, before);
  insertFn(childIndex + 1, after);
  element.removeFromParent();

  return {before, after, betweenIndex:childIndex + 1, parent};
}

function renderSplitItem(item, matchElement) {
  const parent = matchElement.rangeElement.getElement().getParent();
  const splitted = splitElement(parent, matchElement.start, matchElement.endInclusive);
}

function renderTextItem(item, matchElement, callback) {
  const attrs = getMatchAttributes(matchElement);
  replaceMatchText(matchElement, item.value);
  if (callback) callback(item, matchElement);
  setMatchAttributes(matchElement, matchElement.start, matchElement.start + item.value.length - 1, attrs);
  // update the position
  matchElement.start = matchElement.start + item.value.length;
  matchElement.endInclusive = matchElement.start;
}

function renderImageBlob(item, matchElement, blob) {
  const parent = matchElement.rangeElement.getElement().getParent();

  const splitted = splitElement(parent, matchElement.start, matchElement.endInclusive);
  const inlineImage = splitted.before.appendInlineImage(blob);
  splitted.after.merge();

  if (item.width || item.height) {
    const w = inlineImage.getWidth();
    const h = inlineImage.getHeight();
    const ratio = w / h;
    if (item.width) {
      inlineImage.setWidth(item.width);
      if (!item.height) inlineImage.setHeight(item.width / ratio);
    }    
    if (item.height) {
      inlineImage.setHeight(item.height);
      if (!item.width) inlineImage.setWidth(item.height * ratio);
    }
  }

  return inlineImage;
}

function renderImageItem(item, matchElement) {
  let blob = null;
  if (item.fileId) {
    const file = DriveApp.getFileById(item.fileId);
    blob = file.getAs('image/png');
  } else if (item.url) {
    const response = UrlFetchApp.fetch(item.url);
    blob = response.getBlob();
  }

  const inlineImage = renderImageBlob(item, matchElement, blob);
}

function renderChartItem(item, matchElement) {
  const sheet = G.ss.getSheetByName(item.src);
  const charts = sheet.getCharts();
  const chart = charts[0];
  const blob = chart.getAs('image/png');
  
  const inlineImage = renderImageBlob(item, matchElement, blob);
}

function renderTableItem(item, matchElement) {
  replaceMatchText(matchElement, "");
  const p = matchElement.rangeElement.getElement().getParent().asParagraph();
  const namedRange = G.ss.getRangeByName(item.src);
  const dataRange = namedRange ? namedRange : G.ss.getRange(item.src);
  if (dataRange == null) return;

  const values = dataRange.getDisplayValues();
  const body = matchElement.context.body;
  const childIndex = body.getChildIndex(p);

  if (!item.format) {
    body.insertTable(childIndex, values);
  } else {
    const table = body.insertTable(childIndex);
    const rangeRuns = getRangeRuns(dataRange);
    
    rangeRuns.forEach(row => {
      let tableRow = table.appendTableRow();
      row.forEach(cell => {
        let tableCell = tableRow.appendTableCell();
        let text = tableCell.editAsText();
        const cellText = cell.map((run) => run.text).join("");
        text.setText(cellText);
        cell.forEach(run => {
          const bg = run.attrs["BACKGROUND_COLOR"];
          run.attrs["BACKGROUND_COLOR"] = null;
          tableCell.setBackgroundColor(bg);
          setTextAttributes(text, run.start, run.endInclusive, run.attrs);
        });
      })
    });
  }
}

function renderPageBreakItem(item, matchElement) {
  replaceMatchText(matchElement, "");
  const p = matchElement.rangeElement.getElement().getParent();
  p.insertPageBreak(0);
}

function renderLinkItem(item, matchElement) {
  const start = matchElement.start;
  renderTextItem(item, matchElement, (item, matchElement) => {
    const endInclusive = start + item.value.length - 1;
    matchElement.textElement.setLinkUrl(start, endInclusive, item.url);
  });
}

function renderToMatchElement(items, matchElement) {
  for (const item of items) {
    Logger.log(JSON.stringify([matchElement.matched, item]));
    switch (item.kind) {
      case "text": {
        renderTextItem(item, matchElement);
        break;
      }

      case "image": {
        renderImageItem(item, matchElement);
        break;
      }

      case "chart": {
        renderChartItem(item, matchElement);
        break;
      }

      case "table": {
        renderTableItem(item, matchElement);
        break;
      }

      case "pagebreak": {
        renderPageBreakItem(item, matchElement);
        break;
      }

      case "link": {
        renderLinkItem(item, matchElement);
        break;
      }

      case "split": {
        renderSplitItem(item, matchElement);
        break;
      }

      default:
        
    }
  }
}

