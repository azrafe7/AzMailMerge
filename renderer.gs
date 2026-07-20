class TextRenderer {
  constructor() { }

  render(items) {
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
}

class DocRenderer {
  constructor() { 
    this.cursor = null;
  }

  render(items, matchElement) {
    this.attrs = this.getMatchAttributes(matchElement);
    Logger.log(JSON.stringify(this.attrs));
    this.cursor = this.splitMatch(matchElement);

    for (const item of items) {
      Logger.log(JSON.stringify([matchElement.matched, item]));
      switch (item.kind) {
        case "text": {
          this.renderTextItem(item, matchElement);
          break;
        }

        case "image": {
          this.renderImageItem(item, matchElement);
          break;
        }

        case "chart": {
          this.renderChartItem(item, matchElement);
          break;
        }

        case "table": {
          this.renderTableItem(item, matchElement);
          break;
        }

        case "pagebreak": {
          this.renderPageBreakItem(item, matchElement);
          break;
        }

        case "link": {
          this.renderLinkItem(item, matchElement);
          break;
        }

        case "split": {
          this.renderSplitItem(item, matchElement);
          break;
        }

        default:
          
      }
    }

    this.cursor.after.merge();
  }

  getMatchFromRangeElement(rangeElement, context) {
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

  getMatchAttributes(matchElement) {
    // get formatting attributes (into a clone)
    const attrs = Object.assign({}, matchElement.textElement.getAttributes(matchElement.start));
    //Logger.log(JSON.stringify([matched, start, endInclusive]));
    //Logger.log(JSON.stringify(["BEFORE:", matchElement.textElement.getAttributes(matchElement.start)]));
    return attrs;
  }

  setTextAttributes(textElement, start, endInclusive, attrs, onlyNonNull=true) {  
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

  setMatchAttributes(matchElement, start, endInclusive, attrs) {  
    this.setTextAttributes(matchElement.textElement, start, endInclusive, attrs);
  }

  splitMatch(matchElement) {
    const parent = matchElement.rangeElement.getElement().getParent();
    const textElementIndex = parent.getChildIndex(matchElement.textElement);
    const before = parent;
    const after = parent.copy();
    const numChildren = parent.getNumChildren();

    // remove matched text and everything after it
    for (let childIndex = numChildren - 1; childIndex > textElementIndex; childIndex--) {
      before.getChild(childIndex).removeFromParent();
    }
    let matchedTextElement = matchElement.textElement;
    matchedTextElement.deleteText(matchElement.start, matchElement.text.length - 1);

    // remove matched text and everything before it
    for (let childIndex = textElementIndex - 1; childIndex >= 0; childIndex--) {
      after.getChild(childIndex).removeFromParent();
    }
    matchedTextElement = after.getChild(0);
    matchedTextElement.deleteText(0, matchElement.endInclusive);

    const grandParent = parent.getParent();
    const childIndex = grandParent.getChildIndex(parent);
    const type = parent.getType();
    let insertFn = () => {};
    switch (type) {
      case DocumentApp.ElementType.PARAGRAPH:
        insertFn = grandParent.insertParagraph;
        break;
      case DocumentApp.ElementType.LIST_ITEM:
        insertFn = grandParent.insertListItem;
        break;
      default:
        throw new Error(`[split] Unsupported type ${type}`);
    }
    insertFn(childIndex + 1, after);

    return {before, after, betweenIndex:childIndex + 1, grandParent};
  }

  renderSplitItem(item, matchElement) {
  }

  renderTextItem(item, matchElement, callback) {
    this.cursor.before.appendText(item.value);
    matchElement.textElement = this.cursor.before.editAsText();

    if (callback) callback(item, matchElement);
    this.setMatchAttributes(matchElement, matchElement.start, matchElement.start + item.value.length - 1, this.attrs);
  }

  renderImageBlob(item, matchElement, blob) {
    const inlineImage = this.cursor.before.appendInlineImage(blob);

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

    if (item.linkUrl) {
      inlineImage.setLinkUrl(item.linkUrl);
    }

    return inlineImage;
  }

  renderImageItem(item, matchElement) {
    let blob = null;
    if (item.fileId) {
      const file = DriveApp.getFileById(item.fileId);
      blob = file.getAs('image/png');
    } else if (item.url) {
      const response = UrlFetchApp.fetch(item.url);
      blob = response.getBlob();
    }

    const inlineImage = this.renderImageBlob(item, matchElement, blob);
  }

  renderChartItem(item, matchElement) {
    const sheet = G.ss.getSheetByName(item.src);
    const charts = sheet.getCharts();
    const chart = charts[0];
    const blob = chart.getAs('image/png');
    
    const inlineImage = this.renderImageBlob(item, matchElement, blob);
  }

  renderTableItem(item, matchElement) {
    const namedRange = G.ss.getRangeByName(item.src);
    const dataRange = namedRange ? namedRange : G.ss.getRange(item.src);
    if (dataRange == null) return;

    const values = dataRange.getDisplayValues();
    const childIndex = this.cursor.parent.getChildIndex(this.cursor.before);

    if (!item.format) {
      this.cursor.parent.insertTable(childIndex, values);
    } else {
      const table = this.cursor.parent.insertTable(childIndex);
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
            this.setTextAttributes(text, run.start, run.endInclusive, run.attrs);
          });
        })
      });
    }
  }

  renderPageBreakItem(item, matchElement) {
    this.cursor.before.appendPageBreak();
  }

  renderLinkItem(item, matchElement) {
    const start = matchElement.start;
    this.renderTextItem(item, matchElement, (item, matchElement) => {
      const endInclusive = start + item.value.length - 1;
      matchElement.textElement.setLinkUrl(start, endInclusive, item.url);
    });
  }  
}

class SlidesRenderer {
  constructor() { 
    this.cursor = null;
  }

  render(items, matchElement) {
    this.attrs = this.getMatchAttributes(matchElement);
    //Logger.log(JSON.stringify(this.attrs));
    this.cursor = this.splitMatch(matchElement);

    for (const item of items) {
      Logger.log(JSON.stringify([matchElement.matched, item]));
      switch (item.kind) {
        case "text": {
          this.renderTextItem(item, matchElement);
          break;
        }

        case "image": {
          this.renderImageItem(item, matchElement);
          break;
        }

        case "replace_image": {
          this.renderReplaceImageItem(item, matchElement);
          break;
        }

        case "chart": {
          this.renderChartItem(item, matchElement);
          break;
        }

        case "table": {
          this.renderTableItem(item, matchElement);
          break;
        }

        case "pagebreak": {
          this.renderPageBreakItem(item, matchElement);
          break;
        }

        case "link": {
          this.renderLinkItem(item, matchElement);
          break;
        }

        case "split": {
          this.renderSplitItem(item, matchElement);
          break;
        }

        default:
          
      }
    }

    //this.cursor.after.merge();
  }

  getMatchFromRangeElement(rangeElement, context) {
    return {
      context,
      ...rangeElement
    }
  }

  getMatchAttributes(matchElement) {
    // get formatting attributes (into a clone)
    const runsAttrs = getTextRangeAttrs(matchElement.textRange, matchElement.slide.getColorScheme());
    const attrs = Object.assign({}, runsAttrs[0]);
    //Logger.log(JSON.stringify([matched, start, endInclusive]));
    //Logger.log(JSON.stringify(["BEFORE:", matchElement.textElement.getAttributes(matchElement.start)]));
    return attrs;
  }

  setTextAttributes(textElement, attrs, onlyNonNull=true) {  
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

  setMatchAttributes(matchElement, attrs, onlyNonNull=true) {  
    setTextRangeAttrs(matchElement.textRange, attrs, onlyNonNull);
  }

  splitMatch(matchElement) {
    const before = matchElement.textRange;
    before.setText("");
    return {before};
  }

  renderSplitItem(item, matchElement) {
  }

  renderTextItem(item, matchElement, callback) {
    this.cursor.before = this.cursor.before.appendText(item.value);

    if (callback) callback(item, matchElement);

    this.setMatchAttributes(matchElement, this.attrs.attrs);
  }

  renderImageBlob(item, matchElement, blob, replace=false) {
    const image = replace ? 
      matchElement.parent.replaceWithImage(blob)
      : matchElement.slide.insertImage(blob);

    if (item.width || item.height) {
      const w = image.getWidth();
      const h = image.getHeight();
      const ratio = w / h;
      if (item.width) {
        image.setWidth(item.width);
        if (!item.height) image.setHeight(item.width / ratio);
      }    
      if (item.height) {
        image.setHeight(item.height);
        if (!item.width) image.setWidth(item.height * ratio);
      }
    }
    if (item.left != null) {
      image.setLeft(item.left);
    }
    if (item.top != null) {
      image.setTop(item.top);
    }

    if (item.linkUrl) {
      image.setLinkUrl(item.linkUrl);
    }

    return image;
  }

  renderImageItem(item, matchElement, replace=false) {
    let blob = null;
    if (item.fileId) {
      const file = DriveApp.getFileById(item.fileId);
      blob = file.getAs('image/png');
    } else if (item.url) {
      const response = UrlFetchApp.fetch(item.url);
      blob = response.getBlob();
    }

    const image = this.renderImageBlob(item, matchElement, blob, replace);
  }

  renderReplaceImageItem(item, matchElement) {
    if (!item.width && !item.height) {
      item.width = matchElement.parent.getWidth();
      item.height = matchElement.parent.getHeight();
    }
    if (item.left == null) {
      item.left = matchElement.parent.getLeft();
    }
    if (item.top == null) {
      item.top = matchElement.parent.getTop();
    }
    this.renderImageItem(item, matchElement, true);
  }

  renderChartItem(item, matchElement, replace) {
    const sheet = G.ss.getSheetByName(item.src);
    const charts = sheet.getCharts();
    const chart = charts[0];
    const blob = chart.getAs('image/png');
    
    const image = this.renderImageBlob(item, matchElement, blob, replace);
  }

  renderTableItem(item, matchElement) {
    const namedRange = G.ss.getRangeByName(item.src);
    const dataRange = namedRange ? namedRange : G.ss.getRange(item.src);
    if (dataRange == null) return;

    const values = dataRange.getDisplayValues();
    const childIndex = this.cursor.parent.getChildIndex(this.cursor.before);

    if (!item.format) {
      this.cursor.parent.insertTable(childIndex, values);
    } else {
      const table = this.cursor.parent.insertTable(childIndex);
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
            this.setTextAttributes(text, run.start, run.endInclusive, run.attrs);
          });
        })
      });
    }
  }

  renderPageBreakItem(item, matchElement) {
    this.cursor.before.appendPageBreak();
  }

  renderLinkItem(item, matchElement) {
    this.renderTextItem(item, matchElement, (item, matchElement) => {
      //const textRange = matchElement.textElement.getRange(matchElement.start, matchElement.end);
      Logger.log(matchElement.textRange.asString());
      this.cursor.before.getTextStyle().setLinkUrl(item.url);
    });
  }  
}

