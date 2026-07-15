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
  constructor() { }

  render(items, matchElement) {
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

  replaceMatchText(matchElement, text) {
    // delete text and insert replacement
    if (matchElement.endInclusive > matchElement.start) { // as deleteText(10, 10) will still erase one char
      matchElement.textElement.deleteText(matchElement.start, matchElement.endInclusive);
    }
    matchElement.textElement.insertText(matchElement.start, text);
  }

  splitElement(element, start, endInclusive) {
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

  renderSplitItem(item, matchElement) {
    const parent = matchElement.rangeElement.getElement().getParent();
    const splitted = this.splitElement(parent, matchElement.start, matchElement.endInclusive);
  }

  renderTextItem(item, matchElement, callback) {
    const attrs = this.getMatchAttributes(matchElement);
    this.replaceMatchText(matchElement, item.value);
    if (callback) callback(item, matchElement);
    this.setMatchAttributes(matchElement, matchElement.start, matchElement.start + item.value.length - 1, attrs);
    // update the position
    matchElement.start = matchElement.start + item.value.length;
    matchElement.endInclusive = matchElement.start;
  }

  renderImageBlob(item, matchElement, blob) {
    const parent = matchElement.rangeElement.getElement().getParent();

    const splitted = this.splitElement(parent, matchElement.start, matchElement.endInclusive);
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
    this.replaceMatchText(matchElement, "");
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
      const rangeRuns = this.getRangeRuns(dataRange);
      
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
    this.replaceMatchText(matchElement, "");
    const p = matchElement.rangeElement.getElement().getParent();
    p.insertPageBreak(0);
  }

  renderLinkItem(item, matchElement) {
    const start = matchElement.start;
    this.renderTextItem(item, matchElement, (item, matchElement) => {
      const endInclusive = start + item.value.length - 1;
      matchElement.textElement.setLinkUrl(start, endInclusive, item.url);
    });
  }  
}

