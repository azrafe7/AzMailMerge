function uiAvailable() {
  try {
    SpreadsheetApp.getUi();
    return true;
  } catch (error) {
    return false;
  }
}

function getSheetColumnsMap(sheetOrData) {
  let columnsMap = new Map();
  const headers = Array.isArray(sheetOrData) ? sheetOrData[0] : sheetOrData.getRange("1:1").getValues()[0];
  headers.forEach((name) => columnsMap.set(name, headers.indexOf(name)));
  return columnsMap;
}

/**
 * Converts row/col numbers to A1 notation
 * @param {number} row - Row number (1-indexed)
 * @param {number} col - Column number (1-indexed)
 * @returns {Array<string>} A1 notation (e.g., ['A', '1'], ['B', '5'])
 */
function toA1Notation(row, col) {
  let colStr = '';
  let tempCol = col;
  
  while (tempCol > 0) {
    tempCol--;
    colStr = String.fromCharCode(65 + (tempCol % 26)) + colStr;
    tempCol = Math.floor(tempCol / 26);
  }
  
  return [colStr, row];
}

/**
 * Parses A1 notation to row/col numbers
 * @param {string} a1 - A1 notation (e.g., 'A1', 'B5', 'AA10')
 * @returns {Object} Object with row and col properties
 */
function parseA1Notation(a1) {
  const match = a1.match(/^\$?([A-Z]+)?\$?(\d+)?$/);
  if (!match) throw new Error(`Invalid A1 notation: ${a1}`);
  
  const colStr = match[1] || null;
  const row = parseInt(match[2]) || null;
  
  let col = null;
  for (let i = 0; colStr != null && i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }
  
  return {row, col};
}

function getFileNameFromId(id) {
  const file = DriveApp.getFileById(id);
  const name = file.getName();
  return name;
}

function transpose(matrix) {
  return matrix[0].map((_, colIndex) => matrix.map(row => row[colIndex]));
}

function toBool(value) {
  const strValue = String(value).toLowerCase();
  return strValue === 'true' || strValue === 'yes' || strValue === '1' ? true : false;
}

function* iterateChildrenOf(section) {
  let numElements = section.getNumChildren();
  for (let i = 0; i < numElements; i++) {
    let child = section.getChild(i);
    let type = child.getType();
    yield { child, type, index:i };
  }
}  

function appendSectionTo(sourceSection, targetSection) {
  let numElements = sourceSection.getNumChildren();
  
  // since a table cannot be the last element, the Google API auto-appends a paragraph when that condition is met.
  // we collect them here, so they can be removed at the end
  const targetSectionType = targetSection.getType();
  let paraToRemove = [];

  for (let i = 0; i < numElements; i++) {
    let child = sourceSection.getChild(i);
    let copy = child.copy();
    let type = child.getType();

    Logger.log(`Appending type ${type}`);
    switch (type) {
      case DocumentApp.ElementType.PARAGRAPH: {
        targetSection.appendParagraph(copy);
        break;
      }
      
      case DocumentApp.ElementType.TABLE: {
        targetSection.appendTable(copy);
        const nextSibling = copy.getNextSibling();
        if (nextSibling) paraToRemove.push(nextSibling);
        break;
      }
      
      case DocumentApp.ElementType.LIST_ITEM: {
        targetSection.appendListItem(copy);
        copy.setGlyphType(child.getGlyphType());
        break;
      }

      /*
      case DocumentApp.ElementType.INLINE_IMAGE:
        targetSection.appendImage(copy.asInlineImage());
        break;
      
      case DocumentApp.ElementType.HORIZONTAL_RULE:
        targetSection.appendTable(copy.asHorizontalRule());
        break;
      */
      
      default:
        throw new Error(`Unsupported type ${type}`);
        break;
    }
  }

  //Logger.log(`Removed ${paraToRemove.length} para after tables`);
  if (paraToRemove.length > 0) paraToRemove.forEach((p) => p.removeFromParent())
}

function appendDocTo(sourceDoc, targetDoc, { header = false, footer = false }) {
  const sourceBody = sourceDoc.getBody();
  const targetBody = targetDoc.getBody();

  appendSectionTo(sourceBody, targetBody);

  // get header/footer from first tab
  if (header || footer) {
    const sourceTabs = sourceDoc.getTabs();
    const sourceFirstTab = sourceTabs[0].asDocumentTab();
    const targetTabs = targetDoc.getTabs();
    const targetFirstTab = targetTabs[0].asDocumentTab();

    if (header) {
      const sourceHeader = sourceFirstTab.getHeader();
      if (sourceHeader) {
        Logger.log(`Trying to append type ${sourceHeader.getType()}`);
        let targetHeader = targetFirstTab.getHeader() ?? targetFirstTab.addHeader();
        appendSectionTo(sourceHeader, targetHeader);
      }
    }
    if (footer) {
      const sourceFooter = sourceFirstTab.getFooter();
      if (sourceFooter) {
        Logger.log(`Trying to append type ${sourceFooter.getType()}`);
        let targetFooter = targetFirstTab.getFooter() ?? targetFirstTab.addFooter();
        appendSectionTo(sourceFooter, targetFooter);
      }
    }
  }
}

function convertDocToPdf(pdfName, doc, outputFolder) {
  let blob = null;
  const file = DriveApp.getFileById(doc.getId());
  if (file.getMimeType() === DOC_TYPE.DOCUMENT) {
    blob = doc.getAs('application/pdf');
  } else {
    blob = file.getBlob();
  }
  const pdfFile = outputFolder.createFile(blob);
  pdfFile.setName(pdfName);
  return pdfFile;
}

function findPresentationPlaceholders(presentation, pattern) {
  let matches = [];
  const slides = presentation.getSlides();
  for (let slideIdx=0; slideIdx<slides.length; slideIdx++) {
    const slide = slides[slideIdx];
    
    let textNodes = [];
    const { tree, maxStepsReached } = getPageElementStructure(slide.getPageElements(), (node) => {
      if (node.typedElement.getText) textNodes.push(node);
    });

    for (let textNode of textNodes) {
      const textElement = textNode.typedElement.getText();
      let foundMatches = textElement.find(pattern);
      if (foundMatches.length > 0) {
        foundMatches.forEach((tr) => {
          const start = tr.getStartIndex();
          const end = tr.getEndIndex();
          Logger.log(tr.asString());
          matches.push({
            slideIndex: slideIdx,
            slide,
            textElement: textElement,
            parent: textNode.typedElement,
            textRange: tr,
            start,
            end,
            text: textElement.asString(),
            matched: tr.asString()
          });
        });
      }
    }
  }
  return matches;
}

function findPlaceholders(body, pattern) {
  return Array.from(findAllText(body, pattern));
}

/**
 * Generator that finds all the entries when searching.
 * @param {DocumentApp.Body} body Body of the file.
 * @param {string} text Text to find.
 * @returns {Iterator.<DocumentApp.RangeElement>}
 */
function* findAllText(body, text) {
  let entry = body.findText(text);
  while(entry != null) {
    yield entry;
    entry = body.findText(text, entry);
  }
}

function getRangeRuns(range) {
  const numberFormats = range.getNumberFormats();
  
  // temporarily set all formats to String (otherwise cells formatted as numbers report empty richTextValues)
  const stringFormats = numberFormats.map(row => row.map(cell => '@'));
  range.setNumberFormats(stringFormats);

  const richTextValues = range.getRichTextValues();
  const backgrounds = range.getBackgrounds();

  let docCells = [];
  for (let row = 0; row < richTextValues.length; row++) {
    let docRow = [];
    for (let col = 0; col < richTextValues[0].length; col++) {
      let docCell = [];
      const richTextValue = richTextValues[row][col];
      const background = backgrounds[row][col];
      for (let run of richTextValue.getRuns()) {
        const text = run.getText();
        let docRun = { text: text };
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
        attrs["BOLD"] = null;

        let start = run.getStartIndex();
        let endInclusive = run.getEndIndex() - 1;
        if (start < endInclusive) {
          let textStyle = run.getTextStyle(start, endInclusive);
          let fontFamily = textStyle.getFontFamily();
          let fontSize = textStyle.getFontSize();
          let fontColor = textStyle.getForegroundColorObject().asRgbColor().asHexString();
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
          attrs["BOLD"] = bold;
        }

        docRun.start = start;
        docRun.endInclusive = endInclusive;
        docRun["attrs"] = attrs;
        docCell.push(docRun);
      }
      docRow.push(docCell);
    }
    docCells.push(docRow);
  }

  // restore formats
  range.setNumberFormats(numberFormats);

  Logger.log(JSON.stringify(docCells));
  return docCells;
}

function getPageElementStructure(section, callback) {
  if (Array.isArray(section)) {
    const array = section.slice();
    section = {
      getChildren: () => array,
      getPageElementType: () => "DUMMY_ROOT",
    }
  }
  let tree = {};
  let parentsMap = {};
  let id = 0;
  let stack = [[{elementId:id, element:section}]];
  let step = 0;
  let level = 0;
  const MAX_STEPS = 500;
  let maxStepsReached = false;
  while (stack.length > 0 && !maxStepsReached) {
    const stackElements = stack.pop();
    for (let stackElement of stackElements) {
      const { elementId, element } = stackElement;
      const { typed: typedElement, type } = getAsTypedPageElement(element);
      Logger.log(type);
      
      let getChildren = typedElement.getChildren ? typedElement.getChildren : () => [];
      switch (type) {
        case SlidesApp.PageElementType.GROUP: {
          getChildren = typedElement.getChildren;
          break;
        }

        case SlidesApp.PageElementType.TABLE: {
          const numRows = typedElement.getNumRows();
          const numColumns = typedElement.getNumColumns();
          getChildren = () => {
            let cells = [];
            for (let r=0; r<numRows; r++) {
              for (let c=0; c<numColumns; c++) {
                let cell = typedElement.getCell(r, c);
                cell["getPageElementType"] = () => "CELL";
                cells.push(cell);
              }
            }
            return cells;
          };
          break;
        }
      }

      const children = getChildren();
      const numChildren = children.length;
      const text = typedElement.getText ? typedElement.getText().asString() : null;

      let node = { elementId, type, level:null, children: [], text };
      //Logger.log(JSON.stringify(node));

      let parent = null;
      if (step == 0) {
        tree = node;
        node.level = 0;
      } else {
        parent = parentsMap[node.elementId];
        parent.children.push(node);
        node.level = parent.level + 1;
        level = Math.max(level, node.level);
      }
      delete node.elementId;
      if (node.text == null) delete node.text;

      if (callback) callback({typedElement, parent, ...node});

      let stackChildren = [];
      for (let childIndex = 0; childIndex < numChildren; childIndex++) {
        id++;
        const child = children[childIndex];
        stackChildren.push({ elementId:id, element:child });
        parentsMap[id] = node;
      }
      if (stackChildren.length > 0) {
        stack.push(stackChildren);
      }
    }
    step++;
    maxStepsReached = step >= MAX_STEPS;
  }

  Logger.log(`numElements:${id + 1} steps:${step} maxLevel:${level}`);
  return { tree, maxStepsReached };
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

function getSectionStructure(section) {
  let tree = {};
  let parentsMap = {};
  let id = 0;
  let stack = [[{elementId:id, element:section}]];
  let step = 0;
  let level = 0;
  const MAX_STEPS = 500;
  let maxStepsReached = false;
  while (stack.length > 0 && !maxStepsReached) {
    const stackElements = stack.pop();
    for (let stackElement of stackElements) {
      const { elementId, element } = stackElement;
      const numChildren = element.getNumChildren ? element.getNumChildren() : 0;
      const type = element.getType ? element.getType() : null;
      const text = type === DocumentApp.ElementType.TEXT ? element.getText() : null;

      let node = { elementId, type, level:null, children: [], text };
      //Logger.log(JSON.stringify(node));

      if (step == 0) {
        tree = node;
        node.level = 0;
      } else {
        const parent = parentsMap[node.elementId];
        parent.children.push(node);
        node.level = parent.level + 1;
        level = Math.max(level, node.level);
      }
      delete node.elementId;
      if (node.text == null) delete node.text;

      let children = [];
      for (let childIndex = 0; childIndex < numChildren; childIndex++) {
        id++;
        const child = element.getChild(childIndex);
        children.push({ elementId:id, element:child });
        parentsMap[id] = node;
      }
      if (children.length > 0) {
        stack.push(children);
      }
    }
    step++;
    maxStepsReached = step >= MAX_STEPS;
  }

  Logger.log(`numElements:${id + 1} steps:${step} maxLevel:${level}`);
  return { tree, maxStepsReached };
}

function getAsTypedPageElement(element, ignoreErrors=true) {
  const type = (ignoreErrors && !element.getPageElementType) ? "DUMMY" : element.getPageElementType();
  let typed = ignoreErrors ? element : null;

  switch (type) {
    case SlidesApp.PageElementType.SHAPE: {
      typed = element.asShape();
      break;
    }	
    case SlidesApp.PageElementType.IMAGE: {
      typed = element.asImage();
      break;
    }	
    case SlidesApp.PageElementType.VIDEO: {
      typed = element.asVideo();
      break;
    }	
    case SlidesApp.PageElementType.TABLE: {
      typed = element.asTable();
      break;
    }	
    case SlidesApp.PageElementType.GROUP: {
      typed = element.asGroup();
      break;
    }	
    case SlidesApp.PageElementType.LINE: {
      typed = element.asLine();
      break;
    }	
    case SlidesApp.PageElementType.WORD_ART: {
      typed = element.asWordArt();
      break;
    }	
    case SlidesApp.PageElementType.SHEETS_CHART: {
      typed = element.asSheetsChart();
      break;
    }	
    case SlidesApp.PageElementType.SPEAKER_SPOTLIGHT: {
      typed = element.asSpeakerSpotlight();
      break;
    }	
    case SlidesApp.PageElementType.UNSUPPORTED: 
    default:
      if (!ignoreErrors) throw new Error("UNSUPPORTED ELEMENT TYPE");
  }

  return { typed, type };
}

/**
 * Returns `element` as properly typed.
 * `typedAs` can be:
 *  - null: type will be what's returned by `element.getType()`
 *  - one of DocumentApp.ElementType
 *  - another element: type will be what's returned by `typedAs.getType()`
 */
function getAsTyped(element, typedAs=null) {
  const type = typedAs == null ? element.getType() : (typeof typedAs.getType === "function" ? typedAs.getType() : typedAs);
  let typed = null;

  switch (type) {
    case DocumentApp.ElementType.BODY_SECTION: {
      typed = element.asBody();
      break;
    }
    case DocumentApp.ElementType.COMMENT_SECTION: {
      typed = element.asCommentSection();
      break;
    }
    case DocumentApp.ElementType.DATE: {D
      typed = element.asate();
      break;
    }
    case DocumentApp.ElementType.EQUATION: {
      typed = element.asEquation();
      break;
    }
    case DocumentApp.ElementType.EQUATION_FUNCTION: {
      typed = element.asEquationFunction();
      break;
    }
    case DocumentApp.ElementType.EQUATION_FUNCTION_ARGUMENT_SEPARATOR: {
      typed = element.asEquationFunctionArgumentSeparator();
      break;
    }
    case DocumentApp.ElementType.EQUATION_SYMBOL: {
      typed = element.asEquationSymbol();
      break;
    }
    case DocumentApp.ElementType.RICH_LINK: {
      typed = element.asRichLink();
      break;
    }
    case DocumentApp.ElementType.FOOTER_SECTION: {
      typed = element.asFooterSection();
      break;
    }
    case DocumentApp.ElementType.FOOTNOTE: {
      typed = element.asFootnote();
      break;
    }
    case DocumentApp.ElementType.FOOTNOTE_SECTION: {
      typed = element.asFootnoteSection();
      break;
    }
    case DocumentApp.ElementType.HEADER_SECTION: {
      typed = element.asHeaderSection();
      break;
    }
    case DocumentApp.ElementType.HORIZONTAL_RULE: {
      typed = element.asHorizontalRule();
      break;
    }
    case DocumentApp.ElementType.INLINE_DRAWING: {
      typed = element.asInlineDrawing();
      break;
    }
    case DocumentApp.ElementType.INLINE_IMAGE: {
      typed = element.asInlineImage();
      break;
    }
    case DocumentApp.ElementType.LIST_ITEM: {
      typed = element.asListItem();
      break;
    }
    case DocumentApp.ElementType.PAGE_BREAK: {
      typed = element.asPageBreak();
      break;
    }
    case DocumentApp.ElementType.PARAGRAPH: {
      typed = element.asParagraph();
      break;
    }
    case DocumentApp.ElementType.PERSON: {
      typed = element.asPerson();
      break;
    }
    case DocumentApp.ElementType.TABLE: {
      typed = element.asTable();
      break;
    }
    case DocumentApp.ElementType.TABLE_CELL: {
      typed = element.asTableCell();
      break;
    }
    case DocumentApp.ElementType.TABLE_OF_CONTENTS: {
      typed = element.asTableOfContents();
      break;
    }
    case DocumentApp.ElementType.TABLE_ROW: {
      typed = element.asTableRow();
      break;
    }
    case DocumentApp.ElementType.TEXT: {
      typed = element.asText();
      break;
    }
    case DocumentApp.ElementType.UNSUPPORTED:
    default:
      throw new Error("UNSUPPORTED ELEMENT TYPE");
  }

  return typed;
}

function logToSheet(rowData) {
  const sheet = G.ss.getSheetByName(LOG_SHEET);
  sheet.appendRow(Array.isArray(rowData) ? rowData : [rowData]);
}

