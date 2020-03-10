/**
 * Created by laosong on 2020/2/18.
 * create htXml from docx
 */
'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.parseDocx = parseDocx;
exports.createHT = createHT;
exports.Docx = void 0;

const _ = require('lodash/core');

const xmlDOM = require('xmldom');

const DOMParser = xmlDOM.DOMParser;

const elem = require('./elem');

const helper = require('./helper');

function getChildrenByTagName(parent, tagName) {
  let matched = [];

  for (let i = 0; i < parent.childNodes.length; i++) {
    const childNode = parent.childNodes[i];
    if (!childNode) continue;
    if (childNode.nodeType !== 1) continue;
    if (childNode.tagName !== tagName) continue;
    matched.push(childNode);
  }

  return matched;
}

function firstElementByTagName(parent, tagName) {
  let elem = null;

  for (let i = 0; i < parent.childNodes.length; i++) {
    const childNode = parent.childNodes[i];
    if (!childNode) continue;
    if (childNode.nodeType !== 1) continue;

    if (childNode.tagName === tagName) {
      elem = childNode;
    }

    if (elem) {
      return elem;
    }

    elem = firstElementByTagName(childNode, tagName);

    if (elem) {
      return elem;
    }
  }

  return elem;
}

class Settings {
  constructor(xml) {
    this.xml = xml;
    this.defaultTabStop = null;
    this.zoom = null;
  }

  static parse(settingsXml) {
    if (!settingsXml) {
      return null;
    }

    let settings = new Settings(settingsXml);
    const dom = new DOMParser().parseFromString(settingsXml, 'text/xml'); //let nodes = getChildrenByTagName(dom.documentElement, 'w:defaultTabStop');

    let node = firstElementByTagName(dom.documentElement, 'w:defaultTabStop');

    if (node) {
      settings.defaultTabStop = parseInt(node.getAttribute('w:val'));
    }

    let nodes = getChildrenByTagName(dom.documentElement, 'w:zoom');
    return settings;
  }

}

class Relations {
  constructor(relsXml) {
    this.xml = relsXml;
    this.rels = [];
  }

  static parse(relsXml) {
    if (!relsXml) {
      return null;
    }

    let relations = new Relations(relsXml);
    const dom = new DOMParser().parseFromString(relsXml, 'text/xml');
    let relElements = getChildrenByTagName(dom.documentElement, 'Relationship');

    for (const relElement of relElements) {
      let relation = {};
      relation.Id = relElement.getAttribute('Id');
      relation.Target = relElement.getAttribute('Target');
      relations.rels.push(relation);
    }

    return relations;
  }

  getRelByRId(rId) {
    let rel = _.find(this.rels, {
      Id: rId
    });

    if (!rel) {}

    return rel;
  }

}

class Numbering {
  constructor(numberingXml) {
    this.xml = numberingXml;
    this.abstractNums = [];
    this.nums = [];
    this.numbering = {};
  }

  static parseAbstractNum(abNumElem) {
    let abstractNum = {};
    abstractNum.id = abNumElem.getAttribute('w:abstractNumId');
    let levels = [];
    let lvlElements = getChildrenByTagName(abNumElem, 'w:lvl');

    for (const lvlElement of lvlElements) {
      let level = {};
      level.ilvl = lvlElement.getAttribute('w:ilvl');

      for (let i = 0; i < lvlElement.childNodes.length; i++) {
        const childNode = lvlElement.childNodes[i];
        if (!childNode) continue;
        if (childNode.nodeType !== 1) continue;
        let val = childNode.getAttribute('w:val');

        switch (childNode.tagName) {
          case 'w:start':
            level.start = parseInt(val);
            break;

          case 'w:numFmt':
            level.numFmt = val;
            break;

          case 'w:lvlText':
            level.lvlText = val;
            break;

          case 'w:lvlJc':
            level.lvlJc = val;
            break;
        }
      }

      let indElem = firstElementByTagName(lvlElement, 'w:ind');

      if (indElem) {
        let left = indElem.getAttribute('w:left');
        let hanging = indElem.getAttribute('w:hanging');

        if (left) {
          level.indentLeft = parseInt(left);
        }

        if (hanging) {
          level.indentLeft -= parseInt(hanging);
        }
      }

      levels.push(level);
    }

    abstractNum.levels = levels;
    return abstractNum;
  }

  static parseNum(numElem) {
    let num = {};
    num.id = numElem.getAttribute('w:numId');

    for (let i = 0; i < numElem.childNodes.length; i++) {
      const childNode = numElem.childNodes[i];
      if (!childNode) continue;
      if (childNode.nodeType !== 1) continue;
      let val = childNode.getAttribute('w:val');

      switch (childNode.tagName) {
        case 'w:abstractNumId':
          num.abstractNumId = val;
          break;
      }
    }

    return num;
  }

  static parse(numberingXml) {
    if (!numberingXml) {
      return null;
    }

    let numbering = new Numbering(numberingXml);
    const dom = new DOMParser().parseFromString(numberingXml, 'text/xml');
    let abNumElements = getChildrenByTagName(dom.documentElement, 'w:abstractNum');

    for (const abNumElement of abNumElements) {
      const abstractNum = Numbering.parseAbstractNum(abNumElement);

      if (!abstractNum) {
        console.warn('Numbers.parseAbstractNum return null', abNumElement);
        continue;
      }

      numbering.abstractNums.push(abstractNum);
    }

    let numElements = getChildrenByTagName(dom.documentElement, 'w:num');

    for (const numElement of numElements) {
      const num = Numbering.parseNum(numElement);

      if (!num) {
        console.warn('Numbers.parseNum return null', numElement);
        continue;
      }

      numbering.nums.push(num);
    }

    for (let num of numbering.nums) {
      const abstractNum = _.find(numbering.abstractNums, {
        id: num.abstractNumId
      });

      if (!abstractNum) {
        console.warn('num', num, 'no abstractNum');
      }

      num.abstractNum = abstractNum;
    }

    return numbering;
  }

  setNumLevel(numId, ilvl) {
    const num = _.find(this.nums, {
      id: numId
    });

    if (!num) {
      return null;
    }

    if (!num.abstractNum) {
      return null;
    }

    const level = _.find(num.abstractNum.levels, {
      ilvl: ilvl
    });

    if (!level) {
      return null;
    }

    let newNumObj = {
      number: 1,
      text: null,
      indentLeft: 0
    };
    let numberingOfNum = this.numbering[numId];

    if (!numberingOfNum) {
      numberingOfNum = {};
      this.numbering[numId] = numberingOfNum;
    }

    let curNo = 0;
    let numberingOfNumLevel = numberingOfNum[ilvl];

    if (!Array.isArray(numberingOfNumLevel)) {
      numberingOfNumLevel = [0];
      numberingOfNum[ilvl] = numberingOfNumLevel;
    }

    curNo = _.last(numberingOfNumLevel);
    curNo += 1;
    numberingOfNumLevel.push(curNo);

    for (let k in numberingOfNum) {
      if (k > ilvl) {
        numberingOfNum[k].push(0);
      }
    }

    newNumObj.number = curNo + (level.start - 1 || 0);
    newNumObj.indentLeft = level.indentLeft;
    let fmtNumber = null;

    switch (level.numFmt) {
      case 'decimal':
        fmtNumber = String(newNumObj.number);
        break;

      case 'lowerLetter':
        fmtNumber = helper.lowerLetterOfNum(newNumObj.number);
        break;

      case 'upperLetter':
        fmtNumber = helper.upperLetterOfNum(newNumObj.number);
        break;

      case 'lowerRoman':
        fmtNumber = helper.lowerRomanOfNum(newNumObj.number);
        break;

      case 'upperRoman':
        fmtNumber = helper.upperRomanOfNum(newNumObj.number);
        break;

      default:
        fmtNumber = helper.simpleChineseOfNum(newNumObj.number);
        break;
    }

    if (level.lvlText) {
      newNumObj.text = level.lvlText.replace(/%\d+/g, fmtNumber);
    } else {
      newNumObj.text = fmtNumber;
    }

    return newNumObj;
  }

}

class Styles {
  constructor(stylesXml) {
    this.xml = stylesXml;
    this.defaultFont = null;
    this.defaultFontSize = 0;
    this.paragraphStyles = [];
    this.characterStyles = [];
    this.tableStyles = [];
    this.numberingStyles = [];
  }

  static parseParagraphStyle(styleElem) {
    let paragraphStyle = {};
    paragraphStyle.id = styleElem.getAttribute('w:styleId');

    for (let i = 0; i < styleElem.childNodes.length; i++) {
      const childNode = styleElem.childNodes[i];
      if (!childNode) continue;
      if (childNode.nodeType !== 1) continue;
      let val = childNode.getAttribute('w:val');

      switch (childNode.tagName) {
        case 'w:name':
          paragraphStyle.name = val;
          break;

        case 'w:basedOn':
          paragraphStyle.basedOn = val;
          break;
      }
    }

    return paragraphStyle;
  }

  static parseCharacterStyle(styleElem) {//not implement
  }

  static parseTableStyle(styleElem) {//not implement
  }

  static parseNumberingStyle(styleElem) {//not implement
  }

  static parse(stylesXml) {
    if (!stylesXml) {
      return null;
    }

    let styles = new Styles(stylesXml);
    const dom = new DOMParser().parseFromString(stylesXml, 'text/xml');
    let rPrDefaultElem = firstElementByTagName(dom.documentElement, 'w:rPrDefault');

    if (rPrDefaultElem) {
      let rPrElem = firstElementByTagName(rPrDefaultElem, 'w:rPr');
      let szCsElem = null;

      if (rPrElem && (szCsElem = firstElementByTagName(rPrElem, 'w:szCs'))) {
        styles.defaultFontSize = parseInt(szCsElem.getAttribute('w:val'));
      }
    }

    let styleElements = getChildrenByTagName(dom.documentElement, 'w:style');

    for (const styleElement of styleElements) {
      const type = styleElement.getAttribute('w:type');

      switch (type) {
        case 'paragraph':
          const paragraphStyle = Styles.parseParagraphStyle(styleElement);

          if (!paragraphStyle) {
            break;
          }

          styles.paragraphStyles.push(paragraphStyle);
          break;

        case 'character':
          const characterStyle = Styles.parseCharacterStyle(styleElement);

          if (!characterStyle) {
            break;
          }

          styles.characterStyles.push(characterStyle);
          break;

        case 'table':
          const tableStyle = Styles.parseTableStyle(styleElement);

          if (!tableStyle) {
            break;
          }

          styles.tableStyles.push(tableStyle);
          break;

        case 'numbering':
          const numberingStyle = Styles.parseNumberingStyle(styleElement);

          if (!numberingStyle) {
            break;
          }

          styles.numberingStyles.push(numberingStyle);
          break;
      }
    }

    return styles;
  }

  getParagraphStyle(styleId) {
    let paragraphStyle = _.find(this.paragraphStyles, {
      id: styleId
    });

    if (!paragraphStyle) {}

    return paragraphStyle;
  }

}

function addChild2Parent(parent, child) {
  if (!child) {
    return;
  }

  if (!parent || !parent.addChild) {
    return;
  }

  parent.addChild(child);
}

function parseText(parent, tElem) {
  if (!tElem || !tElem.getAttribute) {
    return null;
  }

  let run = parent;
  let paragraph = run.parent;

  if (!run.children) {
    let lastRun = _.last(paragraph.children);

    if (lastRun && lastRun.children && lastRun.children.length === 1) {
      let tRun = _.last(lastRun.children);

      if (tRun.textContent && _.isEqual(lastRun.properties, run.properties)) {
        tRun.textContent += tElem.textContent;
        return null;
      }
    }
  }

  let textRun = new elem.TextRun(parent, null, tElem.textContent);
  addChild2Parent(parent, textRun);
  return textRun;
}

function parseDrawing(parent, drawingElem) {
  if (!drawingElem || !drawingElem.getAttribute) {
    return null;
  }

  let drawingPr = {};
  let floatPr = {};
  let wpRoot = null;
  let wpInline = firstElementByTagName(drawingElem, 'wp:inline');

  if (wpInline) {
    drawingPr[elem.DrawingPr.floatPr] = null;
    wpRoot = wpInline;
  } else {
    let wpAnchor = firstElementByTagName(drawingElem, 'wp:anchor');

    if (wpAnchor) {
      for (let i = 0; i < wpAnchor.childNodes.length; i++) {
        const childNode = wpAnchor.childNodes[i];
        if (!childNode) continue;
        if (childNode.nodeType !== 1) continue;

        switch (childNode.tagName) {
          case 'wp:positionH':
            floatPr[elem.DrawingFloatPr.horizontalPosition] = childNode.getAttribute('relativeFrom');
            break;

          case 'wp:positionV':
            floatPr[elem.DrawingFloatPr.verticalPosition] = childNode.getAttribute('relativeFrom');
            break;

          case 'wp:wrapNone':
            floatPr[elem.DrawingFloatPr.textWrap] = elem.TextWrapMode.wrapNone;
            break;

          case 'wp:wrapSquare':
            floatPr[elem.DrawingFloatPr.textWrap] = elem.TextWrapMode.wrapSquare;
            break;

          case 'wp:wrapThrough':
            floatPr[elem.DrawingFloatPr.textWrap] = elem.TextWrapMode.wrapThrough;
            break;

          case 'wp:wrapTight':
            floatPr[elem.DrawingFloatPr.textWrap] = elem.TextWrapMode.wrapTight;
            break;

          case 'wp:wrapTopAndBottom':
            floatPr[elem.DrawingFloatPr.textWrap] = elem.TextWrapMode.wrapTopAndBottom;
            break;
        }
      }

      drawingPr[elem.DrawingPr.floatPr] = floatPr;
    }

    wpRoot = wpAnchor;
  }

  if (!wpRoot) {
    return null;
  }

  let wpExtent = firstElementByTagName(wpRoot, 'wp:extent');

  if (wpExtent) {
    drawingPr[elem.DrawingPr.cx] = parseInt(wpExtent.getAttribute('cx'));
    drawingPr[elem.DrawingPr.cy] = parseInt(wpExtent.getAttribute('cy'));
  }

  let aGraphic = firstElementByTagName(wpRoot, 'a:graphic');

  if (!aGraphic) {
    return null;
  }

  let aBlip = firstElementByTagName(aGraphic, 'a:blip');

  if (!aBlip) {
    return null;
  }

  drawingPr[elem.DrawingPr.imgRId] = aBlip.getAttribute('r:embed');
  let drawingRun = new elem.DrawingRun(parent, drawingPr);
  addChild2Parent(parent, drawingRun);
  return drawingRun;
}

function parseObject(parent, objectElem) {
  if (!objectElem || !objectElem.getAttribute) {
    return null;
  }

  let objectPr = {};
  let vShape = firstElementByTagName(objectElem, 'v:shape');

  if (!vShape) {
    return null;
  }

  objectPr[elem.ObjectPr.shapeId] = vShape.getAttribute('id');
  objectPr[elem.ObjectPr.shapeStyle] = vShape.getAttribute('style');
  let vImageData = firstElementByTagName(vShape, 'v:imagedata');

  if (!vImageData) {
    return null;
  }

  objectPr[elem.ObjectPr.imgRId] = vImageData.getAttribute('r:id');
  let vOLEObject = firstElementByTagName(objectElem, 'o:OLEObject');

  if (!vOLEObject) {
    return null;
  }

  objectPr[elem.ObjectPr.oleProgId] = vOLEObject.getAttribute('ProgID');
  objectPr[elem.ObjectPr.oleShapeId] = vOLEObject.getAttribute('ShapeID');
  objectPr[elem.ObjectPr.oleRId] = vOLEObject.getAttribute('r:id');
  let objectRun = new elem.ObjectRun(parent, objectPr);
  addChild2Parent(parent, objectRun);
  return objectRun;
}

function parseSpacingPr(spacingPrElem) {
  if (!spacingPrElem || !spacingPrElem.getAttribute) {
    return null;
  }

  let spacingPr = {};
  spacingPr[elem.SpacingPr.before] = parseInt(spacingPrElem.getAttribute('w:before'));
  spacingPr[elem.SpacingPr.after] = parseInt(spacingPrElem.getAttribute('w:after'));
  spacingPr[elem.SpacingPr.line] = parseInt(spacingPrElem.getAttribute('w:line'));
  return spacingPr;
}

function parseBorderPr(borderElem) {
  if (!borderElem || !borderElem.getAttribute) {
    return null;
  }

  let borderPr = {};
  borderPr[elem.BorderPr.style] = borderElem.getAttribute('w:val');
  borderPr[elem.BorderPr.width] = parseInt(borderElem.getAttribute('w:sz'));
  borderPr[elem.BorderPr.color] = borderElem.getAttribute('w:color');
  return borderPr;
}

function parseRunPR(runPrElem) {
  if (!runPrElem || !runPrElem.hasChildNodes) {
    return null;
  }

  let rPrObj = {};
  let textPr = {};

  for (let i = 0; i < runPrElem.childNodes.length; i++) {
    const childNode = runPrElem.childNodes[i];
    if (!childNode) continue;
    if (childNode.nodeType !== 1) continue;

    switch (childNode.tagName) {
      case 'w:pStyle':
        rPrObj[elem.RunPr.styleId] = childNode.getAttribute('w:val');
        break;

      case 'w:bdr':
        rPrObj[elem.RunPr.border] = parseBorderPr(childNode);
        break;

      case 'w:highlight':
        rPrObj[elem.RunPr.highlight] = childNode.getAttribute('w:val');
        break;

      case 'w:shd':
        rPrObj[elem.RunPr.shdFillColor] = childNode.getAttribute('w:fill');
        break;

      case 'w:rFonts':
        let fontName = childNode.getAttribute('w:eastAsia');

        if (fontName) {
          rPrObj[elem.RunPr.fontName] = fontName;
        }

        let fontName2 = childNode.getAttribute('w:ascii');

        if (fontName2) {
          if (!fontName || fontName2 === 'Symbol') {
            rPrObj[elem.RunPr.fontName] = fontName2;
          }
        }

        break;

      case 'w:sz':
      case 'w:szCs':
        let fontSize = childNode.getAttribute('w:val');

        if (fontSize) {
          rPrObj[elem.RunPr.fontSize] = parseInt(fontSize);
        }

        break;

      case 'w:color':
        rPrObj[elem.RunPr.textColor] = childNode.getAttribute('w:val');
        break;

      case 'w:b':
        textPr[elem.TextPr.bold] = true;
        break;

      case 'w:i':
        textPr[elem.TextPr.italics] = true;
        break;

      case 'w:u':
        textPr[elem.TextPr.underline] = true;
        break;

      case 'w:vertAlign':
        textPr[elem.TextPr.vertAlign] = childNode.getAttribute('w:val');
        break;
    }
  }

  rPrObj[elem.RunPr.textPr] = textPr;
  return rPrObj;
}

function parseRun(parent, runElem) {
  if (!runElem || !runElem.getAttribute) {
    return null;
  }

  let rPrElem = firstElementByTagName(runElem, 'w:rPr');
  let rPr = parseRunPR(rPrElem);
  let run = new elem.Run(parent, rPr);

  for (let i = 0; i < runElem.childNodes.length; i++) {
    const childNode = runElem.childNodes[i];
    if (!childNode) continue;
    if (childNode.nodeType !== 1) continue;
    let child = null;

    switch (childNode.tagName) {
      case 'w:t':
        child = parseText(run, childNode);
        if (!child) run = null;
        break;

      case 'w:tab':
        child = new elem.SPCharRun(run, null, elem.SPChars.tab);
        addChild2Parent(run, child);
        break;

      case 'w:br':
        child = new elem.SPCharRun(run, null, elem.SPChars.tab);
        addChild2Parent(run, child);
        break;

      case 'w:drawing':
        parseDrawing(run, childNode);
        break;

      case 'w:object':
        parseObject(run, childNode);
        break;
    }
  }

  addChild2Parent(parent, run);
  return run;
}

function getChildRuns(parent, pElem) {
  if (!pElem || !pElem.getAttribute) {
    return null;
  }

  let runs = [];

  for (let i = 0; i < pElem.childNodes.length; i++) {
    const childNode = pElem.childNodes[i];
    if (!childNode) continue;
    if (childNode.nodeType !== 1) continue;

    switch (childNode.tagName) {
      case 'w:r':
        let run = parseRun(parent, childNode);
        runs.push(run);
        break;
    }
  }

  return runs;
}

function parseHyperlink(parent, hyperlinkElem) {
  let hyperLink = new elem.HyperLink(parent, null);
  let runs = getChildRuns(hyperLink, hyperlinkElem);
  addChild2Parent(parent, hyperLink);
  return hyperLink;
}

function parseSmartTag(parent, smartTagElem) {
  return getChildRuns(parent, smartTagElem);
}

function parseSimpleField(parent, fldSimpleElem) {
  return getChildRuns(parent, fldSimpleElem);
}

function parseOMath(parent, oMathElem) {
  if (!oMathElem || !oMathElem.getAttribute) {
    return null;
  }

  let oMathPr = {};
  oMathPr[elem.OMathPr.omml] = oMathElem.toString();
  let oMath = new elem.OMath(parent, oMathPr);
  addChild2Parent(parent, oMath);
  return oMath;
}

function parseIdentPr(identPrElem) {
  if (!identPrElem || !identPrElem.getAttribute) {
    return null;
  }

  let identPr = {};
  let left = identPrElem.getAttribute('w:left') || identPrElem.getAttribute('w:start');
  let right = identPrElem.getAttribute('w:right') || identPrElem.getAttribute('w:end');
  let hanging = identPrElem.getAttribute('w:hanging');
  let firstLine = identPrElem.getAttribute('w:firstLine');
  identPr[elem.IndentPr.left] = parseInt(left);
  identPr[elem.IndentPr.right] = parseInt(right);
  identPr[elem.IndentPr.hanging] = parseInt(hanging);
  identPr[elem.IndentPr.firstLine] = parseInt(firstLine);
  return identPr;
}

function parsePBorderPr(pBorderElem) {
  if (!pBorderElem || !pBorderElem.getAttribute) {
    return null;
  }

  let pBorderPr = {};

  for (let i = 0; i < pBorderElem.childNodes.length; i++) {
    const childNode = pBorderElem.childNodes[i];
    if (!childNode) continue;
    if (childNode.nodeType !== 1) continue;

    switch (childNode.tagName) {
      case 'w:top':
        pBorderPr[elem.PBorderPr.top] = parseBorderPr(childNode);
        break;

      case 'w:bottom':
        pBorderPr[elem.PBorderPr.bottom] = parseBorderPr(childNode);
        break;

      case 'w:left':
        pBorderPr[elem.PBorderPr.left] = parseBorderPr(childNode);
        break;

      case 'w:right':
        pBorderPr[elem.PBorderPr.right] = parseBorderPr(childNode);
        break;
    }
  }

  return pBorderPr;
}

function parseNumPr(numPrElem) {
  if (!numPrElem || !numPrElem.getAttribute) {
    return null;
  }

  let numPr = {};

  for (let i = 0; i < numPrElem.childNodes.length; i++) {
    const childNode = numPrElem.childNodes[i];
    if (!childNode) continue;
    if (childNode.nodeType !== 1) continue;

    switch (childNode.tagName) {
      case 'w:ilvl':
        numPr[elem.NumPr.level] = childNode.getAttribute('w:val');
        break;

      case 'w:numId':
        numPr[elem.NumPr.numId] = childNode.getAttribute('w:val');
        break;
    }
  }

  return numPr;
}

function parsePPR(pPrElem) {
  if (!pPrElem || !pPrElem.hasChildNodes) {
    return null;
  }

  let pPrObj = {};

  for (let i = 0; i < pPrElem.childNodes.length; i++) {
    const childNode = pPrElem.childNodes[i];
    if (!childNode) continue;
    if (childNode.nodeType !== 1) continue;

    switch (childNode.tagName) {
      case 'w:pStyle':
        pPrObj[elem.ParagraphPr.styleId] = childNode.getAttribute('w:val');
        break;

      case 'w:spacing':
        pPrObj[elem.ParagraphPr.spacing] = parseSpacingPr(childNode);
        break;

      case 'w:jc':
        pPrObj[elem.ParagraphPr.textAlign] = childNode.getAttribute('w:val');
        break;

      case 'w:ind':
        pPrObj[elem.ParagraphPr.ident] = parseIdentPr(childNode);
        break;

      case 'w:pBdr':
        pPrObj[elem.ParagraphPr.border] = parsePBorderPr(childNode);
        break;

      case 'w:rPr':
        pPrObj[elem.ParagraphPr.runPr] = parseRunPR(childNode);
        break;

      case 'w:numPr':
        pPrObj[elem.ParagraphPr.numPr] = parseNumPr(childNode);
        break;
    }
  }

  return pPrObj;
}

function parseWP(parent, pElem) {
  let pPrElem = firstElementByTagName(pElem, 'w:pPr');
  let pPr = parsePPR(pPrElem);
  let p = new elem.Paragraph(parent, pPr);

  for (let i = 0; i < pElem.childNodes.length; i++) {
    const childNode = pElem.childNodes[i];
    if (!childNode) continue;
    if (childNode.nodeType !== 1) continue;

    switch (childNode.tagName) {
      case 'w:r':
        parseRun(p, childNode);
        break;

      case 'w:hyperlink':
        parseHyperlink(p, childNode);
        break;

      case 'w:smartTag':
        parseSmartTag(p, childNode);
        break;

      case 'w:fldSimple':
        parseSimpleField(p, childNode);
        break;

      case 'm:oMath':
      case 'm:oMathPara':
        parseOMath(p, childNode);
        break;

      default:
        break;
    }
  }

  addChild2Parent(parent, p);
  return p;
}

function parseTableCellPr(tcPrElem) {
  if (!tcPrElem || !tcPrElem.hasChildNodes) {
    return null;
  }

  let tcPrObj = {};

  for (let i = 0; i < tcPrElem.childNodes.length; i++) {
    const childNode = tcPrElem.childNodes[i];
    if (!childNode) continue;
    if (childNode.nodeType !== 1) continue;

    switch (childNode.tagName) {
      case 'w:tcW':
        tcPrObj[elem.TableCellPr.width] = parseInt(childNode.getAttribute('w:w'));
        break;

      case 'w:vAlign':
        tcPrObj[elem.TableCellPr.vAlign] = childNode.getAttribute('w:val');
        break;

      case 'w:gridSpan':
        tcPrObj[elem.TableCellPr.gridSpan] = parseInt(childNode.getAttribute('w:val'));
        break;

      case 'w:vMerge':
        tcPrObj[elem.TableCellPr.vMerge] = childNode.getAttribute('w:val') || 'continue';
        break;
    }
  }

  return tcPrObj;
}

function parseTableCell(parent, tcElem) {
  let tcPrElem = firstElementByTagName(tcElem, 'w:tcPr');
  let tcPr = parseTableCellPr(tcPrElem);
  let tableCell = new elem.TableCell(parent, tcPr);

  for (let i = 0; i < tcElem.childNodes.length; i++) {
    const childNode = tcElem.childNodes[i];
    if (!childNode) continue;
    if (childNode.nodeType !== 1) continue;

    switch (childNode.tagName) {
      case 'w:p':
        let p = parseWP(tableCell, childNode);
        break;
    }
  }

  addChild2Parent(parent, tableCell);
  let gridSpan = 1;

  if (tcPr && tcPr[elem.TableCellPr.gridSpan]) {
    gridSpan = tcPr[elem.TableCellPr.gridSpan];
  }

  for (let i = 0; i < gridSpan - 1; i++) {
    addChild2Parent(parent, {});
  }

  return tableCell;
}

function parseTableRowPr(trPrElem) {
  if (!trPrElem || !trPrElem.hasChildNodes) {
    return null;
  }

  let trPrObj = {};

  for (let i = 0; i < trPrElem.childNodes.length; i++) {
    const childNode = trPrElem.childNodes[i];
    if (!childNode) continue;
    if (childNode.nodeType !== 1) continue;

    switch (childNode.tagName) {
      case 'w:tblHeader':
        trPrObj[elem.TableRowPr.tblHeader] = true;
        break;

      case 'w:trHeight':
        trPrObj[elem.TableRowPr.height] = parseInt(childNode.getAttribute('w:val'));
        break;

      case 'w:jc':
        trPrObj[elem.TableRowPr.hAlign] = childNode.getAttribute('w:val');
        break;
    }
  }

  return trPrObj;
}

function parseTableRow(parent, trElem) {
  let trPrElem = firstElementByTagName(trElem, 'w:trPr');
  let trPr = parseTableRowPr(trPrElem);
  let tableRow = new elem.TableRow(parent, trPr);

  for (let i = 0; i < trElem.childNodes.length; i++) {
    const childNode = trElem.childNodes[i];
    if (!childNode) continue;
    if (childNode.nodeType !== 1) continue;
    let tableCell = null;

    switch (childNode.tagName) {
      case 'w:tc':
        tableCell = parseTableCell(tableRow, childNode);
        break;
    }
  }

  addChild2Parent(parent, tableRow);
  return tableRow;
}

function parseTableBorderPr(tblBorderElem) {
  if (!tblBorderElem || !tblBorderElem.getAttribute) {
    return null;
  }

  let tblBorderPr = {};

  for (let i = 0; i < tblBorderElem.childNodes.length; i++) {
    const childNode = tblBorderElem.childNodes[i];
    if (!childNode) continue;
    if (childNode.nodeType !== 1) continue;

    switch (childNode.tagName) {
      case 'w:top':
        tblBorderPr[elem.TblBorderPr.top] = parseBorderPr(childNode);
        break;

      case 'w:bottom':
        tblBorderPr[elem.TblBorderPr.bottom] = parseBorderPr(childNode);
        break;

      case 'w:left':
      case 'w:start':
        tblBorderPr[elem.TblBorderPr.left] = parseBorderPr(childNode);
        break;

      case 'w:right':
      case 'w:end':
        tblBorderPr[elem.TblBorderPr.right] = parseBorderPr(childNode);
        break;

      case 'w:insideH':
        tblBorderPr[elem.TblBorderPr.insideH] = parseBorderPr(childNode);
        break;

      case 'w:insideV':
        tblBorderPr[elem.TblBorderPr.insideV] = parseBorderPr(childNode);
        break;
    }
  }

  return tblBorderPr;
}

function parseTableCellMarPr(tblCellMarElem) {
  if (!tblCellMarElem || !tblCellMarElem.getAttribute) {
    return null;
  }

  let tblCellMarPr = {};

  for (let i = 0; i < tblCellMarElem.childNodes.length; i++) {
    const childNode = tblCellMarElem.childNodes[i];
    if (!childNode) continue;
    if (childNode.nodeType !== 1) continue;
    let val = childNode.getAttribute('w:w');

    switch (childNode.tagName) {
      case 'w:top':
        tblCellMarPr[elem.TblCellMarginPr.top] = parseInt(val);
        break;

      case 'w:bottom':
        tblCellMarPr[elem.TblCellMarginPr.bottom] = parseInt(val);
        break;

      case 'w:left':
      case 'w:start':
        tblCellMarPr[elem.TblCellMarginPr.left] = parseInt(val);
        break;

      case 'w:right':
      case 'w:end':
        tblCellMarPr[elem.TblCellMarginPr.right] = parseInt(val);
        break;
    }
  }

  return tblCellMarPr;
}

function parseTableFloatPr(tblpElem) {
  if (!tblpElem || !tblpElem.getAttribute) {
    return null;
  }

  let tblFloatPr = {};
  tblFloatPr[elem.TblFloatPr.horizontalAnchor] = tblpElem.getAttribute('w:horzAnchor');
  tblFloatPr[elem.TblFloatPr.verticalAnchor] = tblpElem.getAttribute('w:vertAnchor');
  return tblFloatPr;
}

function parseTablePR(tblPrElem) {
  if (!tblPrElem || !tblPrElem.hasChildNodes) {
    return null;
  }

  let tblPrObj = {};

  for (let i = 0; i < tblPrElem.childNodes.length; i++) {
    const childNode = tblPrElem.childNodes[i];
    if (!childNode) continue;
    if (childNode.nodeType !== 1) continue;

    switch (childNode.tagName) {
      case 'w:pStyle':
        tblPrObj[elem.TablePr.styleId] = childNode.getAttribute('w:val');
        break;

      case 'w:tblW':
        let tblW = 0;

        if (childNode.getAttribute('w:type') === 'dxa') {
          tblW = parseInt(childNode.getAttribute('w:w'));
        }

        tblPrObj[elem.TablePr.width] = tblW;
        break;

      case 'w:jc':
        tblPrObj[elem.TablePr.alignment] = childNode.getAttribute('w:val');
        break;

      case 'w:tblBorders':
        tblPrObj[elem.TablePr.border] = parseTableBorderPr(childNode);
        break;

      case 'w:tblCellMar':
        tblPrObj[elem.TablePr.cellMargin] = parseTableCellMarPr(childNode);
        break;

      case 'w:tblCellSpacing':
        tblPrObj[elem.TablePr.cellSpacing] = parseInt(childNode.getAttribute('w:w'));
        break;

      case 'w:tblpPr':
        tblPrObj[elem.TablePr.floatPr] = parseTableFloatPr(childNode);
        break;
    }
  }

  return tblPrObj;
}

function parseTableGrid(tblGridElem) {
  if (!tblGridElem || !tblGridElem.hasChildNodes) {
    return null;
  }

  let colsWidth = [];

  for (let i = 0; i < tblGridElem.childNodes.length; i++) {
    const childNode = tblGridElem.childNodes[i];
    if (!childNode) continue;
    if (childNode.nodeType !== 1) continue;

    switch (childNode.tagName) {
      case 'w:gridCol':
        let w = parseInt(childNode.getAttribute('w:w'));
        colsWidth.push(w);
        break;
    }
  }
}

function parseWTBL(parent, tblElem) {
  let tblPrElem = firstElementByTagName(tblElem, 'w:tblPr');
  let tablePr = parseTablePR(tblPrElem);
  let tblGridElem = firstElementByTagName(tblElem, 'w:tblGrid');
  let colsWidth = parseTableGrid(tblGridElem);
  let table = new elem.Table(parent, tablePr, colsWidth);

  for (let i = 0; i < tblElem.childNodes.length; i++) {
    const childNode = tblElem.childNodes[i];
    if (!childNode) continue;
    if (childNode.nodeType !== 1) continue;
    let tableRow = null;

    switch (childNode.tagName) {
      case 'w:tr':
        tableRow = parseTableRow(table, childNode);
        break;
    }
  }

  addChild2Parent(parent, table);
  return table;
}

function parseDocument(documentXml) {
  if (!documentXml) {
    return null;
  }

  const dom = new DOMParser().parseFromString(documentXml, 'text/xml');
  let bodyElem = firstElementByTagName(dom.documentElement, 'w:body');

  if (!bodyElem) {
    return null;
  }

  let doc = new elem.Document(null, null);
  let body = new elem.Body(doc, null);
  let bodyChild = null;

  for (let i = 0; i < bodyElem.childNodes.length; i++) {
    const childNode = bodyElem.childNodes[i];
    if (!childNode) continue;
    if (childNode.nodeType !== 1) continue;

    switch (childNode.tagName) {
      case 'w:p':
        bodyChild = parseWP(body, childNode);
        break;

      case 'w:tbl':
        bodyChild = parseWTBL(body, childNode);
        break;
    }
  }

  addChild2Parent(doc, body);
  return doc;
}

class Docx {
  constructor(settings, relations, numbering, styles, document) {
    this.settings = settings;
    this.relations = relations;
    this.numbering = numbering;
    this.styles = styles;
    this.document = document;
  }

  getRelByRId(rId) {
    if (!this.relations) return null;
    return this.relations.getRelByRId(rId);
  }

  setNumLevel(numId, ilvl) {
    if (!this.numbering) return null;
    return this.numbering.setNumLevel(numId, ilvl);
  }

  getParagraphStyle(styleId) {
    if (!this.styles) return null;
    return this.styles.getParagraphStyle(styleId);
  }

}

exports.Docx = Docx;

function parseDocx(xmls, files) {
  const settingsXml = xmls['settings'];
  const relationsXml = xmls['relations'];
  const numberingXml = xmls['numbering'];
  const stylesXml = xmls['styles'];
  const documentXml = xmls['document'];
  const settings = Settings.parse(settingsXml);
  console.log('settings=', settings);
  const relations = Relations.parse(relationsXml);
  console.log('relations=', relations);
  const numbering = Numbering.parse(numberingXml);
  console.log('numbering=', numbering);
  const styles = Styles.parse(stylesXml);
  console.log('styles=', styles);
  const document = parseDocument(documentXml);
  console.log('document=', document);
  let docx = new Docx(settings, relations, numbering, styles, document);

  if (!docx) {}

  return docx;
}

function getDrawingAndObjects(htElem, elems) {
  if (!htElem) return;
  let imgRId = null;
  let oleRId = null;

  if (htElem.properties) {
    imgRId = htElem.properties[elem.DrawingPr.imgRId];
    oleRId = htElem.properties[elem.ObjectPr.oleRId];
  }

  if (imgRId) elems.push(htElem);

  if (!Array.isArray(htElem.children)) {
    return;
  }

  for (let htChild of htElem.children) {
    getDrawingAndObjects(htChild, elems);
  }
}

async function dealDrawingOrObject(htElem, docx, pConvtHelper) {
  if (!htElem) return;
  let imgRId = null;
  let oleRId = null;

  if (htElem.properties) {
    imgRId = htElem.properties[elem.DrawingPr.imgRId];
    oleRId = htElem.properties[elem.ObjectPr.oleRId];
  }

  let svg = null;

  if (oleRId) {
    const rel = docx.getRelByRId(oleRId);

    if (rel && rel.Target) {
      const oleName = 'word/' + rel.Target;
      const latex = await pConvtHelper.getLatex(oleName);

      if (latex) {
        //svg = await pConvtHelper.latex2svg(latex);
        htElem.properties[elem.ObjectPr.imgSrc] = svg;
      }

      htElem.properties[elem.ObjectPr.latex] = latex;
    } else {
      console.warn('no rel for oleRId=', imgRId);
    }
  }

  if (imgRId && !svg) {
    const rel = docx.getRelByRId(imgRId);

    if (rel && rel.Target) {
      const imgName = 'word/' + rel.Target;
      const imgSrc = await pConvtHelper.getImgSrc(imgName, oleRId);

      if (imgSrc) {}

      htElem.properties[elem.DrawingPr.imgSrc] = imgSrc;
    } else {
      console.warn('no rel for imgRId=', imgRId);
    }
  }
}

async function createHT(docx, htDocImpl, pConvtHelper) {
  if (!docx.document) return null;
  let elems = [];
  getDrawingAndObjects(docx.document, elems);

  for (const htElem of elems) {
    if (!htElem || !pConvtHelper) {
      continue;
    }

    await dealDrawingOrObject(htElem, docx, pConvtHelper);
  }

  const htDocElem = docx.document.toHTElem(htDocImpl, docx);
  return htDocElem.toString();
}