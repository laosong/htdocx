/**
 * Created by laosong on 2020/2/18.
 * create docx from htXml
 */
'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.createDocx = createDocx;

const _ = require('lodash/core');

const xmlDOM = require('xmldom');

const docx = require('docx');

const helper = require('./helper');

const omath = require('./ext/omath');
/**
 * 200px => 200
 * @param pxStr
 */


function px2Num(pxStr) {
  return parseInt(pxStr, 10);
}

function pxNum2DXA(pxNum) {
  return pxNum * 15;
}
/**
 * rgb(255,255,255) => FFFFFF
 * @param rgbStr
 * @returns {string}
 */


function rgb2hex(rgbStr) {
  const rgb = rgbStr.match(/^rgb\((\d+),\s*(\d+),\s*(\d+)\)$/);
  return ('0' + parseInt(rgb[1], 10).toString(16)).slice(-2) + ('0' + parseInt(rgb[2], 10).toString(16)).slice(-2) + ('0' + parseInt(rgb[3], 10).toString(16)).slice(-2);
}

function fontPX2PT(pxStr) {
  const pxNum = parseInt(pxStr);
  let pt = helper.px2ptTab[String(pxNum)];

  if (!pt) {
    pt = pxNum > 33 ? 25 : 10;
  }

  return pt;
}

function createTextRun(textContent, spanNode) {
  let textRunOpts = {
    text: textContent
  };
  const fontFamily = spanNode.getAttribute('fontFamily');
  const fontSize = spanNode.getAttribute('fontSize');
  const fontWeight = spanNode.getAttribute('fontWeight');
  const color = spanNode.getAttribute('color');
  const backgroundColor = spanNode.getAttribute('backgroundColor');
  const verticalAlign = spanNode.getAttribute('verticalAlign');

  if (fontSize) {
    let fontPt = fontPX2PT(fontSize);
    textRunOpts.size = 2 * fontPt;
  }

  if (parseInt(fontWeight) >= 700) {
    textRunOpts.bold = true;
  }

  if (color) {
    textRunOpts.color = rgb2hex(color);
  }

  if (backgroundColor) {
    textRunOpts.highlight = 'yellow';
  }

  if (verticalAlign === 'sub') {
    textRunOpts.subScript = true;
  } else if (verticalAlign === 'sup') {
    textRunOpts.superScript = true;
  }

  let textRun = new docx.TextRun(textRunOpts);

  if (!textRun) {}

  if (spanNode.hasAttribute('br')) {
    textRun.break();
  }

  return textRun;
}

function createBr(spanNode) {
  let textRunOpts = {
    text: ''
  };
  let textRun = new docx.TextRun(textRunOpts);
  textRun.break();
  return textRun;
}

function createOMath(omml, spanNode) {
  const fontFamily = spanNode.getAttribute('fontFamily');
  const fontSize = spanNode.getAttribute('fontSize');
  const fontWeight = spanNode.getAttribute('fontWeight');
  const color = spanNode.getAttribute('color');
  const backgroundColor = spanNode.getAttribute('backgroundColor');
  const verticalAlign = spanNode.getAttribute('verticalAlign');
  let fontPt = 0;

  if (fontSize) {
    fontPt = fontPX2PT(fontSize);
  }

  let omath1 = new omath.OMath(omml, 2 * fontPt);
  omath1.init();
  return omath1;
}

function createSpan(ctx, spanNode) {
  let runs = [];
  const startData = {
    type: 'text',
    data: spanNode.textContent
  };
  const texts = helper.splitAtDelimiters(startData, '@math-', '@', false);

  for (let i = 0; i < texts.length; i++) {
    const textB = texts[i];
    let textRun = null;

    if (textB.type === 'text') {
      let text = helper.decodeEntities(textB.data);
      text = text.replace(/[\\t]/g, '  ');
      textRun = createTextRun(text, spanNode);
    } else if (textB.type === 'math') {
      const mathId = textB.rawData;
      let omml = null;

      if (ctx.maths.hasOwnProperty(mathId)) {
        omml = ctx.maths[mathId];
      }

      if (!omml) {
        continue;
      }

      textRun = createOMath(omml, spanNode);
    }

    if (textRun) runs.push(textRun);
  }

  return runs;
}

function createImg(ctx, imgNode) {
  const src = imgNode.getAttribute('src');
  const width = px2Num(imgNode.getAttribute('width'));
  const height = px2Num(imgNode.getAttribute('height'));
  const cssFloat = imgNode.getAttribute('cssFloat');
  const verticalAlign = imgNode.getAttribute('verticalAlign');

  if (!width || !height) {
    return;
  }

  let imgBuf = null;

  if (ctx && ctx.images) {
    let findIt = _.find(ctx.images, {
      url: src
    });

    imgBuf = findIt ? findIt.buf : null;
  }

  if (!imgBuf) {
    return;
  }

  let drawingOptions = {};

  if (cssFloat === 'left' || cssFloat === 'right') {
    const hpRelative = docx.HorizontalPositionRelativeFrom.MARGIN;
    const hpAlign = cssFloat === 'left' ? docx.HorizontalPositionAlign.LEFT : docx.HorizontalPositionAlign.RIGHT;
    const vpRelative = docx.VerticalPositionRelativeFrom.PARAGRAPH;
    const vpAlign = verticalAlign === 'middle' ? docx.VerticalPositionAlign.CENTER : docx.VerticalPositionAlign.TOP;
    drawingOptions.floating = {
      horizontalPosition: {
        relative: hpRelative,
        align: hpAlign
      },
      verticalPosition: {
        relative: vpRelative,
        offset: 100
      },
      wrap: {
        type: docx.TextWrappingType.SQUARE,
        side: docx.TextWrappingSide.BOTH_SIDES
      }
    };
  }

  let pictRun = docx.Media.addImage(ctx.htDocx, imgBuf, width, height, drawingOptions);

  if (!pictRun) {}

  return pictRun;
}
/**
 *
 * @param ctx
 * @param pNode
 * @param numbering
 * @returns {docx.Paragraph}
 */


function createParagraph(ctx, pNode, numbering) {
  if (!pNode || !pNode.childNodes) {
    console.warn('createParagraph but pNode', pNode);
    return null;
  }

  let pRuns = [];

  for (let i = 0; i < pNode.childNodes.length; i++) {
    const childNode = pNode.childNodes[i];

    if (childNode.tagName === 'span') {
      let textRun = createSpan(ctx, childNode);

      if (!textRun) {
        continue;
      }

      pRuns = pRuns.concat(textRun);
    } else if (childNode.tagName === 'img') {
      let pictRun = createImg(ctx, childNode);

      if (!pictRun) {
        continue;
      }

      pRuns.push(pictRun);
    }
  }

  if (!pRuns.length) {
    return null;
  }

  const paragraph = new docx.Paragraph({
    children: pRuns
  });

  if (!paragraph) {}

  return paragraph;
}

function createUL(ctx, ulNode) {
  if (!ulNode || !ulNode.childNodes) {
    console.warn('createUL ulNode but ulNode', ulNode);
    return null;
  }

  let paragraphs = [];
  let numbering = null;

  for (let i = 0; i < ulNode.childNodes.length; i++) {
    const childNode = ulNode.childNodes[i];

    if (!childNode || childNode.tagName !== 'li') {
      continue;
    }

    let pNode = childNode.firstChild;

    if (!pNode) {
      continue;
    }

    const liP = createParagraph(ctx, pNode, numbering);

    if (!liP) {
      continue;
    }

    paragraphs.push(liP);
  }

  return paragraphs;
}

function createTableCell(ctx, tableCellNode) {
  if (!tableCellNode || !tableCellNode.childNodes) {
    console.warn('createTableCell but tableCellNode', tableCellNode);
    return null;
  }

  let colSpan = parseInt(tableCellNode.getAttribute('colSpan'));
  let rowSpan = parseInt(tableCellNode.getAttribute('rowSpan'));
  let width = px2Num(tableCellNode.getAttribute('width'));
  let verticalAlign = px2Num(tableCellNode.getAttribute('verticalAlign'));
  let blocks = [];

  for (let i = 0; i < tableCellNode.childNodes.length; i++) {
    const childNode = tableCellNode.childNodes[i];

    if (childNode.tagName === 'div') {//ignore
    } else if (childNode.tagName === 'p') {
      const p = createParagraph(ctx, childNode);
      if (p) blocks.push(p);
    } else if (childNode.tagName === 'ul') {
      const ps = createUL(ctx, childNode);
      if (Array.isArray(ps)) blocks = blocks.concat(ps);
    } else if (childNode.tagName === 'ol') {
      const ps = createUL(ctx, childNode);
      if (Array.isArray(ps)) blocks = blocks.concat(ps);
    } else if (childNode.tagName === 'table') {
      const table = createTable(ctx, childNode);
      if (table) blocks.push(table);
    }
  }

  let tcOptions = {
    children: blocks
  };

  if (colSpan > 1) {
    tcOptions.columnSpan = colSpan;
  }

  if (rowSpan > 1) {
    tcOptions.rowSpan = rowSpan;
  }

  if (width) {
    tcOptions.width = {
      size: pxNum2DXA(width),
      type: docx.WidthType.DXA
    };
  }

  if (verticalAlign) {}

  const tableCell = new docx.TableCell(tcOptions);

  if (!tableCell) {} else if (width && tableCell.properties) {
    tableCell.properties.setWidth(pxNum2DXA(width), docx.WidthType.DXA);
  }

  return tableCell;
}

function createTableRow(ctx, tableRowNode) {
  if (!tableRowNode || !tableRowNode.childNodes) {
    console.warn('createTableRow but tableRowNode', tableRowNode);
    return null;
  }

  let tableCells = [];

  for (let i = 0; i < tableRowNode.childNodes.length; i++) {
    const childNode = tableRowNode.childNodes[i];
    let tableCell = createTableCell(ctx, childNode);

    if (tableCell) {
      tableCells.push(tableCell);
    }
  }

  let trOptions = {
    children: tableCells
  };

  if (tableRowNode.hasAttribute('header')) {
    trOptions.tableHeader = true;
  }

  const tableRow = new docx.TableRow(trOptions);

  if (!tableRow) {}

  return tableRow;
}
/**
 *
 * @param ctx
 * @param tableNode
 * @returns {docx.Table}
 */


function createTable(ctx, tableNode) {
  if (!tableNode || !tableNode.childNodes) {
    console.warn('createTable but createTable', tableNode);
    return null;
  }

  let width = px2Num(tableNode.getAttribute('width'));
  let tableRows = [];

  for (let i = 0; i < tableNode.childNodes.length; i++) {
    const childNode = tableNode.childNodes[i];
    let tableRow = createTableRow(ctx, childNode);

    if (tableRow) {
      tableRows.push(tableRow);
    }
  }

  let tableOptions = {
    rows: tableRows
  };

  if (width) {
    tableOptions.width = {
      size: pxNum2DXA(width),
      type: docx.WidthType.DXA
    };
  }

  tableOptions.margins = {
    marginUnitType: docx.WidthType.DXA,
    top: 75,
    bottom: 75,
    left: 75,
    right: 75
  };
  tableOptions.layout = docx.TableLayoutType.FIXED;
  const table = new docx.Table(tableOptions);

  if (!table) {}

  return table;
}
/**
 * 创建一个不可见表格作为题目之间的分隔
 */


function createBreak(ctx) {
  const paragraph = new docx.Paragraph('');
  const cell = new docx.TableCell({
    children: [paragraph],
    width: {
      size: 8932,
      type: docx.WidthType.DXA
    },
    borders: {
      top: {
        style: docx.BorderStyle.NONE,
        size: 1
      },
      bottom: {
        style: docx.BorderStyle.NONE,
        size: 1
      },
      left: {
        style: docx.BorderStyle.NONE,
        size: 1
      },
      right: {
        style: docx.BorderStyle.NONE,
        size: 1
      }
    }
  });
  const row = new docx.TableRow({
    children: [cell]
  });
  const table = new docx.Table({
    rows: [row],
    width: {
      size: 8932,
      type: docx.WidthType.DXA
    },
    columnWidths: [8932],
    layout: docx.TableLayoutType.FIXED
  });

  if (!table) {
    console.error('createBreak !table');
  }

  return table;
}
/**
 * 根据html（简化）创建docx对象
 * @param htDoc
 * @param images
 * @param maths
 * @param options
 * @returns {*}
 */


function createDocx(htDoc, images, maths, options) {
  const rootNode = htDoc.documentElement;
  const htDocx = new docx.Document(options);

  if (!rootNode || !rootNode.childNodes) {
    console.warn('createDocx but !rootNode');
    return docx;
  }

  let ctx = {
    htDocx,
    images,
    maths
  };
  let blocks = [];

  for (let i = 0; i < rootNode.childNodes.length; i++) {
    const childNode = rootNode.childNodes[i];

    if (childNode.tagName === 'div') {
      const abreak = createBreak(ctx);
      if (abreak) blocks.push(abreak);
    } else if (childNode.tagName === 'p') {
      const p = createParagraph(ctx, childNode);
      if (p) blocks.push(p);
    } else if (childNode.tagName === 'ul') {
      const ps = createUL(ctx, childNode);
      if (Array.isArray(ps)) blocks = blocks.concat(ps);
    } else if (childNode.tagName === 'ol') {
      const ps = createUL(ctx, childNode);
      if (Array.isArray(ps)) blocks = blocks.concat(ps);
    } else if (childNode.tagName === 'table') {
      const table = createTable(ctx, childNode);
      if (table) blocks.push(table);
    }
  }

  htDocx.addSection({
    children: blocks
  });
  return htDocx;
}