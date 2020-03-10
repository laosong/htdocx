/**
 * Created by laosong on 2020/2/18.
 * simplified elements of docx
 */
'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.TableCell = exports.TableCellPr = exports.TableRow = exports.TableRowPr = exports.Table = exports.TablePr = exports.TblFloatPr = exports.TblCellMarginPr = exports.TblBorderPr = exports.Paragraph = exports.ParagraphPr = exports.NumPr = exports.PBorderPr = exports.IndentPr = exports.OMath = exports.OMathPr = exports.ComplexField = exports.SimpleField = exports.SmartTag = exports.HyperLink = exports.ObjectRun = exports.ObjectPr = exports.DrawingRun = exports.DrawingPr = exports.DrawingFloatPr = exports.TextWrapMode = exports.SPCharRun = exports.SPChars = exports.TextRun = exports.Run = exports.RunPr = exports.TextPr = exports.BorderPr = exports.SpacingPr = exports.Body = exports.Document = void 0;

const xmlDom = require('xmldom');

const helper = require('./helper');

class HTComponent {
  constructor(parent, properties) {
    this.parent = parent;
    this.properties = properties;
    this.children = null;
  }

  addChild(child) {
    if (!Array.isArray(this.children)) {
      this.children = [];
    }

    this.children.push(child);
  }

  toHTElem(htDoc, docx) {
    console.warn('toHTElem', this.children);
    return null;
  }

}

function buildChildren(htElem, children, docx) {
  if (!htElem) return;
  if (!Array.isArray(children)) return;

  for (let child of children) {
    if (!child || !child.toHTElem) continue;
    let htChild = child.toHTElem(htElem.ownerDocument, docx);
    if (!htChild) continue;
    htElem.appendChild(htChild);
  }
}

class Document extends HTComponent {
  constructor(parent, options) {
    super(parent, null);
    this.options = options;
  }

  toHTElem(htDoc, docx) {
    let htElem = htDoc.createElement('html');
    buildChildren(htElem, this.children, docx);
    return htElem;
  }

}

exports.Document = Document;

class Body extends HTComponent {
  constructor(parent, properties) {
    super(parent, properties);
  }

  toHTElem(htDoc, docx) {
    let htElem = htDoc.createElement('body');
    buildChildren(htElem, this.children, docx);
    return htElem;
  }

}

exports.Body = Body;

function obj2styleString(styleObj) {
  return Object.entries(styleObj).reduce((styleString, [propName, propValue]) => {
    return `${styleString}${propName}:${propValue};`;
  }, '');
}

function dxa2PxNum(dxa) {
  return Math.floor(dxa / 15);
}

function emu2PxNum(emu) {
  return Math.floor(emu / 9525);
}

const SpacingPr = {
  'before': 'before',
  //number dxa
  'after': 'after',
  //number
  'line': 'line' //number

};
exports.SpacingPr = SpacingPr;
const BorderPr = {
  'style': 'style',
  //string
  'width': 'width',
  //number
  'color': 'color' //hex number

};
exports.BorderPr = BorderPr;

function setBorderStyle(styles, borderPr, which) {
  let borderStyleName = 'border-style';
  let borderWidthName = 'border-width';
  let borderColorName = 'border-color';

  if (which) {
    borderStyleName = 'border-' + which + '-style';
    borderWidthName = 'border-' + which + '-width';
    borderColorName = 'border-' + which + '-color';
  }

  styles[borderStyleName] = 'solid';

  switch (borderPr.style) {
    case 'dashed':
      styles[borderStyleName] = 'dashed';
      break;

    case 'dotted':
      styles[borderStyleName] = 'dotted';
      break;
  }

  styles['border-width'] = '1px';

  if (borderPr.width) {
    styles[borderWidthName] = borderPr.width / 8 + 'px';
  }

  if (borderPr.color) {
    styles[borderColorName] = '#' + borderPr.color;
  }
}

const TextPr = {
  'bold': 'bold',
  //boolean
  'italics': 'italics',
  //boolean
  'underline': 'underline',
  //boolean
  'vertAlign': 'vertAlign' //string superscript,subscript

};
exports.TextPr = TextPr;
const RunPr = {
  'styleId': 'styleId',
  //string
  'border': 'border',
  //object BorderPr
  'highlight': 'highlight',
  //string yellow
  'shdFillColor': 'shdFillColor',
  // hex number
  'fontName': 'fontName',
  //string
  'fontSize': 'fontSize',
  //number
  'textColor': 'textColor',
  //hex number
  'textPr': 'textPr' //object

};
exports.RunPr = RunPr;

function fontSZ2PX(fontSize) {
  let px = helper.sz2pxTab[String(fontSize)];

  if (!px) {
    px = fontSize > 50 ? 33 : 14;
  }

  return px;
}

class Run extends HTComponent {
  constructor(parent, properties) {
    super(parent, properties);
  }

  toHTElem(htDoc, docx) {
    let htElem = htDoc.createElement('span');
    let properties = this.properties || {};
    let styles = {};

    if (properties[RunPr.border]) {
      setBorderStyle(styles, properties[RunPr.border], null);
    }

    if (properties[RunPr.highlight]) {
      styles['background-color'] = properties[RunPr.highlight];
    } else if (properties[RunPr.shdFillColor]) {
      styles['background-color'] = '#' + properties[RunPr.shdFillColor];
    }

    if (properties[RunPr.fontName]) {
      styles['font-family'] = properties[RunPr.fontName];
    }

    if (properties[RunPr.fontSize]) {
      styles['font-size'] = fontSZ2PX(properties[RunPr.fontSize]) + 'px';
    }

    if (properties[RunPr.textColor]) {
      styles['color'] = '#' + properties[RunPr.textColor];
    }

    let cp = htElem;
    const textPr = properties[RunPr.textPr];

    if (textPr) {
      if (textPr[TextPr.bold]) {
        styles['font-weight'] = 800;
      }

      if (textPr[TextPr.italics]) {
        styles['font-style'] = 'italic';
      }

      if (textPr[TextPr.underline]) {
        styles['text-decoration'] = 'underline';
      }

      if (textPr[TextPr.vertAlign] === 'superscript') {
        let supElem = htDoc.createElement('sup');
        cp.appendChild(supElem);
        cp = supElem;
      } else if (textPr[TextPr.vertAlign] === 'subscript') {
        let subElem = htDoc.createElement('sub');
        cp.appendChild(subElem);
        cp = subElem;
      }
    }

    let styleStr = obj2styleString(styles);

    if (styleStr) {
      htElem.setAttribute('style', styleStr);
    }

    buildChildren(cp, this.children, docx);
    return htElem;
  }

}

exports.Run = Run;

class TextRun extends HTComponent {
  constructor(parent, properties, content) {
    super(parent, properties);
    this.textContent = content;
  }

  toHTElem(htDoc, docx) {
    let htElem = htDoc.createTextNode(this.textContent);

    if (!htElem) {}

    return htElem;
  }

}

exports.TextRun = TextRun;
const SPChars = {
  'tab': 'tab',
  'br': 'br'
};
exports.SPChars = SPChars;

class SPCharRun extends HTComponent {
  constructor(parent, properties, spChar) {
    super(parent, properties);
    this.spChar = spChar;
  }

  toHTElem(htDoc) {
    let htElem = null;

    if (this.spChar === SPChars.tab) {
      htElem = htDoc.createTextNode('\u00A0\u00A0');
    } else if (this.spChar === SPChars.br) {
      htElem = htDoc.createElement('br');
    }

    return htElem;
  }

}

exports.SPCharRun = SPCharRun;
const TextWrapMode = {
  'wrapNone': 'wrapNone',
  'wrapSquare': 'wrapSquare',
  'wrapThrough': 'wrapThrough',
  'wrapTight': 'wrapTight',
  'wrapTopAndBottom': 'wrapTopAndBottom'
};
exports.TextWrapMode = TextWrapMode;
const DrawingFloatPr = {
  'horizontalPosition': 'horizontalPosition',
  'verticalPosition': 'verticalPosition',
  //string
  'textWrap': 'textWrap'
};
exports.DrawingFloatPr = DrawingFloatPr;
const DrawingPr = {
  'floatPr': 'floatPr',
  'title': 'title',
  //string
  'cx': 'cx',
  //number EMU
  'cy': 'cy',
  //number EMU
  'imgRId': 'imgRId',
  //string
  'imgSrc': 'imgSrc' //string

};
exports.DrawingPr = DrawingPr;

class DrawingRun extends HTComponent {
  constructor(parent, properties) {
    super(parent, properties);
  }

  toHTElem(htDoc) {
    let htElem = htDoc.createElement('img');
    let properties = this.properties || {};
    let rId = properties[DrawingPr.imgRId];
    let src = properties[DrawingPr.imgSrc];
    let alt = properties[DrawingPr.title];
    let width = properties[DrawingPr.cx];
    let height = properties[DrawingPr.cy];
    htElem.setAttribute('src', src);

    if (alt) {
      htElem.setAttribute('alt', alt);
    }

    if (width) {
      htElem.setAttribute('width', emu2PxNum(width));
    }

    if (height) {
      htElem.setAttribute('height', emu2PxNum(height));
    }

    htElem.setAttribute('data-rid', rId);
    let styles = {};
    const floatPr = properties[DrawingPr.floatPr];

    if (floatPr) {
      styles['float'] = 'right';
    } else {
      styles['vertical-align'] = 'middle';
    }

    let styleStr = obj2styleString(styles);

    if (styleStr) {
      htElem.setAttribute('style', styleStr);
    }

    return htElem;
  }

}

exports.DrawingRun = DrawingRun;
const ObjectPr = {
  'shapeId': 'shapeId',
  //string
  'shapeStyle': 'shapeStyle',
  //string
  'imgRId': 'imgRId',
  //string
  'imgSrc': 'imgSrc',
  //string,
  'oleProgId': 'oleProgId',
  'oleShapeId': 'oleShapeId',
  'oleRId': 'oleRId',
  'latex': 'latex',
  //string
  'svgUrl': 'svgUrl' //string

};
exports.ObjectPr = ObjectPr;

class ObjectRun extends HTComponent {
  constructor(parent, properties) {
    super(parent, properties);
  }

  toHTElem(htDoc) {
    let htElem = htDoc.createElement('img');
    let properties = this.properties || {};
    let rId = properties[ObjectPr.imgRId];
    let src = properties[ObjectPr.imgSrc];
    let alt = properties[ObjectPr.latex];
    htElem.setAttribute('src', src);

    if (alt) {
      htElem.setAttribute('alt', alt);
    }

    htElem.setAttribute('data-rid', rId);
    let styles = {};
    const floatPr = properties[DrawingPr.floatPr];

    if (floatPr) {
      styles['float'] = 'right';
    } else {
      styles['vertical-align'] = 'middle';
    }

    let styleStr = obj2styleString(styles);

    if (styleStr) {
      htElem.setAttribute('style', styleStr);
    }

    return htElem;
  }

}

exports.ObjectRun = ObjectRun;

class HyperLink extends HTComponent {
  constructor(parent, properties) {
    super(parent, properties);
  }

  toHTElem(htDoc, docx) {
    let htElem = htDoc.createElement('a');
    buildChildren(htElem, this.children, docx);
    return htElem;
  }

}

exports.HyperLink = HyperLink;

class SmartTag extends HTComponent {}

exports.SmartTag = SmartTag;

class SimpleField extends HTComponent {}

exports.SimpleField = SimpleField;

class ComplexField extends HTComponent {}

exports.ComplexField = ComplexField;
const OMathPr = {
  'omml': 'omml',
  //string
  'mathml': 'mathml' //string

};
exports.OMathPr = OMathPr;

class OMath extends HTComponent {
  constructor(parent, properties) {
    super(parent, properties);
  }

  toHTElem(htDoc) {
    let htElem = htDoc.createElement('math');
    return htElem;
  }

}

exports.OMath = OMath;
const IndentPr = {
  'left': 'left',
  //number dxa
  'right': 'right',
  //number dxa
  'hanging': 'hanging',
  //number dxa
  'firstLine': 'firstLine' //number dxa

};
exports.IndentPr = IndentPr;
const PBorderPr = {
  'top': 'top',
  //object BorderPr
  'bottom': 'bottom',
  'left': 'left',
  'right': 'right'
};
exports.PBorderPr = PBorderPr;
const NumPr = {
  'level': 'level',
  //string
  'numId': 'numId' //string

};
exports.NumPr = NumPr;
const ParagraphPr = {
  'styleId': 'styleId',
  //string
  'pageBreakBefore': 'pageBreakBefore',
  //boolean
  'heading': 'heading',
  //string
  'ident': 'ident',
  //object IndentPr
  'border': 'border',
  //object PBorderPr
  'spacing': 'spacing',
  //object SpacingPr
  'textAlign': 'textAlign',
  //string left,center,right,both
  'runPr': 'runPr',
  //object
  'numPr': 'numPr' //object

};
exports.ParagraphPr = ParagraphPr;
const PStyleName2Tag = {
  'heading 1': 'h1',
  'heading 2': 'h2',
  'heading 3': 'h3',
  'heading 4': 'h4',
  'heading 5': 'h5',
  'heading 6': 'h6'
};

function setPBorderStyle(styles, pBorderPr) {
  const top = pBorderPr[PBorderPr.top];

  if (top) {
    setBorderStyle(styles, top, 'top');
  }

  const bottom = pBorderPr[PBorderPr.bottom];

  if (top) {
    setBorderStyle(styles, bottom, 'bottom');
  }

  const left = pBorderPr[PBorderPr.left];

  if (top) {
    setBorderStyle(styles, left, 'left');
  }

  const right = pBorderPr[PBorderPr.right];

  if (right) {
    setBorderStyle(styles, right, 'right');
  }
}

class Paragraph extends HTComponent {
  constructor(parent, properties) {
    super(parent, properties);
  }

  toHTElem(htDoc, docx) {
    let tag = 'p';
    let properties = this.properties || {};
    const styleId = properties[ParagraphPr.styleId];

    if (styleId) {
      let pStyle = docx.getParagraphStyle(styleId);

      if (pStyle && pStyle.name) {
        tag = PStyleName2Tag[pStyle.name] || 'p';
      }
    }

    let htElem = htDoc.createElement(tag);
    let styles = {};
    const identPr = properties[ParagraphPr.ident];

    if (identPr) {
      const left = identPr[IndentPr.left];
      const right = identPr[IndentPr.right];
      const hanging = identPr[IndentPr.hanging];
      const firstLine = identPr[IndentPr.firstLine];

      if (firstLine) {
        styles['text-indent'] = dxa2PxNum(firstLine) + 'px';
      }

      if (left) {
        styles['margin-left'] = dxa2PxNum(left) + 'px';
      }

      if (right) {
        styles['margin-right'] = dxa2PxNum(right) + 'px';
      }
    }

    const borderPr = properties[ParagraphPr.border];

    if (borderPr) {
      styles['padding'] = '5px';
      setPBorderStyle(styles, borderPr);
    }

    const spacingPr = properties[ParagraphPr.spacing];

    if (spacingPr) {
      const before = spacingPr[SpacingPr.before];
      const after = spacingPr[SpacingPr.after];
      const line = spacingPr[SpacingPr.line];

      if (before) {
        styles['margin-top'] = dxa2PxNum(before) + 'px';
      }

      if (after) {
        styles['margin-bottom'] = dxa2PxNum(after) + 'px';
      }

      if (line) {
        styles['line-height'] = dxa2PxNum(line) + 'px';
      }
    }

    if (!styles['line-height']) {
      styles['line-height'] = '1.8';
    }

    const textAlign = properties[ParagraphPr.textAlign];

    if (textAlign) {
      styles['text-align'] = textAlign;
    }

    const numPr = properties[ParagraphPr.numPr];

    if (numPr) {
      let numObj = docx.setNumLevel(numPr[NumPr.numId], numPr[NumPr.level]);

      if (numObj) {
        let numElem = htDoc.createElement('span');
        let numStyle = {};
        numStyle['padding-left'] = dxa2PxNum(numObj.indentLeft) + 'px';
        numStyle['text-align'] = 'right';
        numElem.setAttribute('style', obj2styleString(numStyle));
        numElem.appendChild(htDoc.createTextNode(numObj.text));
        htElem.appendChild(numElem);
        htElem.appendChild(htDoc.createTextNode('\u00A0\u00A0'));
      }
    }

    let styleStr = obj2styleString(styles);

    if (styleStr) {
      htElem.setAttribute('style', styleStr);
    }

    if (!this.children || this.children.length <= 0) {
      let span = htDoc.createElement('span');
      htElem.appendChild(span);
    } else {
      buildChildren(htElem, this.children, docx);
    }

    return htElem;
  }

}

exports.Paragraph = Paragraph;
const TblBorderPr = {
  'top': 'top',
  //object BorderPr
  'bottom': 'bottom',
  'left': 'left',
  'right': 'right',
  'insideH': 'insideH',
  'insideV': 'insideV'
};
exports.TblBorderPr = TblBorderPr;
const TblCellMarginPr = {
  'top': 'top',
  //number dxa
  'bottom': 'bottom',
  'left': 'left',
  'right': 'right'
};
exports.TblCellMarginPr = TblCellMarginPr;
const TblFloatPr = {
  'horizontalAnchor': 'horizontalAnchor',
  //string
  'verticalAnchor': 'verticalAnchor' //string

};
exports.TblFloatPr = TblFloatPr;
const TablePr = {
  'styleId': 'styleId',
  //string
  'width': 'width',
  // number dxa（percent not implement)
  'alignment': 'alignment',
  //string left,center,right
  'border': 'border',
  //object TblBorderPr
  'cellMargin': 'cellMargin',
  //object TblCellMarginPr
  'cellSpacing': 'cellSpacing',
  // number dxa
  'floatPr': 'floatPr' //object TblFloatPr

};
exports.TablePr = TablePr;

class Table extends HTComponent {
  constructor(parent, properties, colsWidth) {
    super(parent, properties);
    this.colsWidth = colsWidth;
  }

  toHTElem(htDoc, docx) {
    let htElem = htDoc.createElement('table');
    let properties = this.properties || {};
    let width = properties[TablePr.width];
    let alignment = properties[TablePr.alignment];
    let border = properties[TablePr.border];
    let cellMargin = properties[TablePr.cellMargin];

    if (width) {
      htElem.setAttribute('width', dxa2PxNum(width));
    }

    htElem.setAttribute('border', 1);

    if (cellMargin) {
      let cellPadding = cellMargin[TblCellMarginPr.top] || cellMargin[TblCellMarginPr.bottom];

      if (cellPadding) {
        htElem.setAttribute('cellpadding', dxa2PxNum(cellPadding) + 'px');
      }
    }

    let styles = {};
    styles['border-collapse'] = 'collapse';

    if (alignment === 'center') {
      styles['margin'] = 'auto';
    }

    const floatPr = properties[TablePr.floatPr];

    if (floatPr) {
      styles['float'] = 'right';
      styles['margin-top'] = '30px';
    }

    let styleStr = obj2styleString(styles);

    if (styleStr) {
      htElem.setAttribute('style', styleStr);
    } //calc rowspan


    let tableRows = this.children || [];

    for (let i = 0; i < tableRows.length; i++) {
      const tableRow = tableRows[i];

      if (!Array.isArray(tableRow.children)) {
        continue;
      }

      for (let j = 0; j < tableRow.children.length; j++) {
        const tableCell = tableRow.children[j];

        if (!tableCell || !tableCell.properties) {
          continue;
        }

        if (tableCell.properties[TableCellPr.vMerge] !== 'restart') {
          continue;
        }

        let rowspan = 1;

        for (let k = i + 1; i < tableRows.length; k++) {
          const belowRow = tableRows[k];
          let belowCell = null;

          if (!belowRow || !belowRow.children) {
            break;
          }

          belowCell = belowRow.children[j];

          if (!belowCell) {
            break;
          }

          let vMerge = belowCell.properties ? belowCell.properties[TableCellPr.vMerge] : null;

          if (vMerge === 'continue') {
            rowspan += 1;
            continue;
          }

          break;
        }

        if (rowspan > 1) {
          //console.log('rowspan=', rowspan);
          tableCell.properties[TableCellPr.rowSpan] = rowspan;
        }
      }
    }

    buildChildren(htElem, this.children, docx);
    return htElem;
  }

}

exports.Table = Table;
const TableRowPr = {
  'tblHeader': 'tblHeader',
  //boolean
  'height': 'height',
  //number dxa
  'hAlign': 'hAlign' //left,center,right

};
exports.TableRowPr = TableRowPr;

class TableRow extends HTComponent {
  constructor(parent, properties) {
    super(parent, properties);
  }

  toHTElem(htDoc, docx) {
    let tag = 'tr';
    let properties = this.properties || {};
    let tblHeader = properties[TableRowPr.tblHeader];
    let height = properties[TableRowPr.height];
    let hAlign = properties[TableRowPr.hAlign];

    if (tblHeader) {
      tag = 'th';
    }

    let htElem = htDoc.createElement(tag);
    let styles = {};

    if (height) {
      styles['height'] = dxa2PxNum(height) + 'px';
    }

    styles['vertical-align'] = 'top';
    let styleStr = obj2styleString(styles);

    if (styleStr) {
      htElem.setAttribute('style', styleStr);
    }

    buildChildren(htElem, this.children, docx);
    return htElem;
  }

}

exports.TableRow = TableRow;
const TableCellPr = {
  'width': 'width',
  //number dxa
  'vAlign': 'vAlign',
  //string top center bottom
  'gridSpan': 'gridSpan',
  //number
  'vMerge': 'vMerge',
  //string restart, continue
  'rowSpan': 'rowSpan' //number

};
exports.TableCellPr = TableCellPr;

class TableCell extends HTComponent {
  constructor(parent, properties) {
    super(parent, properties);
  }

  toHTElem(htDoc, docx) {
    let properties = this.properties || {};
    let width = properties[TableCellPr.width];
    let vAlign = properties[TableCellPr.vAlign];
    let gridSpan = properties[TableCellPr.gridSpan];
    let vMerge = properties[TableCellPr.vMerge];
    let rowSpan = properties[TableCellPr.rowSpan];

    if (vMerge === 'continue') {
      return null;
    }

    let htElem = htDoc.createElement('td');

    if (gridSpan) {
      htElem.setAttribute('colspan', gridSpan);
    }

    if (rowSpan) {
      htElem.setAttribute('rowspan', rowSpan);
    }

    let styles = {};

    if (width) {
      styles['width'] = dxa2PxNum(width) + 'px';
    }

    if (vAlign) {
      styles['vertical-align'] = vAlign === 'center' ? 'middle' : vAlign;
    }

    let styleStr = obj2styleString(styles);

    if (styleStr) {
      htElem.setAttribute('style', styleStr);
    }

    buildChildren(htElem, this.children, docx);
    return htElem;
  }

}

exports.TableCell = TableCell;