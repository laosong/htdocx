"use strict";

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.OMath = void 0;

// songchunwen add omath
const docx = require('docx');

const xml2Js = require('xml-js');

class OMath extends docx.XmlComponent {
  constructor(text, fontSize) {
    super('m:oMath');
    this.omml = text;
    this.fontSize = fontSize;
  }

  init() {
    this.root = [];
    const ommlObj = xml2Js.xml2js(this.omml, {
      compact: false
    });
    let omathXmlElement = null;

    for (const xmlElm of ommlObj.elements || []) {
      if (xmlElm.name === 'm:oMath') {
        omathXmlElement = xmlElm;
      }
    }

    if (omathXmlElement === undefined) {
      throw new Error('can not find omath element');
    }

    if (this.fontSize) {
      const fontEle = {
        type: 'element',
        name: 'w:rPr',
        elements: [{
          type: 'element',
          name: 'w:sz',
          attributes: {
            'w:val': this.fontSize
          }
        }, {
          type: 'element',
          name: 'w:szCs',
          attributes: {
            'w:val': this.fontSize
          }
        }]
      };
      this.setFont(omathXmlElement, fontEle);
    }

    const omathElements = omathXmlElement.elements || [];
    omathElements.map(childElm => {
      const ixc = docx.convertToXmlComponent(childElm);
      this.root.push(ixc);
    });
  }

  setFont(ele, fontEle) {
    if (!ele) {
      return;
    }

    if (!Array.isArray(ele.elements)) {
      return;
    }

    if (ele.type === 'element' && ele.name === 'm:r') {
      ele.elements.unshift(fontEle);
    }

    for (let eleChild of ele.elements) {
      this.setFont(eleChild, fontEle);
    }
  }

}

exports.OMath = OMath;