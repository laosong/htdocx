/**
 * Created by laosong on 2018/7/5.
 */

'use strict';

const {describe} = require('mocha');
const {assert} = require('chai');

const xmlDOM = require('xmldom');

const {convt, docx2ht} = require('../src/index');

let fs = require('fs');
let path = require('path');

describe('docx2ht', () => {

  describe('parseDocx', () => {
    it('should check success', async function() {

      console.log(convt, docx2ht);

      const fp = 'D:\\Work\\demo\\docx\\20200307\\1583380543247\\word\\document.xml';
      const xml = fs.readFileSync(fp, {encoding: 'utf-8'});

      let xmls = {'document': xml};

      let result = docx2ht.parseDocx(xmls);

      console.log(result);

      let htDocImpl = new xmlDOM.DOMImplementation().createDocument(null, 'htXML', null);

      result = await docx2ht.createHT(result, htDocImpl);

      console.log(result);

      return 0;
    }).timeout(10000);
  });


});
