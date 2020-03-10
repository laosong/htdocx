'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.docx2ht = exports.ht2docx = exports.IConvtHelper = exports.helper = void 0;

const helper = require('./helper');

exports.helper = helper;

const convt = require('./convt');

const IConvtHelper = convt.IConvtHelper;
exports.IConvtHelper = IConvtHelper;

const ht2docx = require('./ht2docx');

exports.ht2docx = ht2docx;

const docx2ht = require('./docx2ht');

exports.docx2ht = docx2ht;