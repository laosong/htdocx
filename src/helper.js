/**
 * Created by laosong on 2020/2/18.
 * some helper function
 */

'use strict';

const translate_re = /&(nbsp|amp|quot|lt|gt);/g;
const translate = {
  'nbsp': ' ',
  'amp': '&',
  'quot': '"',
  'lt': '<',
  'gt': '>'
};

export function decodeEntities(encodedString) {
  return encodedString.replace(translate_re, function(match, entity) {
    return translate[entity];
  });
}

/* eslint no-constant-condition:0 */
const findEndOfMath = function(delimiter, text, startIndex) {
  // Adapted from
  // https://github.com/Khan/perseus/blob/master/src/perseus-markdown.jsx
  let index = startIndex;
  let braceLevel = 0;

  const delimLength = delimiter.length;

  while (index < text.length) {
    const character = text[index];

    if (braceLevel <= 0 &&
      text.slice(index, index + delimLength) === delimiter) {
      return index;
    } else if (character === '\\') {
      index++;
    } else if (character === '{') {
      braceLevel++;
    } else if (character === '}') {
      braceLevel--;
    }

    index++;
  }

  return -1;
};

export function splitAtDelimiters(startData, leftDelim, rightDelim, display) {
  const finalData = [];

  if (startData.type !== 'text') {
    return finalData;
  }

  const text = startData.data;

  let lookingForLeft = true;
  let currIndex = 0;
  let nextIndex;

  nextIndex = text.indexOf(leftDelim);
  if (nextIndex !== -1) {
    currIndex = nextIndex;
    finalData.push({
      type: 'text',
      data: text.slice(0, currIndex),
    });
    lookingForLeft = false;
  }

  while (true) {
    if (lookingForLeft) {
      nextIndex = text.indexOf(leftDelim, currIndex);
      if (nextIndex === -1) {
        break;
      }

      finalData.push({
        type: 'text',
        data: text.slice(currIndex, nextIndex),
      });

      currIndex = nextIndex;
    } else {
      nextIndex = findEndOfMath(rightDelim, text, currIndex + leftDelim.length);
      if (nextIndex === -1) {
        break;
      }

      finalData.push({
        type: 'math',
        data: text.slice(currIndex + leftDelim.length, nextIndex),
        rawData: text.slice(currIndex, nextIndex + rightDelim.length),
        display: display,
      });

      currIndex = nextIndex + rightDelim.length;
    }

    lookingForLeft = !lookingForLeft;
  }

  finalData.push({
    type: 'text',
    data: text.slice(currIndex),
  });

  return finalData;
}

export function replaceMathML(htText, startId, maths) {
  let result;

  const texts = splitAtDelimiters({type: 'text', data: htText}, '<math>', '</math>', false);

  let mathId = startId;

  for (let i = 0; i < texts.length; i++) {
    const textB = texts[i];

    if (textB.type === 'text') {
      result += textB.data;
    } else if (textB.type === 'math') {
      const k = '@math-ml-' + mathId + '@';

      maths[k] = {mathml: textB.rawData};

      mathId += 1;
      result += k;
    }
  }

  return result;
}

export function visitNode(node, callback) {
  if (callback(node)) {
    return true;
  }
  if (node = node.firstChild) {
    do {
      if (visitNode(node, callback)) {
        return true
      }
    } while (node = node.nextSibling)
  }
}

export function replaceLatex(htDoc, startId, maths) {
  let mathId = startId;

  visitNode(htDoc.documentElement, function(node) {
    if (node.nodeType === 3) {//text
      let result;

      const startData = {type: 'text', data: node.textContent};

      const texts = splitAtDelimiters(startData, '$', '$', false);

      for (let i = 0; i < texts.length; i++) {
        const textB = texts[i];

        if (textB.type === 'text') {
          result += textB.data;
        } else if (textB.type === 'math') {
          const k = '@math-la-' + mathId + '@';

          maths[k] = {latex: textB.data};

          mathId += 1;
          result += k;
        }
      }
      node.textContent = result;
    }
  });
}

export function upperLetterOfNum(num) {
  if (isNaN(num))
    return NaN;
  return String.fromCharCode(65 + (num - 1));
}

export function lowerLetterOfNum(num) {
  if (isNaN(num))
    return NaN;
  return String.fromCharCode(97 + (num - 1));
}

export function upperRomanOfNum(num) {
  if (isNaN(num))
    return NaN;
  let key = ['', 'C', 'CC', 'CCC', 'CD', 'D', 'DC', 'DCC', 'DCCC', 'CM',
    '', 'X', 'XX', 'XXX', 'XL', 'L', 'LX', 'LXX', 'LXXX', 'XC',
    '', 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX'];
  let digits = String(+num).split(''),
    roman = '',
    i = 3;
  while (i--)
    roman = (key[+digits.pop() + (i * 10)] || '') + roman;
  return Array(+digits.join('') + 1).join('M') + roman;
}

export function lowerRomanOfNum(num) {
  if (isNaN(num))
    return NaN;
  let key = ['', 'c', 'cc', 'ccc', 'cd', 'd', 'dc', 'dcc', 'dccc', 'cm',
    '', 'x', 'xx', 'xxx', 'xl', 'l', 'lx', 'lxx', 'lxxx', 'xc',
    '', 'i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix'];
  let digits = String(+num).split(''),
    roman = '',
    i = 3;
  while (i--)
    roman = (key[+digits.pop() + (i * 10)] || '') + roman;
  return Array(+digits.join('') + 1).join('M') + roman;
}

export function simpleChineseOfNum(num) {
  if (isNaN(num))
    return NaN;
  const digits = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九'];
  const positions = ['', '十', '百', '千', '万', '十万', '百万', '千万', '亿', '十亿', '百亿', '千亿'];
  const charArray = String(num).split('');
  let result = '';
  let prevIsZero = false;
  //处理0  deal zero
  for (let i = 0; i < charArray.length; i++) {
    const ch = charArray[i];
    if (ch !== '0' && !prevIsZero) {
      result += digits[parseInt(ch)] + positions[charArray.length - i - 1];
    } else if (ch === '0') {
      prevIsZero = true;
    } else if (ch !== '0' && prevIsZero) {
      result += '零' + digits[parseInt(ch)] + positions[charArray.length - i - 1];
    }
  }
  //处理十 deal ten
  if (num < 100) {
    result = result.replace('一十', '十');
  }
  return result;
}

export const px2ptTab = {
  '9': '7',
  '10': '7.5',
  '11': '8.5',
  '12': '9',
  '13': '10',
  '14': '10.5',
  '15': '11.5',
  '16': '12',
  '17': '13',
  '18': '13.5',
  '19': '14.5',
  '20': '15',
  '21': '16',
  '22': '16.5',
  '23': '17.5',
  '24': '18',
  '25': '19',
  '26': '19.5',
  '27': '20.5',
  '28': '21',
  '29': '22',
  '30': '22.5',
  '31': '23.5',
  '32': '24',
  '33': '25',
};

export const sz2pxTab = {
  '14': '9',
  '15': '10',
  '16': '10',
  '17': '11',
  '18': '12',
  '19': '12',
  '20': '13',
  '21': '14',
  '22': '14',
  '23': '15',
  '24': '16',
  '25': '16',
  '26': '17',
  '27': '18',
  '28': '18',
  '29': '19',
  '30': '20',
  '31': '20',
  '32': '21',
  '33': '22',
  '34': '22',
  '35': '23',
  '36': '24',
  '37': '24',
  '38': '25',
  '39': '26',
  '40': '26',
  '41': '27',
  '42': '28',
  '43': '28',
  '44': '29',
  '45': '30',
  '46': '30',
  '47': '31',
  '48': '32',
  '49': '32',
  '50': '33',
};
