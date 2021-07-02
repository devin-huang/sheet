// 过滤对象属性
export const pickOne = (origin, item) => {
  return origin.reduce((prev, key) => {
    if (item[key]) {
      return { ...prev, ...{ [key]: item[key] } };
    }
    return prev;
  }, {});
};

// 合并封装
export const assignAll = (origin, ...other) => {
  if (!origin) origin = {};
  const result = Object.assign.call(this, origin, ...other);
  return result;
};

/**
 * 数字转换成excel表头。 （递归处理）
 * @param {Number} num 需要转换的数字
 */
// eslint-disable-next-line no-unused-vars
export const numToString = (num) => {
  let strArray = [];
  let numToStringAction = function(o) {
    let temp = o - 1;
    let a = parseInt(temp / 26);
    let b = temp % 26;
    strArray.push(String.fromCharCode(64 + parseInt(b + 1)));
    if (a > 0) {
      numToStringAction(a);
    }
  };
  numToStringAction(num);
  return strArray.reverse().join("");
};

/**
 * 表头字母转换成数字。（进制转换）
 * @param {string} str 需要装换的字母
 */
export const stringToNum = (str) => {
  let temp = str.toLowerCase().split("");
  let len = temp.length;
  let getCharNumber = function(charx) {
    return charx.charCodeAt() - 96;
  };
  let numout = 0;
  let charnum = 0;
  for (let i = 0; i < len; i++) {
    charnum = getCharNumber(temp[i]);
    numout += charnum * Math.pow(26, len - i - 1);
  }
  return numout;
};

/**
 * worksheet转成ArrayBuffer
 * @param {worksheet} s xlsx库中的worksheet
 */
export const s2ab = (s) => {
  if (typeof ArrayBuffer !== "undefined") {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i !== s.length; ++i) {
      view[i] = s.charCodeAt(i) & 0xff;
    }
    return buf;
  } else {
    const buf = new Array(s.length);
    for (let i = 0; i !== s.length; ++i) {
      buf[i] = s.charCodeAt(i) & 0xff;
    }
    return buf;
  }
};

// 是否为对象
export const isPlainObject = (value) =>
  Object.prototype.toString.call(value) === "[object Object]";

// 设置合并单元格
export const setMergeData = (prop) => {
  return (mockData) => {
    return mockData.reduce((prev, cur) => {
      if (Object.prototype.hasOwnProperty.call(prev, cur[prop])) {
        const num = prev[cur[prop]] + 1;
        return { ...prev, ...{ [cur[prop]]: num } };
      } else if (cur[prop]) {
        return { ...prev, ...{ [cur[prop]]: 1 } };
      }
      return prev;
    }, {});
  };
};
