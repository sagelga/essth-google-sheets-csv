/* To view a documentation on how to use each function visit : https://sagelga.github.io/approval-google-addons/ */

/**
 * Converts number into a character. Perfect for converting a column number into 'A1' format.
 *
 * @tutorial https://stackoverflow.com/questions/21229180/convert-column-index-into-corresponding-column-letter
 *
 * @param {String|Number} column - can be alphabet or number
 *
 * @returns {String} result, as English alphabet. Starts with A as 1.
 *
 * @example
 * columnToLetter(2);
 * => 'B'
 *
 * @example
 * columnToLetter('A');
 * => 'A'
 *
 */
function columnToLetter(column) {
  if (RegExp("[0-9]").test(column)) {
    var temp = "";
    var letter = "";

    while (column > 0) {
      temp = (column - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = (column - temp - 1) / 26;
    }

    return letter;
  }
  return column;
}

/**
 * Converts alphabet to column number
 * @param {*} letter
 */
function letterToColumn(letter) {
  //  Code source : https://stackoverflow.com/questions/21229180/convert-column-index-into-corresponding-column-letter
  var column = 0;
  var length = letter.length;

  for (var i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }

  return column;
}

/**
 * Converts 1 dimension arrays (normal ones) to String
 * by dividing it via ', ' to separate the string values in different array list
 * @param {Array} array - 1 dimension array. nothing too fancy.
 * @returns {String} as comma separated
 *
 * @example
 * arrayToComma(['1', '2', '3'])
 * => '1, 2, 3'
 */
function arrayToComma(arrays) {
  var result = "";

  for (var i = 0; i < arrays.length; i += 1) {
    result += i < arrays.length - 1 ? arrays[i] + ", " : arrays[i];
  }

  return result;
}

/**
 * Converts Google Range to Array
 * @deprecated as they are not being used anymore
 * @tutorial https://developers.google.com/apps-script/reference/spreadsheet/range
 * @param {Range} range - (Must be Google Range datatype)
 */
function rangeToArray(range) {
  var result = [];

  for (var i = 1; i < range.length; i++) {
    var rowResult = "";
    for (var j = 1; j < range[i].length; j++) {
      if (range[i][j]) {
        rowResult = rowResult + ", " + range[i][j];
      }
    }
    result.append(rowResult);
  }

  return result;
}

/**
 * Converts millisecond (from time difference calculations) to Dd HHh MMm
 * @tutorial https://stackoverflow.com/questions/8528382/javascript-show-milliseconds-as-dayshoursmins-without-seconds
 * @param {*} t
 * @returns {String} time string from milisecond calculations
 */
function milisecToTime(t) {
  var cd = 24 * 60 * 60 * 1000,
    ch = 60 * 60 * 1000,
    d = Math.floor(t / cd),
    h = Math.floor((t - d * cd) / ch),
    m = Math.round((t - d * cd - h * ch) / 60000),
    pad = function(n) {
      return n < 10 ? "0" + n : n;
    };
  if (m === 60) {
    h++;
    m = 0;
  }
  if (h === 24) {
    d++;
    h = 0;
  }

  return "" + d + "d " + pad(h) + "h " + pad(m) + "m";
  // return [d, pad(h), pad(m)].join(':');
}
