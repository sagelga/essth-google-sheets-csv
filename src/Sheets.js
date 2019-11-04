/* To view a documentation on how to use each function visit : https://sagelga.github.io/approval-google-addons/ */

/**
 *
 * @param {*} columnNumber
 * @param {*} rowNumber
 * @param {*} sheetObject
 */
function getCellValue(columnNumber, rowNumber, sheetObject) {
  try {
    columnNumber = columnToLetter(columnNumber);

    return sheetObject.getRange(columnNumber + "" + rowNumber).getValue();
  } catch (error) {
    console.error("getCellValue() : " + error);

    console.log("columnNumber: " + columnNumber);
    console.log("rowNumber: " + rowNumber);
    console.log("sheetObject: " + sheetObject);

    throw new Error(error);
  }
}

/**
 * Retries column number that matches the `searchQuery`
 * @param {String|Number} searchQuery
 * @param {Object} args
 * @param {Number} args.rowNumber
 * @param {Object} args.sheetObject
 */
function getColumnNumber(searchQuery, args) {
  var result =
    getRowRange(args.rowNumber, args.sheetObject).indexOf(searchQuery) + 1;

  if (result) {
    return result;
  } else {
    return null;
  }
}

/**
 *
 * @param {Number} rowNumber
 * @param {Object} sheetObject
 */
function getRowRange(rowNumber, sheetObject) {
  var result = sheetObject
    .getRange(
      "A" +
        rowNumber +
        ":" +
        columnToLetter(sheetObject.getLastColumn()) +
        rowNumber
    )
    .getValues()[0];

  return result;
}

/**
 *
 * @param {String|Number} columnNumber
 * @param {Object} sheetObject
 */
function getColumnRange(columnNumber, sheetObject) {
  columnNumber = columnToLetter(columnNumber);

  var result = sheetObject
    .getRange(
      columnNumber + "1" + ":" + columnNumber + sheetObject.getLastRow()
    )
    .getValues();

  var arr1d = [].concat.apply([], result);

  return arr1d;
}

/**
 * Pull the value from the sheet, and put it into the *response* Object.
 * DEBUG NOTE : If the error of : 'pullValue() : Exception: Range not found'. Make sure that there is a value in the `.cell` thing
 * @param {Object} response
 * @returns {Object} response
 */
function pullValue(response) {
  try {
    var mandatoryColumn = ["status", "timestamp", "comments", "formUrl"];
    for (var level = 1; level <= 2; level++) {
      for (var i = 0; i < mandatoryColumn.length; i++) {
        response["step".concat(level)][
          mandatoryColumn[i]
        ].value = sheet
          .getRange(response["step".concat(level)][mandatoryColumn[i]].cell)
          .getValue();
      }
    }

    mandatoryColumn = ["id", "skipRow", "email", "requestType", "timestamp"];
    for (i = 0; i < mandatoryColumn.length; i++) {
      response.payloads[mandatoryColumn[i]].value = sheet
        .getRange(response.payloads[mandatoryColumn[i]].cell)
        .getValue();
    }

    response.step1.email.value = sheet
      .getRange(response.step1.email.cell)
      .getValue();
    response.payloads.bodyField = getBodyField(response.payloads.rowNumber);

    var deadline = response.payloads.timestamp.value;
    deadline = deadline.setDate(deadline.getDate() + CONFIG.responseTimeout);
    response.payloads.requestTimeout = deadline;

    return response;
  } catch (error) {
    console.error("pullValue() : " + error);
    console.log(response);

    throw new Error(error);
  }
}

/**
 *
 * @param {*} response
 * @param {*} value
 */
function pushValue(response, value) {
  response.value = value;
  setCellValue(response.cell, value, sheet);
}

/**
 *
 * @param {String} a1Notation
 * @param {String} value
 * @param {Object} [sheetObject=sheet]
 */
function setCellValue(a1Notation, value, sheetObject) {
  if (sheetObject === undefined) {
    sheetObject = sheet;
  }

  sheetObject.getRange(a1Notation).setValue(value);
}

/**
 *
 * @param {*} name
 * @param {Object} [args]
 * @param {Number} [args.columnNumber = sheetObject.getLastColumn()]
 * @param {Boolean} [args.appendAfter = true]
 * @param {Sheet} [args.sheetObject = sheet]
 * @param {Boolean} [args.checkExistance = false]
 */
function createNewColumn(name, args) {
  /* Check wheather the column is already exists or not */
  var colNum = getColumnNumber(name, {
    rowNumber: 1,
    sheetObject: args.sheetObject
  });

  if (args.checkExistance && colNum) {
    console.warn(name + "already exists");
    return colNum;
  }

  try {
    /* Setup default values */
    var sheetObject = args.sheetObject || sheet;
    var location = args.columnNumber || sheetObject.getLastColumn();
    var appendAfter = args.appendAfter || true;

    /* Insert a new column at the end of the list */
    if (appendAfter) {
      sheetObject.insertColumnAfter(location);
      location += 1;
    } else {
      sheetObject.insertColumnBefore(location);
    }

    var headerCell = columnToLetter(location) + "" + CONFIG.headerRow;
    setCellValue(headerCell, name, sheetObject);
    console.log(headerCell);
    // console.log("Insertion of " + name + " is completed.");
  } catch (error) {
    console.error("createNewColumn() : " + error);
    console.log("appendAfter: " + appendAfter);
    console.log("location: " + location);
  }
}
