function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu(
    'Shift Maker'
  ).addItem(
    'Copy Formatting', 'copyFormattingSidebar'
  ).addItem(
    'Build Schedule', 'buildScheduleSidebar'
  ).addItem(
    'Build Event List', 'buildEventListSidebar'
  ).addToUi();
}

function getActiveRangeA1Notation() {
  var range = SpreadsheetApp.getActiveRange();
  var sheetName = range.getSheet().getSheetName();
  var notation = range.getA1Notation();
  return "'" + sheetName + "'" + '!' + notation;
}

function flatten2D(values) {
  var result = [];
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      result.push(values[i][j]);
    }
  }
  return result;
}

function getHourFrom(value) {
  if (typeof value == 'number') {
    while (value < 0) {
      value += 24;
    }
    return value;
  } else if (typeof value == 'string') {
    return parseInt(value.split(':')[0]);
  } else {
    // Dates not handled yet
  }
}

function getDurationFrom(value) {
  if (typeof value == 'number') {
    return value;
  } else {
    // Dates not handled yet
  }
}

function array2index(values) {
  var result = {};
  for (var i = 0; i < values.length; i++) {
    result[values[i]] = i;
  }
  return result;
}