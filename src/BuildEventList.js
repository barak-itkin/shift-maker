function buildEventList(
  dataRanges, nameHeader, colHeader, rowHeader, lengthHeader, contentHeader,
  destRange
) {
  var result = [
    [nameHeader, colHeader, rowHeader, lengthHeader, contentHeader]
  ];

  for (var rangeName in dataRanges) {
    var dataRange = dataRanges[rangeName];
    var merges = dataRange.getMergedRanges();
    var data = dataRange.getValues();
    
    var lengths = [];
    for (var i = 0; i < data.length; i++) {
      var row = [];
      for (var j = 0; j < data[i].length; j++) {
        row.push(1);
      }
      lengths.push(row);
    }
    for (var k = 0; k < merges.length; k++) {
      var merge = merges[k];
      for (var i = 0; i < merge.getHeight(); i++) {
        for (var j = 0; j < merge.getWidth(); j++) {
          lengths[merge.getRow() + i - 1][merge.getColumn() + j - 1] = merge.getHeight();
        }
      }
    }

    for (var i = 1; i < data.length; i++) {
      for (var j = 1; j < data[i].length; j++) {
        if (!data[i][j]) {
          continue;
        }
        result.push([
          rangeName, data[0][j], data[i][0], lengths[i][j], data[i][j]
        ]);
      }
    }
  }

  var destSheet = destRange.getSheet();
  destRange = destSheet.getRange(
    destRange.getRow(), destRange.getColumn(), result.length, result[0].length
  );
  removeMerges(destRange);
  destRange.setWrap(true);
  destRange.setValues(result);
}

function buildEventListSidebarCallback(
  dataRanges, nameHeader, colHeader, rowHeader, lengthHeader, contentHeader,
  destRange
)
{
  try {
    dataRanges = JSON.parse(dataRanges);
  } catch(e) {
    SpreadsheetApp.getUi().alert(
      "ERROR: Failed parsing the JSON of the data ranges: " + e.message
    );
    return;
  }

  if (dataRanges.constructor != ({}).constructor) {
    SpreadsheetApp.getUi().alert(
      "ERROR: The JSON of the data ranges is not an object, it's a " + (typeof dataRanges)
    );
    return;
  }

  var realDataRanges = {};
  for (var name in dataRanges) {
    try {
      realDataRanges[name] = SpreadsheetApp.getActive().getRange(dataRanges[name]);
    } catch(e) {
      SpreadsheetApp.getUi().alert(
        "ERROR: Failed getting the data range for " + JSON.stringify(name)
      + " (which is " + JSON.stringify(dataRanges[name]) + ")"
      );
      return;
    }
  }

  try {
    var destRangeObj = SpreadsheetApp.getActive().getRange(destRange);
  } catch(e) {
      SpreadsheetApp.getUi().alert(
        "ERROR: Failed getting the destination range (" + JSON.stringify(destRange) + ")"
      );
      return;
  }

  try {
    buildEventList(
      realDataRanges,
      nameHeader,
      colHeader,
      rowHeader,
      lengthHeader,
      contentHeader,
      destRangeObj
    );
  } catch(e) {
      SpreadsheetApp.getUi().alert(
        "ERROR: Failed building an event list: " + e.message
      );
      return;
  }
}

function buildEventListSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('BuildEventListSidebar').setTitle('Build Event List');
  SpreadsheetApp.getUi().showSidebar(html);
}
