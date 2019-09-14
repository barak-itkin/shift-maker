function buildSchedule(
  dataRange, startTime, rooms, filters,
  textColHeader, startColHeader, lengthColHeader, roomColHeader,
  destRange
) {
  var data = dataRange.getValues();
  
  var headers = data[0];
  var headersIndex = array2index(headers);
  
  var roomIndex = array2index(rooms);
  
  var result = buildSkeleton(rooms, getHourFrom(startTime));
  var merges = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    var match = true;
    for (var field in filters) {
      if (row[headersIndex[field]] != filters[field]) {
        match = false;
        break;
      }
    }
    if (!match) {
      continue;
    }
    
    start = getHourFrom(row[headersIndex[startColHeader]]);
    length = getDurationFrom(row[headersIndex[lengthColHeader]]);
    place = row[headersIndex[roomColHeader]];
    text = row[headersIndex[textColHeader]];
    
    rowIndex = start - getHourFrom(startTime);
    row = (rowIndex < 0 ? rowIndex + 24 : rowIndex) + 1;
    col = roomIndex[place] + 1;
    
    result[row][col] = text;
    if (length > 1) {
      merges.push([row, col, length]);
    }
  }
  
  var destSheet = destRange.getSheet();
  destRange = destSheet.getRange(destRange.getRow(), destRange.getColumn(), result.length, result[0].length);
  removeMerges(destRange);
  destRange.setWrap(true);
  destRange.setValues(result);
  for (var i = 0; i < merges.length; i++) {
    destSheet.getRange(
      merges[i][0] + destRange.getRow(),
      merges[i][1] + destRange.getColumn(),
      merges[i][2], 1
    ).merge();
  }  
}

function buildScheduleSidebarCallback(
  dataRange, startTime, roomsRange, filtersJSON,
  textColHeader, startColHeader, lengthColHeader, roomColHeader,
  destRange
)
{
  buildSchedule(
    SpreadsheetApp.getActive().getRange(dataRange),
    getHourFrom(startTime),
    flatten2D(SpreadsheetApp.getActive().getRange(roomsRange).getValues()),
    JSON.parse(filtersJSON),
    textColHeader,
    startColHeader,
    lengthColHeader,
    roomColHeader,
    SpreadsheetApp.getActive().getRange(destRange)
  );
}

function buildScheduleSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('BuildScheduleSidebar').setTitle('Build Schedule');
  SpreadsheetApp.getUi().showSidebar(html);
}

function buildSkeleton(rooms, startTime) {
  var result = [];
  // Create title row
  var resultHeader = [''];
  for (var i = 0; i < rooms.length; i++) {
    resultHeader.push(rooms[i]);
  }
  result.push(resultHeader);
  // Create the rest of the rows with the time column
  for (var i = 0; i < 24; i++) {
    var row = [(startTime + i) % 24];
    while (row.length < resultHeader.length) {
      row.push('');
    }
    result.push(row);
  }
  return result;
}

function removeMerges(range) {
  var merges = range.getMergedRanges();
  for (var i = 0; i < merges.length; i++) {
    merges[i].breakApart();
  }
}