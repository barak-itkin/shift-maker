function copyFormatting(
  sourceRange, destRange, ruleConditionSetter
) {
  var dest = destRange.getSheet();
  var rules = dest.getConditionalFormatRules();
  
  var rowCount = sourceRange.getHeight();
  var colCount = sourceRange.getWidth();
  
  for (var i = 0; i < rowCount; i++) {
    for (var j = 0; j < colCount; j++) {
      var cell = sourceRange.getCell(i + 1, j + 1);
      var bg = cell.getBackground();
      var fg = cell.getFontColor();
      var ruleBuilder = SpreadsheetApp.newConditionalFormatRule().setBackground(
        bg
      ).setFontColor(
        fg
      ).setRanges(
        [destRange]
      );
      ruleBuilder = ruleConditionSetter(ruleBuilder, cell.getValue());
      rules.push(ruleBuilder.build());
    }
  }
  dest.setConditionalFormatRules(rules);
}

function copyFormattingSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('CopyFormattingSidebar').setTitle('Copy Formatting');
  SpreadsheetApp.getUi().showSidebar(html);
}

function copyFormattingSidebarCallback(source, dest) {
  try {
    var sourceRange = SpreadsheetApp.getActive().getRange(source);
  } catch(e) {
      SpreadsheetApp.getUi().alert(
        "ERROR: Failed getting the source range (" + JSON.stringify(source) + ")"
      );
      return;
  }

  try {
    var destRange = SpreadsheetApp.getActive().getRange(dest);
  } catch(e) {
      SpreadsheetApp.getUi().alert(
        "ERROR: Failed getting the destination range (" + JSON.stringify(dest) + ")"
      );
      return;
  }

  function ruleConditionSetter(ruleBuilder, value) {
    return ruleBuilder.whenTextContains(value)
  }

  try {
    copyFormatting(
      sourceRange, destRange, ruleConditionSetter
    );
  } catch(e) {
      SpreadsheetApp.getUi().alert(
        "ERROR: Failed copying the formatting: " + e.message
      );
      return;
  }
}
