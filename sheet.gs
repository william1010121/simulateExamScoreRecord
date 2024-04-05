function checkIsChooseType(range) {
  var validationRule = range.getDataValidation();
  if( validationRule != null ) {
    var criterialType = validationRule.getCriteriaType();
    return criterialType == SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST;
  }
}
function columnLetterToNumber(columnLetter) {
  var columnNumber = 0;
  var length = columnLetter.length;

  for( var i = 0; i < length; i++ ) {
    columnNumber *= 26;
    columnNumber += (columnLetter.charCodeAt(i)-"A".charCodeAt(0)) + 1;
  }
  return columnNumber;
}
function getStartPosition(sheet) {
  var returnDicOf = function(startColumnPosition, rowCountPosition) {
    return {
      "startColumn": columnLetterToNumber( sheet.getRange(startColumnPosition).getValue() ),
      "rowCount" : sheet.getRange(rowCountPosition).getValue()
    };
  };
  return {
    "數甲" : returnDicOf("Y5", "Z5"),
    "物理" : returnDicOf("Y7", "Z7"),
    "化學" : returnDicOf("Y9", "Z9")
  };
}
function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();

  if( checkIsChooseType(range) ) {
    Logger.log(getStartPosition(sheet));
    var startPosition = getStartPosition(sheet);

    var row = range.getRow();
    var col = range.getColumn();
    var value = range.getValue();
    Logger.log(value);
    var values = sheet.getRange(row, startPosition[value]["startColumn"]+1, 1, startPosition[value]["rowCount"]).getValues();
    sheet.getRange(row,col+1, 1, startPosition[value]["rowCount"]).setValues(values);
  }
}
