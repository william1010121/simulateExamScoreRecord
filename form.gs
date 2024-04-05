function tableNameWithUrl () {
  return {
    "110-2-中": "https://drive.google.com/drive/folders/1jqDMmlaBX0Oh2T8VkArKqyNx0tFeajW_",
    "110-1-北": "https://drive.google.com/drive/folders/1CspQmLA3YZTKvPDwBgD1o3TUh-ayPSfo",
    "110-2-北": "https://drive.google.com/drive/folders/1OJVmozur2tUjDbp9sAHLfFGr2F_11YVX",
    "110-1-全": "https://drive.google.com/drive/folders/1fwr3wmU6VBrExJ1lrFE5BR5CQFIq55NR",
    "110-1-全-1": "https://drive.google.com/drive/folders/14eBovFxkpfdN6zxlNZ5DRPO6dP7ij0xN",
    "110-2-全" : "https://drive.google.com/drive/folders/1-1KnKYCC2omwtGD4ChONkDBz4Um7RHDX",
    "110-2-全-1" : "https://drive.google.com/drive/folders/1Q-pVx51WsdXN0Ff58aKlmhgLyZOAik9x",
    "110-3-全" : "https://drive.google.com/drive/folders/1G8nkvNBWyfyqBJXw_hYxnc7y08SUWJIb",

    "111-1-中" : "https://drive.google.com/drive/folders/1Ze37QGP1D8EUln1AwU4xwRcLmZ5WcJmO",
    "111-3-北" : "https://drive.google.com/drive/folders/1079oiA0HBuvrkQKQK4d7YNXAswYpYZLQ",
    "111-1-全" : "https://drive.google.com/drive/folders/1W3yuIZDVgBFrXiuZET5Ulamgs3Lf1mMm",
    "111-2-全" : "https://drive.google.com/drive/folders/1GixfFsuvhfXTqsIMnZliW0u57O3c0ybN",
    "111-3-全" : "https://drive.google.com/drive/folders/1GixfFsuvhfXTqsIMnZliW0u57O3c0ybN"
  };
}
function findOrCreateTableName(tableName,sheet ){
  var urlDic = tableNameWithUrl();

  var startRow = sheet.getRange("Y2").getValue();
  var lastRow = sheet.getLastRow();
  for( let i = startRow; i <= lastRow; ++i ) {
    Logger.log("row:" + i);
    let currentValue = sheet.getRange(i,1).getValue();
    if( currentValue == tableName ) {
      sheet.getRange(i,1).setFormula(`=HYPERLINK( "${urlDic[tableName]}" , "${tableName}")`)
      return i;
    }
    if( currentValue == "") {
      sheet.getRange(i,1).setFormula(`=HYPERLINK( "${urlDic[tableName]}" , "${tableName}")`)
      return i;
    }
  }

  sheet.getRange(lastRow+1,1).setFormula(`=HYPERLINK( "${urlDic[tableName]}" , "${tableName}")`)
  return lastRow+1;
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

function onFormSubmit(e) {
  var url = "https://docs.google.com/spreadsheets/d/1iLLx02ua9b5arzsN975FDh4uFD4MtvcRCAkrwAwp3kQ/edit#gid=1614157971";
  var spreadsheets = SpreadsheetApp.openByUrl(url);
  var sheet = spreadsheets.getSheetByName("模考統整");


  var formResponse = e.response;
  var itemResponses = formResponse.getItemResponses();
  var object = itemResponses[1].getResponse();
  var dataRow = findOrCreateTableName(itemResponses[0].getResponse(),sheet);

  var startPositions = getStartPosition(sheet);
  var startColumn = startPositions[object]["startColumn"];
  var rowCount = startPositions[object]["rowCount"];

  sheet.getRange(dataRow, startColumn+1).setValue(itemResponses[2].getResponse());
  sheet.getRange(dataRow, startColumn+2).setValue(itemResponses[3].getResponse());
  sheet.getRange(dataRow, startColumn+3).setValue(itemResponses[4].getResponse());

}

