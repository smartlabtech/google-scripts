function LastNonEmptyCell(sheet, Column) {
  var numberRow = sheet.getMaxRows()
  var lastRow = numberRow
  var counter = 0
  var values = sheet.getRange(1, Column, lastRow, Column).getValues()
  for (; values[lastRow - 1] == "" && lastRow > 0; lastRow--) {
    counter++
  }
  return numberRow - counter
}

function FindInAnArray(Array, item) {
  for (var i = 0; i < Array.length; i++) if (Array[i] == item) return i
  return -1
}

function columnToLetter(column) {
  var temp,
    letter = ""
  while (column > 0) {
    temp = (column - 1) % 26
    letter = String.fromCharCode(temp + 65) + letter
    column = (column - temp - 1) / 26
  }
  return letter
}

function letterToColumn(letter) {
  var column = 0,
    length = letter.length
  for (var i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1)
  }
  return column
}

function deleteRowsWithCondition() {
  var Sheet = (Sheet = SpreadsheetApp.openById(
    "1OwFAg_TkuYTUW2IwSpoZMbTDtsK1gkk2CbiZ027rEFc"
  ).getSheetByName("Need To be Confirmed"))
  var startRow = 1 // Start at second row because the first row contains the data labels
  var numRows = LastNonEmptyCell(Sheet, 1) // Put in here the number of rows you want to process
  var numCols = 10
  if (numRows > 0) {
    var dataRange = Sheet.getRange(startRow, 1, numRows, numCols).getValues()

    var HeadersRowData = Sheet.getRange(
      1,
      1,
      1,
      Sheet.getMaxColumns()
    ).getValues()
    var InPendingOrders_Col = FindInAnArray(
      HeadersRowData[0],
      "In Pending Orders?"
    )

    for (var i = numRows - 1; i >= 1; i--) {
      if (dataRange[i][InPendingOrders_Col] == 0) Sheet.deleteRow(i + 1)
    }
  }
}

/*
  function onEdit() {
    var rowno= SpreadsheetApp.getActiveRange().getRowIndex();
    var Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("All orders");
    var startRow = 2; // Start at second row because the first row contains the data labels
    var numRows = LastNonEmptyCell(Sheet,1)-1; // Put in here the number of rows you want to process
    var numCols =50; 
    
    var HeadersRowData = Sheet.getRange(1,1,1,Sheet.getMaxColumns()).getValues();
    var updated_at_Col = FindInAnArray(HeadersRowData[0],"Updated Date")+1;
    var Status_Col = FindInAnArray(HeadersRowData[0],"Status");
    var timestamp=Utilities.formatDate(new Date(), "GMT+2", "M/d/yy HH:mm:ss");  
  if(rowno>1)
  var dataRange = Sheet.getRange(startRow,1,numRows,numCols).getValues();
  if(dataRange[rowno-2][Status_Col]!="Delivered")
    Sheet.getRange(rowno, updated_at_Col).setValue(timestamp);
      
    if(SpreadsheetApp.getActiveRange().getValue() == "y") {
      SpreadsheetApp.getActiveRange().setValue('=CHAR(10004)');
      SpreadsheetApp.getActiveRange().setBackgroundRGB(0,255, 0);
      SpreadsheetApp.flush();
    }
    
    if(SpreadsheetApp.getActiveRange().getValue() == "") {
      SpreadsheetApp.getActiveRange().setBackgroundRGB(255, 255, 255);
    }
  }*/
