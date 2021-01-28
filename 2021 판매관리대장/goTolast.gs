function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var screen = ss.getSheetByName("2021");
  var range = screen.getRange(screen.getLastRow(), 1);

  screen.setActiveRange(range);
}
