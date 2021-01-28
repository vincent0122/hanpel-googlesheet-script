function third() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var activeSheet = ss.getActiveSheet();
  var aa = ss.getSheetByName("입력");
  var bb = ss.getSheetByName("재고현황");
  var now = new Date();
  var month = now.getMonth();
  var day = 1;
  var date = month + "/" + day;

  for (var i = 3; i < 50; i++) {
    var lastRow = bb.getLastRow();
    var valName = aa.getRange(i, 20).getValue();
    var valPlu = aa.getRange(i, 21).getValue();
    var valUni = aa.getRange(i, 22).getValue();

    if (valName && valPlu && valUni != "") {
      // 제품명이랑 수입단가가 공란이 아니면 시작!
      var purDat = aa.getRange(i, 20, 1, 3).getValues();
      purDat[0].splice(0, 0, "생산", date, "");
      purDat[0].splice(5, 0, "", "", "");
      purDat[0].splice(9, 0, "1");

      for (var o = 4; o < lastRow + 1; o++) {
        purDat[0][9] = "=E" + o + "*" + "I" + o;

        if (bb.getRange(o, 4).getValue() === valName) {
          bb.insertRowBefore(o);
          bb.getRange(o, 1, 1, 10).setValues(purDat);
          break;
        } else if (o === lastRow) {
          bb.getRange(lastRow + 1, 1, 1, 10).setValues(purDat);
        }
      }
    }
    Logger.log(lastRow);
    Logger.log(valName);
    Logger.log(valPlu);
    Logger.log(valUni);

    if (valName === "") break;
  }
}
