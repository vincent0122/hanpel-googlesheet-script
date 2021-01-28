function sending() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var secondSheet = ss.getSheetByName("공장일지");
  var targetSheet = ss.getSheetByName("원료제품누적");
  var half = ss.getSheetByName("기타제품누적");
  var add = ss.getSheetByName("시간외근무누적");
  var etc = ss.getSheetByName("기타활동누적");
  var date = secondSheet.getRange("C1").getValue();

  var rowRange = "I6:I49";
  var etcRange = "A6:A19";
  var addRange = "A23:A32";
  endRow = [];

  function findRow(range) {
    var rowLen = secondSheet.getRange(range).getValues();
    for (rows = 0; rows < rowLen.length; rows++) {
      if (!rowLen[rows].join("")) {
        endRow.push(rows);
        break;
      }
    }
  }

  findRow(rowRange);
  findRow(etcRange);
  findRow(addRange);

  Logger.log(endRow);
  function inputting(sheetName, colNum, values) {
    var emp = sheetName.getRange("a2:a30000").getValues();
    for (row = 0; row < emp.length; row++) {
      if (!emp[row].join("")) break;
    }
    sheetName.getRange(row + 2, 1, 1).setValue(date);
    sheetName.getRange(row + 2, 2, 1, colNum).setValues(values);
  }

  // 원료, 제품 입력
  for (i = 6; i < endRow[0] + 6; i++) {
    var data1 = secondSheet.getRange(i, 9, 1, 11).getValues();
    if (data1[0][0] === "원료" || data1[0][0] === "제품") {
      inputting(targetSheet, 11, data1);
    } else if (data1[0][0] === "기타" || data1[0][0] === "지대") {
      inputting(half, 11, data1);
    }
  }
  // 기타활동누적 입력
  for (i = 6; i < endRow[1] + 6; i++) {
    var data3 = secondSheet.getRange(i, 1, 1, 7).getValues();
    inputting(etc, 7, data3);
  }

  // 시간외근무 입력
  for (i = 23; i < endRow[2] + 23; i++) {
    var data2 = secondSheet.getRange(i, 1, 1, 7).getValues();
    inputting(add, 7, data2);
  }
}
