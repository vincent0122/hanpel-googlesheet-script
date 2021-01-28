function fourth() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var activeSheet = ss.getActiveSheet();
  var aa = ss.getSheetByName("입력");
  var bb = ss.getSheetByName("재고현황");
  var lastRow2 = aa.getLastRow();
  var lastRow = bb.getLastRow();

  //여기서부터 사용품 사용내역 단가구하기
  for (var i = 3; i <= lastRow2; i++) {
    var stockTep = 0;
    var unitFin = 0;
    var unit2 = 0;
    var stock2 = 0;
    var sum = 0;
    var sum2 = 0;

    tarRow = new Array();
    oldSto = new Array();
    oldUni = new Array();
    unit = new Array();

    if (aa.getRange(i, 20).getValue() != "") {
      var itemNam = aa.getRange(i, 20).getValue();
      for (var o = 4; o <= lastRow; o++) {
        if (bb.getRange(o, 4).getValue() === itemNam) {
          tarRow.push(o);
          var v = tarRow.length - 1;
          if (bb.getRange(o, 4).getValue() != bb.getRange(o + 1, 4).getValue())
            break;
        }
      }
      for (k = 0; k < tarRow.length; k++) {
        var sum = sum + bb.getRange(tarRow[k], 5).getValue();
      }

      if (aa.getRange(i, 23).getValue() > sum) {
        SpreadsheetApp.getUi().alert("사용재고가 현재고 보다 많습니다");
        continue;
      }

      for (p = 0; p < tarRow.length; p++) {
        oldSto.push(bb.getRange(o - p, 5).getValue());
        oldUni.push(bb.getRange(o - p, 9).getValue());
      }

      var useSto = aa.getRange(i, 19).getValue();
      var useUnit = aa.getRange(i, 20);

      //1단계 - 한행에서 끝나는 경우

      if (useSto <= oldSto[0]) {
        var stock = oldSto[0] - useSto;
        var unitFin = oldUni[0];
        bb.getRange(o, 5).setValue(stock);
        aa.getRange(i, 24).setValue(unitFin);
      }

      //2단계 - 여러행이 필요한 경우
      if (useSto > oldSto[0]) {
        var sum2 = oldSto[0];
        for (k = 1; k < oldSto.length; k++) {
          var sum2 = sum2 + oldSto[k];
          if (useSto < sum2) {
            var k = k + 1;
            break;
          } //k가 1이면 두개의 행으로 충분하다는 뜻. k가 2면 세개의 행이 필요.
        }

        for (l = 0; l < k - 1; l++) {
          //k가 3이라는 가정. 즉 네개의 행이 필요하다는 뜻임 [0] + [1] + [2] + [3] 했다는 것
          unit[l] = oldUni[l] * oldSto[l]; // unit의 배열에 세개의 행의 amount를 배열로 넣었다.
          var stockTep = stockTep + oldSto[l]; // stockTep 변수에 세개의 행의 재고를 더했다
          var unitFin = unitFin + unit[l]; // unitFin 변수에 세개의 행의 amount를 다 더했다.
        }

        var stock2 = stockTep + oldSto[k - 1] - useSto; // stock 변수에 세개행의 재고에다가 마지막행의 재고를 더하고, 사용한 재고량을 뺐다. ---> 그 칸에 들어가야 할 잔량이지
        var unit2 = oldUni[k - 1] * (oldSto[k - 1] - stock2); //unit2 변수에 마지막행의 단가 * (마지막행의 수량 - stock) 을 했다. ----> 사용한 만큼의 amount를 구한것이다
        var unitFin2 = (unitFin + unit2) / useSto; // unitFin에는 전체 amount를 다 더하고 그것을 사용한 재고량으로 나누었다.
        bb.getRange(o - k + 1, 5).setValue(stock2);
        aa.getRange(i, 24).setValue(unitFin2);

        //여기서부터 행 조정

        bb.deleteRows(o - l + 1, l);

        //여기까지 행 조정
      }
    }
  }
}
