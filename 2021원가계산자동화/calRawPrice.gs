function calRawPrice() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var input = ss.getSheetByName("입력");
  var stock = ss.getSheetByName("재고현황");
  var usedData = ss.getSheetByName("원료제품사용내역");

  const getStockInfo = () => {
    var stockInfo = {};
    stockInfo.lastRow = stock.getLastRow();
    stockInfo.data = stock.getRange(4, 1, stockInfo.lastRow, 9).getValues();

    return stockInfo;
  };

  const getInputInfo = () => {
    var inputInfo = {};
    inputInfo.lastRow = input.getLastRow();
    inputInfo.data = input.getRange(3, 1, inputInfo.lastRow, 9).getValues();

    return inputInfo;
  };

  const getUsedInfo = () => {
    var usedInfo = {};
    usedInfo.lastRow = usedData.getLastRow();
    usedInfo.data = usedData.getRange(3, 1, usedInfo.lastRow, 3).getValues();
    usedInfo.data2 = usedData.getRange(3, 5, usedInfo.lastRow, 9).getValues();

    return usedInfo;
  };

  const inputImportData = () => {
    //잘됨
    var stockInfo = getStockInfo();
    var inputInfo = getInputInfo();
    var finalData = stockInfo.data.concat(inputInfo.data);
    stock.getRange(4, 1, finalData.length, 9).setValues(finalData);
    var range = stock.getRange(4, 1, finalData.length, 9);
    range.sort([
      {
        column: 3,
        ascending: true,
      },
      {
        column: 1,
        ascending: true,
      },
    ]);
  };

  const stockBackup = (dataTo) => {
    var stockData = getStockInfo();
    var dataTo = ss.getSheetByName(dataTo);
    dataTo.getRange(1, 1, stockData.data.length, 9).setValues(stockData.data);
  };

  const inputRawPrice = () => {
    var stockInfo = getStockInfo();
    var usedInfo = getUsedInfo();
    var usedUnitPrice = new Array();

    for (var i = 0; i < usedInfo.lastRow; i++) {
      var stockTep = 0;
      var unit2 = 0;
      var stock2 = 0;
      var sum = 0;
      var sum2 = 0;
      var nowUnitPrice = 0;

      var targetRow = new Array();
      var oldStock = new Array();
      var oldUnitPrice = new Array();
      var unit = new Array();

      for (var k = 0; k < stockInfo.data.length - 3; k++) {
        if (usedInfo.data[i][0] === stockInfo.data[k][2]) {
          targetRow.push(k);
          if (stockInfo.data[k][2] != stockInfo.data[k + 1][2]) {
            break;
          }
        }
      }

      for (r = 0; r < targetRow.legnth; r++) {
        var sum = sum + stockInfo.data[k][4];
      }

      if (usedInfo.data[i][0] > sum) {
        SpreadsheetApp.getUi().alert("사용재고가 현재고 보다 많습니다");
        continue;
      }

      for (var q = 0; q < targetRow.length; q++) {
        oldStock.push(stockInfo.data[targetRow[q]][3]),
          oldUnitPrice.push(stockInfo.data[targetRow[q]][7]);
      }

      var usedStock = usedInfo.data[i][1];

      //1단계 - 한행에서 끝나는 경우
      if (usedStock <= oldStock[0]) {
        var nowStock = oldStock[0] - usedStock;
        stockInfo.data[k][3] = nowStock;
        stockInfo.data[k][8] = stockInfo.data[k][7] * nowStock;
        usedUnitPrice.push([oldUnitPrice[0]]);
      }

      //2단계 - 한행에서 안끝날경우
      else if (usedStock > oldStock[0]) {
        var sum2 = oldStock[0];
        for (o = 1; o < oldStock.length; o++) {
          var sum2 = sum2 + oldStock[o];
          if (usedStock < sum2) {
            var o = o + 1;
            break;
          } //k가 1이면 두개의 행으로 충분하다는 뜻. k가 2면 세개의 행이 필요.
        }

        for (var l = 0; l < o - 1; l++) {
          unit.push(oldUnitPrice[l] * oldStock[l]); // unit의 배열에 세개의 행의 amount를 배열로 넣었다.
          var stockTep = stockTep + oldStock[l]; // stockTep 변수에 세개의 행의 재고를 더했다
          var nowUnitPrice = nowUnitPrice + unit[l]; // unitFin 변수에 세개의 행의 amount를 다 더했다.
        }

        var stock2 = stockTep + oldStock[o - 1] - usedStock; //
        var unit2 = oldUnitPrice[o - 1] * (oldStock[o - 1] - stock2); //
        var nowUnitPrice = (nowUnitPrice + unit2) / usedStock; // unitFin에는

        stockInfo.data[targetRow[0] + o - 1][3] = stock2;
        stockInfo.data[targetRow[0] + o - 1][8] =
          stockInfo.data[targetRow[0] + o - 1][7] * stock2;

        usedUnitPrice.push([nowUnitPrice]);
        stockInfo.data.splice(k - q + 1, o - 1);
      }
    }
    usedData.getRange(3, 3, usedUnitPrice.length).setValues(usedUnitPrice);
    stock.getRange("a4:l300").clearContent();
    stock.getRange(4, 1, stockInfo.data.length, 9).setValues(stockInfo.data);
    return stockInfo;
  };

  stockBackup("first"); //큰따옴표로 하면 오류난다
  inputImportData();
  stockBackup("second");
  inputRawPrice();
}
