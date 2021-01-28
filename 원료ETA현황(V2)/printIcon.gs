function remittance() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var remitSheet = ss.getActiveSheet();
  var activeRow = remitSheet.getActiveCell().getRow();
  if (activeRow < 4 || activeRow > 15) {
    SpreadsheetApp.getUi().alert("송금건에 커서를 놓고 진행하세요!!");
    return;
  }
  var accSheet = ss.getSheetByName("계좌관리");
  var remitReg = ss.getSheetByName("송금신청서");
  var remitHistory = ss.getSheetByName("송금이력");
  var remitLastRow = remitHistory.getLastRow();

  var remitTitle = remitSheet.getRange(3, 1, 1, 14).getValues();
  var remitTitle = remitTitle[0];
  var accTitle = accSheet.getRange(1, 1, 1, 11).getValues();
  var accTitle = accTitle[0];

  const getRemitDatas = () => {
    //송금정보 가져오기
    var remitDatasAll = remitSheet.getRange(4, 1, 12, 14).getValues();
    var remitDatasLength = remitDatasAll.length;
    for (var i = 0; i < remitDatasLength; i++) {
      if (remitDatasAll[i][1] === "") {
        var remitDatas = remitDatasAll.slice(0, i);
        break;
      }
    }
    return remitDatas;
  };

  const getAccDatas = () => {
    //계좌정보 가져오기
    var accSheetGetlastrow = accSheet.getLastRow();
    var accDatas = accSheet.getRange(2, 1, accSheetGetlastrow, 11).getValues();

    return accDatas;
  };

  const getTheData = () => {
    //해당 건에 대한 정보 가져오기
    var remitDatas = getRemitDatas();
    var remitOrder = activeRow - 4;
    var remitData = remitDatas[remitOrder];
    var sum = 0;
    validRow = [];

    for (var i = 0; i < remitDatas.length; i++) {
      if (
        remitDatas[i][11] === remitData[11] &&
        remitDatas[i][1] === remitData[1] &&
        remitDatas[i][7] === remitData[7] &&
        remitDatas[i][9].toString() === remitData[9].toString()
      ) {
        var sum = sum + remitDatas[i][8];
        validRow.push(i);
      }
    }
    remitData[8] = sum;

    var accDatas = getAccDatas();
    for (var i = 1; i <= accDatas.length; i++) {
      if (accDatas[i][1] === remitData[2]) {
        var accData = accDatas[i];
        break;
      }
    }

    var values = remitData.concat(accData);
    Logger.log(values);
    return values;
  };

  const alertError = () => {
    const values = getTheData();
    if (values[0] === "(주)허니텍") {
      SpreadsheetApp.getUi().alert(
        "(주)허니텍은 송금 신청서를 작성하지 않아요!"
      );
      return;
    }

    for (var i = 1; i < 12; i++) {
      if (values[i] === "") {
        var errorMsg = accTitle[0][i];
        SpreadsheetApp.getUi().alert(errorMsg + " 입력하세요!");
        return;
      }
    }

    for (var k = 14; k < 23; k++) {
      if (values[k] === "") {
        SpreadsheetApp.getUi().alert(values[1] + "의 계좌정보를 완성하세요!");
        return;
      }
    }

    if (values[12] != "") {
      remitReg.getRange(33, 1).setValue("v");
    }

    if (values[13] != "") {
      remitReg.getRange(34, 1).setValue("v");
    }
    checkLcOrTt();
  };

  const checkLcOrTt = () => {
    const values = getTheData();
    if (values[7] === "LC") {
      processLc();
    } else {
      makeRemit();
    }
  };

  const processLc = () => {
    SpreadsheetApp.flush();
    var url = "https://www.kbstar.com/";
    var htmlTemplate = HtmlService.createTemplateFromFile("index");
    htmlTemplate.url = url;
    var msg = "KB 홈페이지로 이동 중";
    SpreadsheetApp.getUi().showModalDialog(
      htmlTemplate.evaluate().setHeight(10).setWidth(320),
      msg
    );
    //퍼페티어
    Utilities.sleep(10 * 1000);
    erasev();
    return;
  };

  const ourAccInfo = () => {
    const accInfor = [];

    const hanpelAcc = {};
    const daehanAcc = {};

    hanpelAcc.koreanName = "한펠";
    hanpelAcc.englishName = "HANPEL TECH CO.,LTD";
    hanpelAcc.regiNum = "312-81-14560";
    hanpelAcc.tel = "031-387-7442";
    hanpelAcc.add = "경기 안양 동안 관악대로 298, 도명빌딩 501호";

    daehanAcc.koreanName = "대한산업";
    daehanAcc.englishName = "DAEHAN CORPORATION";
    daehanAcc.regiNum = "535-37-00574";
    daehanAcc.tel = "010-2378-1502";
    daehanAcc.add = "충남 천안시 동남구 수신면 수신로 705-5";

    accInfor.push(hanpelAcc);
    accInfor.push(daehanAcc);

    Logger.log(accInfor);
    return accInfor;
  };

  const makeRemit = () => {
    const finalData = getTheData();
    const now = new Date();
    finalData.push(now);
    remitHistory
      .getRange(remitLastRow + 1, 1, 1, finalData.length)
      .setValues([finalData]);

    const accInformation = ourAccInfo();
    if (finalData[0] === "한펠") {
      remitReg.getRange(9, 6).setValue(accInformation[0].koreanName);
      remitReg.getRange(10, 6).setValue(accInformation[0].regiNum);
      remitReg.getRange(11, 6).setValue(accInformation[0].add);
      remitReg.getRange(9, 10).setValue(accInformation[0].englishName);
      remitReg.getRange(10, 12).setValue(accInformation[0].tel);
    } else if (finalData[0] === "대한산업") {
      remitReg.getRange(9, 6).setValue(accInformation[1].koreanName);
      remitReg.getRange(10, 6).setValue(accInformation[1].regiNum);
      remitReg.getRange(11, 6).setValue(accInformation[1].add);
      remitReg.getRange(9, 10).setValue(accInformation[1].englishName);
      remitReg.getRange(10, 12).setValue(accInformation[1].tel);
    }

    remitReg.getRange(15, 9).setValue(finalData[8]); // 금액
    remitReg.getRange(27, 6).setValue(finalData[9]); // 날짜
    remitReg.getRange(30, 7).setValue(finalData[9]); // 날짜
    remitReg.getRange(27, 11).setValue(finalData[11]); // invoice 넘버
    remitReg.getRange(18, 6).setValue(finalData[16]); // 수취인 성명
    remitReg.getRange(19, 6).setValue(finalData[17]); // 수취인 주소
    remitReg.getRange(19, 12).setValue(finalData[18]); // 수취인 국가
    remitReg.getRange(20, 9).setValue(finalData[19]); // 은행
    remitReg.getRange(21, 9).setValue(finalData[20]); // swift
    remitReg.getRange(24, 12).setValue(finalData[21]); // 은행 국가
    remitReg.getRange(25, 12).setValue("'" + finalData[22]); // 계좌번호
    remitReg.getRange(26, 6).setValue(finalData[23]); // 중계은행

    //여기에다 lc와 tt에 따라서 다르게 진행
    printSelectedRange();
    Utilities.sleep(10 * 1000);
    erasev();
    return;
  };

  var PRINT_OPTIONS = {
    size: 7, // paper size. 0=letter, 1=tabloid, 2=Legal, 3=statement, 4=executive, 5=folio, 6=A3, 7=A4, 8=A5, 9=B4, 10=B
    fzr: false, // repeat row headers
    portrait: true, // false=landscape
    fitw: true, // fit window or actual size
    gridlines: false, // show gridlines
    printtitle: false,
    sheetnames: false,
    pagenum: "UNDEFINED", // CENTER = show page numbers / UNDEFINED = do not show
    attachment: false,
  };

  var PDF_OPTS = objectToQueryString(PRINT_OPTIONS);

  function printSelectedRange() {
    SpreadsheetApp.flush();
    var range = remitReg.getRange("a1:h36");

    var gid = remitReg.getSheetId();
    var printRange = objectToQueryString({
      c1: range.getColumn() - 1,
      r1: range.getRow() - 1,
      c2: range.getColumn() + range.getWidth() + 10,
      r2: range.getRow() + range.getHeight() - 1,
    });
    var url =
      ss.getUrl().replace(/edit$/, "") +
      "export?format=pdf" +
      PDF_OPTS +
      printRange +
      "&gid=" +
      gid;

    var htmlTemplate = HtmlService.createTemplateFromFile("js");
    htmlTemplate.url = url;
    var msg = validRow.length + "건 통합 송금신청서 작성 중";
    SpreadsheetApp.getUi().showModalDialog(
      htmlTemplate.evaluate().setHeight(10).setWidth(320),
      msg
    );
  }

  function objectToQueryString(obj) {
    return Object.keys(obj)
      .map(function (key) {
        return Utilities.formatString("&%s=%s", key, obj[key]);
      })
      .join("");
  }

  function erasev() {
    remitReg.getRange("a33:a34").clearContent();
    for (var o = 0; o < validRow.length; o++) {
      remitSheet.getRange(validRow[o] + 4, 13, 1, 2).clearContent();
    }
  }
  alertError();
}

//     const remitAccData = {};
//     var key = remitTitle.concat(accTitle);
//     var value = remitData.concat(accData);

//     for (var i = 0; i < key.length; i++) {
//         remitAccData[key[i]] = value[i];
//     }

//     Logger.log(remitAccData);
//     return remitAccData;
// }
