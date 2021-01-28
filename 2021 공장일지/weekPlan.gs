const checkMon = () => {
  var sheet = getSheet();
  var day = sheet.getRange("b1").getValue().getDay();
  if (day != 1) {
    SpreadsheetApp.getUi().alert("시작일이 월요일이 아닙니다!");
    return;
  }
  return "ok";
};
const pencil = () => {
  if (checkMon() === "ok") {
    var sheet = getSheet();
    sheet.getRange("l5:w22").clearContent();
    sheet.getRange("m24:w24").clearContent();
    sheet.getRange("m27:w27").clearContent();
    inputItemNameToCheckStock();
    inputPlanner();
    getInItems();
  }
};

const eraser = () => {
  var sheet = getSheet();
  sheet.getRange("l5:w22").clearContent();
  sheet.getRange("m24:w24").clearContent();
  sheet.getRange("m27:w27").clearContent();
  sheet.getRange("c3:h75").clearContent();
};

const getSheet = () => {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  return activeSheet;
};

const getStatusBoard = () => {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var statusBoard = ss.getSheetByName("현황판");
  return statusBoard;
};

const getInItems = () => {
  var sheet = getSheet();
  var statusBoard = getStatusBoard();
  var lastRow = statusBoard.getLastRow();
  var startDate = sheet.getRange("m2").getValue().getTime();
  var inItems = [];
  var colNumber = [3, 7, 11, 15, 19, 23];

  for (var i = 2; i < lastRow; i = i + 28) {
    if (
      statusBoard
        .getRange(i - 1, 3)
        .getValue()
        .getTime() === startDate
    ) {
      for (let r of colNumber) {
        var inItem = statusBoard.getRange(i, r, 4, 1).getValues();
        var col = (r - 1) / 2 + 11;
        sheet.getRange(17, col, 4, 1).setValues(inItem);
      }
    }
  }
};

const getInputItems = () => {
  var sheet = getSheet();
  var detail = [];
  var inputDatas = [];
  var inputDatasEmp = sheet.getRange("a3:h75").getValues();

  for (var i = 0; i < inputDatasEmp.length; i++) {
    var detail = inputDatasEmp[i];
    var detail = detail.slice(2, 7);
    var detail = detail.join();
    if (detail != ",,,,") {
      if (inputDatasEmp[i][1] === "") {
        SpreadsheetApp.getUi().alert(
          "제품명이 없는 곳에 숫자가 기입되어 있습니다!"
        );
        return;
      }
      inputDatas.push(inputDatasEmp[i]);
    }
  }
  return inputDatas;
};

const getItemName = () => {
  var totalItem = getInputItems();
  var itemName = [];

  totalItem.forEach((item) => {
    itemName.push(item[1]);
  });

  return itemName;
};

const inputItemNameToCheckStock = () => {
  var sheet = getSheet();
  var itemName = getItemName();
  var itemNumber = itemName.length;
  if (itemNumber < 1) {
    SpreadsheetApp.getUi().alert("제품별 예상 수량이 없습니다!");
  }
  var startRow = 24;

  if (itemNumber < 11) {
    sheet.getRange(startRow, 13, 1, itemNumber).setValues([itemName]);
  } else {
    sheet.getRange(startRow, 13, 1, 11).setValues([itemName.slice(0, 11)]);
    sheet
      .getRange(startRow + 3, 13, 1, itemNumber - 11)
      .setValues([itemName.splice(11, itemNumber)]);
  }
};

const inputPlanner = () => {
  var sheet = getSheet();
  var inputDatas = getInputItems();
  var inputData = {};

  //월요일인 것만 가져오기
  for (var i = 0; i < inputDatas.length; i++) {
    for (var dayCheck = 2; dayCheck < 8; dayCheck++) {
      if (inputDatas[i][dayCheck] != "") {
        //이게 월요일이야
        inputData.separation = inputDatas[i][0];
        inputData.itemName = inputDatas[i][1];
        inputData.qty = inputDatas[i][dayCheck];

        if (inputData.separation === "기계실") {
          for (var r = 5; r < 9; r++) {
            var savedData = sheet.getRange(r, (dayCheck + 4) * 2).getValue();
            if (savedData === "") {
              sheet
                .getRange(r, (dayCheck + 4) * 2)
                .setValue(inputData.itemName);
              sheet.getRange(r, (dayCheck + 4) * 2 + 1).setValue(inputData.qty);
              break;
            }
          }
        } else if (inputData.separation === "코코넛라인") {
          for (var r = 9; r < 13; r++) {
            var savedData = sheet.getRange(r, (dayCheck + 4) * 2).getValue();
            if (savedData === "") {
              sheet
                .getRange(r, (dayCheck + 4) * 2)
                .setValue(inputData.itemName);
              sheet.getRange(r, (dayCheck + 4) * 2 + 1).setValue(inputData.qty);
              break;
            }
          }
        } else if (inputData.separation === "기타") {
          for (var r = 13; r < 17; r++) {
            var savedData = sheet.getRange(r, (dayCheck + 4) * 2).getValue();
            if (savedData === "") {
              sheet
                .getRange(r, (dayCheck + 4) * 2)
                .setValue(inputData.itemName);
              sheet.getRange(r, (dayCheck + 4) * 2 + 1).setValue(inputData.qty);
              break;
            }
          }
        }
      }
    }
  }
};

function print() {
  SpreadsheetApp.flush();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getSheet();
  var range = sheet.getRange("k2:w28");

  var gid = sheet.getSheetId();
  var printRange = objectToQueryString({
    c1: range.getColumn() - 1,
    r1: range.getRow() - 1,
    c2: range.getColumn() + range.getWidth() - 1,
    r2: range.getRow() + range.getHeight() - 3,
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
  var msg = "주간계획표 인쇄 중";
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

var PRINT_OPTIONS = {
  size: 8, // paper size. 0=letter, 1=tabloid, 2=Legal, 3=statement, 4=executive, 5=folio, 6=A3, 7=A4, 8=A5, 9=B4, 10=B
  fzr: false, // repeat row headers
  portrait: false, // false=landscape
  fitw: true, // fit window or actual size
  gridlines: false, // show gridlines
  printtitle: false,
  sheetnames: false,
  pagenum: "UNDEFINED", // CENTER = show page numbers / UNDEFINED = do not show
  attachment: false,
};

var PDF_OPTS = objectToQueryString(PRINT_OPTIONS);
