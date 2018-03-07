/**
 * ・[メンバーマスタ]シート
 * A1 : [あだ名]
 * B1 : [メールアドレス]
 * C1 : [利用可能ポイント]
 * D1 : [累計ポイント]
 * E1 : [受取済チケット]
 * F1 : [あだ名チェック用、入力ミスとかのをいれるとこ]
 * ※1行目がヘッダ、2行目からデータとしてやってます
 * 
 * ・[投票ログ]シート
 * A1 : [投票元(メールアドレス)]
 * B1 : [投票先(あだ名)]
 * C1 : [投票ポイント]
 * D1 : [投票時刻]
 * E1 : [あだ名チェックエラー(メンバのF列で変換できなかった場合にチェックが入る)]
 */

/**
 * <画面から>
 * 初期画面を取得する
 */
function doGet(request) {
  // 初期表示用の画面を返却する
  return HtmlService.createTemplateFromFile('index.html').evaluate();
}

/**
 * メインのスプレッドシートを取得する
 */
function getSheet(sheetName) {
  var id = "1bYnk1YlGxxpoGcxS-thUs6u4-RoRk3-tETcQSyB1EUo";// ※実際につなぐスプレッドシートのIDを設定
  var spreadSheet = SpreadsheetApp.openById(id);
  return spreadSheet.getSheetByName(sheetName);;
}

/**
 * 『メンバーマスタ』のシート情報を取得する
 */
function getMasterSheet() {
  return getSheet("メンバーマスタ");
}

/**
 * 『投票ログ』のシート情報を取得する
 */
function getLogSheet() {
  return getSheet("投票ログ");
}

/**
 * <画面から>
 * 有効なメールアドレスかをチェック
 */
function validateMailAddresses(form)
{
  var masterSheet = getMasterSheet()
  var mailAddresses = masterSheet.getRange(2, 2, masterSheet.getLastRow() - 1, 1).getValues();
  if (contains(mailAddresses, form.mailAddress)) {
    return HtmlService.createTemplateFromFile('index.html').evaluate();
  } else {
    return HtmlService.createTemplateFromFile('error.html').evaluate();
  }
}

/**
 * 配列内にデータが存在するかをチェック
 */
function contains(array, mailAddress) {
  for(var i = 0; i < array.length; i++) {
    if (array[i][0] == mailAddress) {
      return true;
    }
  }
}

/**
 * マスタシートからデータを取得
 */
function getMyData(mailAddress) {
  var masterSheet = getMasterSheet();
  var range = masterSheet.getRange(2, 2, masterSheet.getLastRow() - 1, 3).getValues();
  for (var i = 0; i < range.length; i++) {
    if (range[i][0] == mailAddress) {
      var row = i + 2;
      return {mailAddress: range[i][0], point: range[i][1], totalPoint: range[i][2], row: row};
    }
  }
  return {mailAddress: "", point: 0, totalPoint: 0, row: 0};
  
}

/**
 * <画面から>
 * 投票を行う
 */
function doPost(form) {
  // 投票を受取り、所持ポイントを超過していないかをチェック
  // 投票元と投票先が同じでないことを確認
  // 問題なければ所持ポイントを減算
  // 投票先があだ名チェックに含まれる場合、本来の名前に置き換え
  // 登録を行う
  var sheet = getLogSheet();
  var myData = getMyData(form.mailAddress);
  if (!isAvailablePayment(myData, form)) {
    return "ERROR";
  }
  var masterSheet = getMasterSheet();
  masterSheet.getRange(myData.row, 3, 1, 1).setValues([[myData.point - form.point]]);
  if (isValidName(form.to)) {
    sheet.appendRow([form.mailAddress, form.to, form.point, new Date()]);
    return "SUCCESS";
  }
  var name = transName(form.to);
  if (name == form.to) {
    sheet.appendRow([form.mailAddress, form.to, form.point, new Date(), "x"]);
  } else {
    sheet.appendRow([form.mailAddress, name, form.point, new Date()]);
  }
  return "SUCCESS";
}

/**
 * 投票かをチェック
 * 1. 所持ポイント以内であること
 * 2. 自分への投票でないこと
 */
function isAvailablePayment(myData, form) {
  if (myData.point < form.point) {
    return false;
  }
  var masterSheet = getMasterSheet();
  var masters = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 2).getValues();
  var name = form.to;
  var mailAddress = form.mailAddress;
  var masterName;
  for (var i = 0; i < masters.length; i++) {
    if (masters[i][0] == name && masters[i][1] == mailAddress) {
      return false;
    } else if (masters[i][1] == mailAddress) {
      masterName = masters[i][0];
    }
  }
  var tmpName = transName(name);
  if (tmpName == masterName) {
    return false;
  } 
  return true;
}

/**
 * 有効なあだ名かをチェック
 */
function isValidName(name) {
  var masterSheet = getMasterSheet();
  var names = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < names.length; i++) {
    if (names[i][0] == name) {
      return true;
    }
  }
  return false;
}

/**
 * 正式名称に変換できるものについて、正式名称を取得する
 */
function transName(name) {
  var masterSheet = getMasterSheet();
  var names = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 6).getValues();
  for (var i = 0; i < names.length; i++) {
    var tempName = names[i];
    if (!names[i][5]) {
      continue;
    }
    var sameNames = names[i][5].split(",");
    for (var j = 0; j < sameNames.length; j++) {
      if (sameNames[j] == name) {
        return names[i][0];
      }
    }
  }
  return name;
}

/**
 * <画面から>
 * 現在の受領合計ポイントを取得する
 */
function getTotalPoint(mailAddress) {
  var sheet = getMasterSheet();
  var range = sheet.getRange(2, 2, sheet.getLastRow() - 1, 3).getValues();
  for (var i = 0; i < range.length; i++) {
    if (range[i][0] == mailAddress) {
      return range[i][2];
    }
  }
  return 0;
}

/**
 * <トリガーから>
 * 月次集計処理
 */
function monthlyCalc() {
  // 月次処理を行う
  var masterSheet = getMasterSheet();
  var logSheet = getLogSheet();
  // 0. 全ての使用可能ポイントをゼロする
  var masterRange = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < masterRange.length; i++) {
    masterSheet.getRange(i + 2, 3, 1, 1).setValues([[0]]);
  }
  // 1. ログから各人の累計ポイントを計算する
  var logRange = logSheet.getRange(2, 2, logSheet.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < masterRange. length; i++) {
    var totalPoint = 0;
    var tmpName = masterRange[i][0];
    for (var j = 0; j < logRange.length; j++) {
      if (logRange[j][0] == tmpName) {
        totalPoint += logRange[j][1];
      }
    }
    masterSheet.getRange(i + 2, 4, 1, 1).setValues([[totalPoint]]);
  }
  // 2. 全ての使用可能ポイントをリセットする
  for (var i = 0; i < masterRange.length; i++) {
    masterSheet.getRange(i + 2, 3, 1, 1).setValues([[1000]]);
  }
}

function test() {
  var form = {mailAddress : "takashima.u@gmail.com", to : "なると", point : 100};
  //doPost(form);
  Logger.log(doPost(form));
}
