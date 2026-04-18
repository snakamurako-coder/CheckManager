/**
 * 初回セットアップ：エディタ上部でこの関数を選択して「実行」してください。
 */
function setupAppEnvironment() {
  const rootFolderName = "課題点検アプリ_システムフォルダ";
  const userEmail = Session.getActiveUser().getEmail();
  
  // 1. フォルダ作成（既存の場合は取得）
  let folder = DriveApp.getFoldersByName(rootFolderName).hasNext() 
    ? DriveApp.getFoldersByName(rootFolderName).next() 
    : DriveApp.createFolder(rootFolderName);

  // 2. Configブック作成
  let configSs = SpreadsheetApp.create("【管理】Config_課題点検アプリ");
  DriveApp.getFileById(configSs.getId()).moveTo(folder);
  
  // Config: ユーザー管理シート
  let userSheet = configSs.getSheets()[0];
  userSheet.setName("ユーザー管理");
  userSheet.appendRow(["メールアドレス", "権限", "備考"]);
  userSheet.appendRow([userEmail, "admin", "作成者（自動登録）"]);
  
  // Config: アプリ設定シート
  let settingsSheet = configSs.insertSheet("アプリ設定");
  settingsSheet.appendRow(["項目", "値", "説明"]);
  settingsSheet.appendRow(["TARGET_SS_ID", "", "現在点検中のスプレッドシートID"]);
  settingsSheet.appendRow(["PASS_SCORE", "80", "小テスト合格点"]);

  // Config: ログシート
  let logSheet = configSs.insertSheet("利用ログ");
  logSheet.appendRow(["タイムスタンプ", "ユーザー", "操作", "クラス", "内容"]);

  // 3. サンプル点検票ブック作成（ご提示の構成を再現）
  let sampleSs = SpreadsheetApp.create("【サンプル】点検票ブック");
  DriveApp.getFileById(sampleSs.getId()).moveTo(folder);
  createSampleSheet(sampleSs.getSheets()[0]);

  // 4. スクリプトプロパティへの登録
  PropertiesService.getScriptProperties().setProperties({
    "CONFIG_SS_ID": configSs.getId(),
    "SAMPLE_SS_ID": sampleSs.getId()
  });
  
  // 初期設定としてサンプルをターゲットに設定
  updateConfigValue("TARGET_SS_ID", sampleSs.getId());

  Logger.log("セットアップ完了！マイドライブの「課題点検アプリ_システムフォルダ」を確認してください。");
}

/**
 * ユーザー指定の5行見出し構成を再現したサンプルシート作成
 */
function createSampleSheet(sheet) {
  sheet.setName("1組");
  const headers = [
    ["組","番号","ID","氏名","性別","提出率","提出数", 1, 2, 3, 4, 5], // 1行目: 通し番号
    ["","","","","","", "提出率→", "=(COUNTIF(H6:H50,\"提\"))/40", "", "", "", ""], // 2行目: 提出率
    ["","","","","","", "返却可否→", "済", "済", "済", "未", "未"], // 3行目: 返却
    ["","","","日付","","","", "4/10", "4/15", "4/20", "5/1", "5/10"], // 4行目: 日付
    ["組","番号","ID","氏名","性別","提出率","提出数", "課題A", "課題B", "小テスト1", "課題C", "小テスト2"] // 5行目: 課題名
  ];
  sheet.getRange(1, 1, 5, headers[0].length).setValues(headers);
  
  // 6行目以降にサンプル生徒データ
  let sampleStudents = [];
  for(let i=1; i<=40; i++) {
    sampleStudents.push([
      1, 
      i, 
      "ID"+(100+i), 
      "生徒氏名"+i, 
      i%2==0 ? "女" : "男", 
      `=IF(COUNTA($H$5:$Z$5)=0, 0, G${i+5}/COUNTA($H$5:$Z$5))`, 
      `=COUNTIF(H${i+5}:Z${i+5},"提")+COUNTIF(H${i+5}:Z${i+5}, 1)`
    ]);
  }
  sheet.getRange(6, 1, 40, 7).setValues(sampleStudents);
  sheet.getRange(6, 6, 40, 1).setNumberFormat("0.0%"); // 提出率を%表示
  
  // 条件付き書式（提:緑, 未:赤, 再:黄, 休:紫）
  let range = sheet.getRange("H6:Z50");
  let rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("提").setBackground("#b7e1cd").setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("未").setBackground("#f4c7c3").setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("再").setBackground("#fce8b2").setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("休").setBackground("#d9d2e9").setRanges([range]).build()
  ];
  sheet.setConditionalFormatRules(rules);
  
  // 見出し固定
  sheet.setFrozenRows(5);
  sheet.setFrozenColumns(7);
}

/**
 * Webアプリ表示の分岐と初期データ渡し
 */
function doGet() {
  const userEmail = Session.getActiveUser().getEmail();
  const role = getUserRole(userEmail);
  
  const template = HtmlService.createTemplateFromFile('index');
  template.isAdmin = (role === 'admin');
  template.userEmail = userEmail;
  
  return template.evaluate()
    .setTitle('課題点検・小テスト入力')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 権限チェック
function getUserRole(email) {
  const configId = PropertiesService.getScriptProperties().getProperty("CONFIG_SS_ID");
  if (!configId) return 'guest'; // セットアップ前
  const sheet = SpreadsheetApp.openById(configId).getSheetByName("ユーザー管理");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) return data[i][1];
  }
  return 'guest';
}

// 設定値の取得（UIから呼ばれる）
function getAdminSettings() {
  const configId = PropertiesService.getScriptProperties().getProperty("CONFIG_SS_ID");
  const sheet = SpreadsheetApp.openById(configId).getSheetByName("アプリ設定");
  const data = sheet.getDataRange().getValues();
  return {
    targetSsId: data[1][1],
    passScore: data[2][1]
  };
}

// 設定値の更新（UIから呼ばれる）
function updateConfigValue(key, value) {
  const configId = PropertiesService.getScriptProperties().getProperty("CONFIG_SS_ID");
  const sheet = SpreadsheetApp.openById(configId).getSheetByName("アプリ設定");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      break;
    }
  }
}

// 共通：現在設定されている点検用スプレッドシートを取得
function getAppSpreadsheet() {
  const configId = PropertiesService.getScriptProperties().getProperty("CONFIG_SS_ID");
  const sheet = SpreadsheetApp.openById(configId).getSheetByName("アプリ設定");
  const ssId = sheet.getRange("B2").getValue(); // TARGET_SS_ID
  // 設定がない場合はサンプルブックを読む
  return SpreadsheetApp.openById(ssId || PropertiesService.getScriptProperties().getProperty("SAMPLE_SS_ID"));
}

// クラス一覧（シート名）を取得
function getClassList() {
  const ss = getAppSpreadsheet();
  return ss.getSheets().map(s => s.getName()).filter(name => name !== '設定');
}

// 選択されたクラスの名簿と既存の課題名を取得
function getClassData(className) {
  const ss = getAppSpreadsheet();
  const sheet = ss.getSheetByName(className);
  const lastRow = sheet.getLastRow();
  
  // 名簿データは6行目〜、A列〜G列（7列分）を取得
  const dataRows = Math.max(1, lastRow - 5);
  const range = sheet.getRange(6, 1, dataRows, 7).getValues();
  // 課題名は5行目のH列(8列目)から取得
  const taskCols = Math.max(1, sheet.getLastColumn() - 7);
  const tasks = sheet.getRange(5, 8, 1, taskCols).getValues()[0];
  
  // 配列からオブジェクトに変換
  const students = range.map(row => ({
    group: row[0],
    no: row[1],
    id: row[2],
    name: row[3] // 氏名列
  })).filter(s => s.name); // 空白行を除外

  return { students: students, tasks: tasks };
}

// データの保存
function submitAttendanceData(className, taskIndex, dataArray) {
  const ss = getAppSpreadsheet();
  const sheet = ss.getSheetByName(className);
  const startRow = 6;
  const col = 8 + parseInt(taskIndex); // H列(8列目)からスタート
  
  // 縦方向の2次元配列に変換
  const values = dataArray.map(val => [val]);
  sheet.getRange(startRow, col, values.length, 1).setValues(values);
  
  // ログ記録
  logToConfig(className, "課題入力", `課題インデックス: ${taskIndex}`);
  
  return ss.getUrl();
}

function logToConfig(className, type, detail) {
  const configId = PropertiesService.getScriptProperties().getProperty("CONFIG_SS_ID");
  if (configId) {
    const logSheet = SpreadsheetApp.openById(configId).getSheetByName("利用ログ");
    logSheet.appendRow([new Date(), Session.getActiveUser().getEmail(), type, className, detail]);
  }
}

// インクルード用
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}