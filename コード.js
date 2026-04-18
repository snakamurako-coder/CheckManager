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
  settingsSheet.appendRow(["ASSIGNMENT_SS_ID", "", "提出物用スプレッドシートID"]);
  settingsSheet.appendRow(["QUIZ_SS_ID", "", "小テスト用スプレッドシートID"]);
  settingsSheet.appendRow(["PASS_SCORE", "80", "小テスト合格点"]);
  settingsSheet.appendRow(["ID_USE_FLAG", "false", "IDから組・番号を抽出するか(true/false)"]);
  settingsSheet.appendRow(["ID_CLASS_START", "1", "組抽出の開始文字位置(1始まり)"]);
  settingsSheet.appendRow(["ID_CLASS_LEN", "1", "組抽出の文字数"]);
  settingsSheet.appendRow(["ID_NUMBER_START", "2", "番号抽出の開始文字位置(1始まり)"]);
  settingsSheet.appendRow(["ID_NUMBER_LEN", "2", "番号抽出の文字数"]);
  settingsSheet.appendRow(["HEADER_ROWS", "5", "見出し行数"]);
  settingsSheet.appendRow(["ROSTER_COLS", "7", "名簿部分の列数"]);

  // Config: 利用ログシート
  let logSheet = configSs.insertSheet("利用ログ");
  logSheet.appendRow(["タイムスタンプ", "ユーザー", "操作", "クラス", "内容"]);

  // 3. サンプル点検票ブック作成
  let sampleSs = SpreadsheetApp.create("【サンプル】点検票ブック");
  DriveApp.getFileById(sampleSs.getId()).moveTo(folder);
  createSampleSheet(sampleSs.getSheets()[0]);

  // 4. スクリプトプロパティへの登録
  PropertiesService.getScriptProperties().setProperties({
    "CONFIG_SS_ID": configSs.getId(),
    "SAMPLE_SS_ID": sampleSs.getId()
  });
  
  // 初期設定としてサンプルをターゲットに設定
  updateConfigValue("ASSIGNMENT_SS_ID", sampleSs.getId());
  updateConfigValue("QUIZ_SS_ID", sampleSs.getId());

  Logger.log("セットアップ完了！マイドライブの「課題点検アプリ_システムフォルダ」を確認してください。");
}

function createSampleSheet(sheet) {
  sheet.setName("1組");
  const headers = [
    ["組","番号","ID","氏名","性別","提出率","提出数", 1, 2, 3, 4, 5],
    ["","","","","","", "提出率→", "=(COUNTIF(H6:H50,\"提\"))/40", "", "", "", ""],
    ["","","","","","", "返却可否→", "済", "済", "済", "未", "未"],
    ["","","","日付","","","", "4/10", "4/15", "4/20", "5/1", "5/10"],
    ["組","番号","ID","氏名","性別","提出率","提出数", "課題A", "課題B", "小テスト1", "課題C", "小テスト2"]
  ];
  sheet.getRange(1, 1, 5, headers[0].length).setValues(headers);
  
  let sampleStudents = [];
  for(let i=1; i<=40; i++) {
    // IDは 101, 102 ... 140 となるようにする
    let idStr = "1" + ("0" + i).slice(-2);
    sampleStudents.push([
      1, 
      i, 
      idStr, 
      "生徒氏名"+i, 
      i%2==0 ? "女" : "男", 
      `=IF(COUNTA($H$5:$Z$5)=0, 0, G${i+5}/COUNTA($H$5:$Z$5))`, 
      `=COUNTIF(H${i+5}:Z${i+5},"提")+COUNTIF(H${i+5}:Z${i+5}, 1)`
    ]);
  }
  sheet.getRange(6, 1, 40, 7).setValues(sampleStudents);
  sheet.getRange(6, 6, 40, 1).setNumberFormat("0.0%");
  
  // 条件付き書式
  let range = sheet.getRange("H6:Z50");
  let rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("提").setBackground("#b7e1cd").setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("未").setBackground("#f4c7c3").setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("再").setBackground("#fce8b2").setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("休").setBackground("#d9d2e9").setRanges([range]).build()
  ];
  sheet.setConditionalFormatRules(rules);
  
  sheet.setFrozenRows(5);
  sheet.setFrozenColumns(7);
}

function doGet() {
  const userEmail = Session.getActiveUser().getEmail();
  const role = getUserRole(userEmail);
  
  const template = HtmlService.createTemplateFromFile('index');
  template.isAdmin = (role === 'admin');
  template.isGuest = (role === 'guest');
  template.userEmail = userEmail;
  
  return template.evaluate()
    .setTitle('課題点検・小テスト入力')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function getUserRole(email) {
  const configId = PropertiesService.getScriptProperties().getProperty("CONFIG_SS_ID");
  if (!configId) return 'guest'; // セットアップ前
  const sheet = SpreadsheetApp.openById(configId).getSheetByName("ユーザー管理");
  if (!sheet) return 'guest';
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) return data[i][1];
  }
  return 'guest';
}

function getAdminSettings() {
  const configId = PropertiesService.getScriptProperties().getProperty("CONFIG_SS_ID");
  const sheet = SpreadsheetApp.openById(configId).getSheetByName("アプリ設定");
  const data = sheet.getDataRange().getValues();
  let settings = {};
  for(let i=1; i<data.length; i++){
    settings[data[i][0]] = data[i][1];
  }
  return settings;
}

function updateConfigValue(key, value) {
  const configId = PropertiesService.getScriptProperties().getProperty("CONFIG_SS_ID");
  const sheet = SpreadsheetApp.openById(configId).getSheetByName("アプリ設定");
  const data = sheet.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      found = true;
      break;
    }
  }
  if (!found) {
    sheet.appendRow([key, value, ""]);
  }
}

function saveUserState(stateObj) {
  PropertiesService.getUserProperties().setProperty("APP_STATE", JSON.stringify(stateObj));
}

function loadUserState() {
  const stateJson = PropertiesService.getUserProperties().getProperty("APP_STATE");
  return stateJson ? JSON.parse(stateJson) : null;
}

function clearUserState() {
  PropertiesService.getUserProperties().deleteProperty("APP_STATE");
}

function getAppSpreadsheet(bookType) {
  const settings = getAdminSettings();
  let targetId = bookType === "quiz" ? settings["QUIZ_SS_ID"] : settings["ASSIGNMENT_SS_ID"];
  return SpreadsheetApp.openById(targetId || PropertiesService.getScriptProperties().getProperty("SAMPLE_SS_ID"));
}

function getClassList(bookType) {
  const ss = getAppSpreadsheet(bookType);
  return ss.getSheets().map(s => s.getName()).filter(name => !['設定', 'ユーザー管理', 'アプリ設定', '利用ログ'].includes(name));
}

function getClassData(bookType, className) {
  const ss = getAppSpreadsheet(bookType);
  const sheet = ss.getSheetByName(className);
  if (!sheet) return null;
  const lastRow = sheet.getLastRow();
  
  const settings = getAdminSettings();
  const headerRows = parseInt(settings.HEADER_ROWS) || 5;
  const rosterCols = parseInt(settings.ROSTER_COLS) || 7;
  const idUseFlag = String(settings.ID_USE_FLAG).toLowerCase() === 'true';
  const idClassStart = parseInt(settings.ID_CLASS_START) - 1 || 0;
  const idClassLen = parseInt(settings.ID_CLASS_LEN) || 1;
  const idNumStart = parseInt(settings.ID_NUMBER_START) - 1 || 0;
  const idNumLen = parseInt(settings.ID_NUMBER_LEN) || 2;
  
  const dataRows = Math.max(1, lastRow - headerRows);
  const range = sheet.getRange(headerRows + 1, 1, dataRows, rosterCols).getValues();
  
  const taskCols = Math.max(1, sheet.getLastColumn() - rosterCols);
  const tasks = sheet.getRange(headerRows, rosterCols + 1, 1, taskCols).getValues()[0];
  
  const students = range.map(row => {
    let group = row[0]; // 1列目 (A列)
    let no = row[1];    // 2列目 (B列)
    let id = row[2];    // 3列目 (C列)
    let name = row[3];  // 4列目 (D列)
    
    if (idUseFlag && id) {
      const idStr = String(id);
      group = idStr.substring(idClassStart, idClassStart + idClassLen);
      no = idStr.substring(idNumStart, idNumStart + idNumLen);
    }
    
    return {
      group: group,
      no: no,
      id: id,
      name: name
    };
  }).filter(s => s.name);

  return { students: students, tasks: tasks };
}

function submitAttendanceData(bookType, className, taskIndex, taskName, dataArray) {
  const ss = getAppSpreadsheet(bookType);
  const sheet = ss.getSheetByName(className);
  
  const settings = getAdminSettings();
  const headerRows = parseInt(settings.HEADER_ROWS) || 5;
  const rosterCols = parseInt(settings.ROSTER_COLS) || 7;
  
  const startRow = headerRows + 1;
  const col = rosterCols + 1 + parseInt(taskIndex);
  
  // 新規課題名の書き込み
  if (taskName) {
    const taskNameCell = sheet.getRange(headerRows, col);
    if (!taskNameCell.getValue()) {
      taskNameCell.setValue(taskName);
    }
  }

  const values = dataArray.map(val => [val === null || val === undefined ? "" : val]);
  sheet.getRange(startRow, col, values.length, 1).setValues(values);
  
  logToConfig(className, bookType === 'quiz' ? "小テスト入力" : "課題入力", `課題列: ${parseInt(taskIndex)+1}, 課題名: ${taskName || '既存'}`);
  
  clearUserState();
  
  return ss.getUrl();
}

function logToConfig(className, type, detail) {
  const configId = PropertiesService.getScriptProperties().getProperty("CONFIG_SS_ID");
  if (configId) {
    const logSheet = SpreadsheetApp.openById(configId).getSheetByName("利用ログ");
    if (logSheet) {
      logSheet.appendRow([new Date(), Session.getActiveUser().getEmail(), type, className, detail]);
    }
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}