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
  settingsSheet.appendRow(["HEADER_ROWS", "5", "見出し行数"]);
  settingsSheet.appendRow(["ROSTER_COLS", "7", "名簿部分の列数"]);
  settingsSheet.appendRow(["COL_MAP_NAME", "3", "氏名列のインデックス(0始まり)"]);
  settingsSheet.appendRow(["COL_MAP_ID", "2", "ID列のインデックス"]);
  settingsSheet.appendRow(["COL_MAP_CLASS", "0", "組列のインデックス"]);
  settingsSheet.appendRow(["COL_MAP_NUMBER", "1", "番号列のインデックス"]);

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
function saveAdminSettingsFromUI(assignId, quizId, mapName, mapId, mapClass, mapNum, rosterCols, headerRows) {
  updateConfigValue("ASSIGNMENT_SS_ID", assignId);
  updateConfigValue("QUIZ_SS_ID", quizId);
  if (mapName !== "") updateConfigValue("COL_MAP_NAME", mapName);
  if (mapId !== "") updateConfigValue("COL_MAP_ID", mapId);
  if (mapClass !== "") updateConfigValue("COL_MAP_CLASS", mapClass);
  if (mapNum !== "") updateConfigValue("COL_MAP_NUMBER", mapNum);
  if (rosterCols !== "") updateConfigValue("ROSTER_COLS", rosterCols);
  if (headerRows !== "") updateConfigValue("HEADER_ROWS", headerRows);
  
  if (assignId) registerBookTypeMarker("assignment", assignId);
  if (quizId) registerBookTypeMarker("quiz", quizId);
  
  return true;
}

function getHeadersFromSheet(headerRow) {
  const settings = getAdminSettings();
  const assignId = settings.ASSIGNMENT_SS_ID || PropertiesService.getScriptProperties().getProperty("SAMPLE_SS_ID");
  if (!assignId) return [];
  
  try {
    const ss = SpreadsheetApp.openById(assignId);
    const sheet = ss.getSheets().find(s => !['設定', 'ユーザー管理', 'アプリ設定', '利用ログ'].includes(s.getName()));
    if (!sheet) return [];
    
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return [];
    
    const rowNum = parseInt(headerRow) || 5;
    const values = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
    
    let headers = [];
    for (let i = 0; i < values.length; i++) {
      if (values[i] === "") break;
      headers.push(String(values[i]));
    }
    return headers;
  } catch(e) {
    Logger.log("getHeadersFromSheet Error: " + e);
    return [];
  }
}

function registerBookTypeMarker(bookType, ssId) {
  try {
    const ss = SpreadsheetApp.openById(ssId);
    const label = bookType === 'quiz' ? '小テスト' : '提出物';
    const sheets = ss.getSheets();
    for (let s of sheets) {
      if (!['設定', 'ユーザー管理', 'アプリ設定', '利用ログ'].includes(s.getName())) {
        s.getRange("A1").setValue(label);
      }
    }
  } catch(e) {
    Logger.log("Failed to register A1 marker: " + e);
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
  
  const mapName = parseInt(settings.COL_MAP_NAME);
  const mapId = parseInt(settings.COL_MAP_ID);
  const mapClass = parseInt(settings.COL_MAP_CLASS);
  const mapNum = parseInt(settings.COL_MAP_NUMBER);

  // default fallbacks if not set
  const colName = isNaN(mapName) ? 3 : mapName;
  const colId = isNaN(mapId) ? 2 : mapId;
  const colClass = isNaN(mapClass) ? 0 : mapClass;
  const colNum = isNaN(mapNum) ? 1 : mapNum;
  
  const dataRows = Math.max(1, lastRow - headerRows);
  const range = sheet.getRange(headerRows + 1, 1, dataRows, rosterCols).getValues();
  
  const taskCols = Math.max(1, sheet.getLastColumn() - rosterCols);
  
  const headerDateRow = headerRows > 1 ? headerRows - 1 : 1;
  const dates = sheet.getRange(headerDateRow, rosterCols + 1, 1, taskCols).getValues()[0];
  const taskNames = sheet.getRange(headerRows, rosterCols + 1, 1, taskCols).getValues()[0];
  
  let taskList = [];
  const now = new Date();
  
  for (let i = 0; i < taskCols; i++) {
    let name = taskNames[i];
    if (!name) continue; // 空の列はスキップ
    
    let dateVal = dates[i];
    let dateStr = "";
    let diff = Infinity;
    
    if (dateVal instanceof Date) {
      dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "M/d");
      diff = Math.abs(now.getTime() - dateVal.getTime());
    } else if (dateVal) {
      dateStr = String(dateVal);
      let currentYear = now.getFullYear();
      let parts = dateStr.split('/');
      if (parts.length === 2) {
        let parsed = new Date(currentYear, parseInt(parts[0])-1, parseInt(parts[1]));
        diff = Math.abs(now.getTime() - parsed.getTime());
      }
    }
    
    taskList.push({
      colIndex: i,
      name: name,
      date: dateStr,
      diff: diff
    });
  }
  
  // 差が小さい順（現在に近い順）にソート
  taskList.sort((a, b) => a.diff - b.diff);
  
  const students = range.map(row => {
    return {
      group: row[colClass],
      no: row[colNum],
      id: row[colId],
      name: row[colName]
    };
  }).filter(s => s.name);

  return { students: students, tasks: taskList };
}

function createNewTask(bookType, className, dateStr, name) {
  const ss = getAppSpreadsheet(bookType);
  const sheet = ss.getSheetByName(className);
  
  const settings = getAdminSettings();
  const headerRows = parseInt(settings.HEADER_ROWS) || 5;
  const rosterCols = parseInt(settings.ROSTER_COLS) || 7;
  const headerDateRow = headerRows > 1 ? headerRows - 1 : 1;
  
  let targetCol = rosterCols + 1;
  const lastCol = sheet.getLastColumn();
  
  if (lastCol > rosterCols) {
    const tasks = sheet.getRange(headerRows, rosterCols + 1, 1, lastCol - rosterCols).getValues()[0];
    let foundEmpty = false;
    for (let i = 0; i < tasks.length; i++) {
      if (!tasks[i]) {
        targetCol = rosterCols + 1 + i;
        foundEmpty = true;
        break;
      }
    }
    if (!foundEmpty) {
      targetCol = lastCol + 1;
    }
  }
  
  if (targetCol > sheet.getMaxColumns()) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), targetCol - sheet.getMaxColumns());
  }
  
  sheet.getRange(headerDateRow, targetCol).setValue(dateStr);
  sheet.getRange(headerRows, targetCol).setValue(name);
  
  logToConfig(className, "新規作成", `ブック: ${bookType}, 課題名: ${name}, 日付: ${dateStr}`);
  
  return true;
}

function submitAttendanceData(bookType, className, taskIndex, taskName, dataArray) {
  const ss = getAppSpreadsheet(bookType);
  const sheet = ss.getSheetByName(className);
  
  const settings = getAdminSettings();
  const headerRows = parseInt(settings.HEADER_ROWS) || 5;
  const rosterCols = parseInt(settings.ROSTER_COLS) || 7;
  
  const startRow = headerRows + 1;
  const col = rosterCols + 1 + parseInt(taskIndex);
  
  if (col > sheet.getMaxColumns()) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), col - sheet.getMaxColumns());
  }
  
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