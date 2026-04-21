/**
 * ②番号参照の入力順を記録するシートのテンプレ名（各点検票ブックに1枚配置推奨。無い場合は入力シートをコピーして自動作成）
 */
var REF_ORDER_TEMPLATE_SHEET_NAME = '【テンプレ】番号参照入力順';

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
  settingsSheet.appendRow(["QUIZ_PF_SS_ID", "", "小テスト（合否）用スプレッドシートID"]);
  settingsSheet.appendRow(["QUIZ_SCORE_SS_ID", "", "小テスト（点数）用スプレッドシートID"]);
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
  let sampleAssignSs = SpreadsheetApp.create("【サンプル】課題点検票");
  DriveApp.getFileById(sampleAssignSs.getId()).moveTo(folder);
  createSampleSheet(sampleAssignSs.getSheets()[0], "1年＠入力", "assignment");
  ensureRefOrderTemplateInWorkbook_(sampleAssignSs);

  let sampleQuizPfSs = SpreadsheetApp.create("【サンプル】小テスト点検票(合否)");
  DriveApp.getFileById(sampleQuizPfSs.getId()).moveTo(folder);
  createSampleSheet(sampleQuizPfSs.getSheets()[0], "1年＠入力", "quiz_pf");
  ensureRefOrderTemplateInWorkbook_(sampleQuizPfSs);

  let sampleQuizScoreSs = SpreadsheetApp.create("【サンプル】小テスト点検票(点数)");
  DriveApp.getFileById(sampleQuizScoreSs.getId()).moveTo(folder);
  createSampleSheet(sampleQuizScoreSs.getSheets()[0], "1年＠入力", "quiz_score");
  ensureRefOrderTemplateInWorkbook_(sampleQuizScoreSs);

  // 4. スクリプトプロパティへの登録
  PropertiesService.getScriptProperties().setProperties({
    "CONFIG_SS_ID": configSs.getId()
  });
  
  // 初期設定としてサンプルをターゲットに設定
  updateConfigValue("ASSIGNMENT_SS_ID", sampleAssignSs.getId());
  updateConfigValue("QUIZ_PF_SS_ID", sampleQuizPfSs.getId());
  updateConfigValue("QUIZ_SCORE_SS_ID", sampleQuizScoreSs.getId());

  Logger.log("セットアップ完了！マイドライブの「課題点検アプリ_システムフォルダ」を確認してください。");
}

function createSampleSheet(sheet, desiredName = "1年＠入力", sheetType = "assignment") {
  try {
    sheet.setName(desiredName);
  } catch(e) {
    Logger.log("Failed to set sheet name: " + desiredName);
  }
  
  let a1Val = "提出物";
  if (sheetType === "quiz_pf") a1Val = "小テスト(合否)";
  if (sheetType === "quiz_score") a1Val = "小テスト(点数)";

  const headers = [
    ["組","","","","","","通し番号→", 1, 2, 3, 4, 5],
    ["","","","","","", "提出率→", "=(COUNTIF(H6:H50,\"提\"))/40", "", "", "", ""],
    ["","","","","","", "返却可否→", "済", "済", "済", "未", "未"],
    ["","","","日付","","", "提出日→", "4/10", "4/15", "4/20", "5/1", "5/10"],
    ["組","番号","ID","氏名","性別","提出率","提出数", "課題A", "課題B", "小テスト1", "課題C", "小テスト2"]
  ];
  sheet.getRange(1, 1, 5, headers[0].length).setValues(headers);
  sheet.getRange("A1").setValue(a1Val);

  if (sheetType === "quiz_pf" || sheetType === "quiz_score") {
    let passScore = 80;
    try { passScore = getAdminSettings().PASS_SCORE || 80; } catch(e) {}
    sheet.getRange("B1").setValue(passScore);
    sheet.getRange("C1").setValue("点合格");
  }
  
  let sampleStudents = [];
  for(let i=1; i<=40; i++) {
    let group = i <= 20 ? "1組" : "2組";
    let num = i <= 20 ? i : i - 20;
    let idStr = (group === "1組" ? "10" : "20") + ("0" + num).slice(-2);
    sampleStudents.push([
      group, 
      num, 
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
  let range = sheet.getRange("H6:Z205");
  let rules = [];
  if (sheetType === "assignment") {
    rules = [
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("提").setBackground("#b7e1cd").setFontColor("#0f5132").setRanges([range]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("未").setBackground("#f4c7c3").setFontColor("#842029").setRanges([range]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("再").setBackground("#fce8b2").setFontColor("#664d03").setRanges([range]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("休").setBackground("#d9d2e9").setFontColor("#351c75").setRanges([range]).build()
    ];
  } else if (sheetType === "quiz_pf") {
    rules = [
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("○").setBackground("#d1e7dd").setFontColor("#0f5132").setRanges([range]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("×").setBackground("#f8d7da").setFontColor("#842029").setRanges([range]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("休").setBackground("#d9d2e9").setFontColor("#351c75").setRanges([range]).build()
    ];
  } else if (sheetType === "quiz_score") {
    rules = [
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("休").setBackground("#d9d2e9").setFontColor("#351c75").setRanges([range]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=AND(ISNUMBER(H6), H6>=$B$1)").setBackground("#d1e7dd").setFontColor("#0f5132").setRanges([range]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=AND(ISNUMBER(H6), H6<$B$1)").setBackground("#f8d7da").setFontColor("#842029").setRanges([range]).build()
    ];
  }
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
}

function saveAdminSettingsFromUI(assignId, quizPfId, quizScoreId, mapName, mapId, mapClass, mapNum, rosterCols, headerRows) {
  updateConfigValue("ASSIGNMENT_SS_ID", assignId);
  updateConfigValue("QUIZ_PF_SS_ID", quizPfId);
  updateConfigValue("QUIZ_SCORE_SS_ID", quizScoreId);
  if (mapName !== "") updateConfigValue("COL_MAP_NAME", mapName);
  if (mapId !== "") updateConfigValue("COL_MAP_ID", mapId);
  if (mapClass !== "") updateConfigValue("COL_MAP_CLASS", mapClass);
  if (mapNum !== "") updateConfigValue("COL_MAP_NUMBER", mapNum);
  if (rosterCols !== "") updateConfigValue("ROSTER_COLS", rosterCols);
  if (headerRows !== "") updateConfigValue("HEADER_ROWS", headerRows);
  
  if (assignId) registerBookTypeMarker("assignment", assignId);
  if (quizPfId) registerBookTypeMarker("quiz_pf", quizPfId);
  if (quizScoreId) registerBookTypeMarker("quiz_score", quizScoreId);
  
  return true;
}

function getHeadersFromSheet(headerRow) {
  const settings = getAdminSettings();
  const assignId = settings.ASSIGNMENT_SS_ID || PropertiesService.getScriptProperties().getProperty("SAMPLE_SS_ID");
  if (!assignId) return [];
  
  try {
    const ss = SpreadsheetApp.openById(assignId);
    const sheet = ss.getSheets().find(s => s.getName().endsWith('＠入力'));
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
    let label = '提出物';
    if (bookType === 'quiz_pf') label = '小テスト(合否)';
    if (bookType === 'quiz_score') label = '小テスト(点数)';
    const sheets = ss.getSheets();
    for (let s of sheets) {
      if (s.getName().endsWith('＠入力')) {
        s.getRange("A1").setValue(label);
      }
    }
  } catch(e) {
    Logger.log("Failed to register A1 marker: " + e);
  }
}

/** @deprecated 入力途中状態はブラウザ localStorage に保存（ユーザー単位ではなく端末・ブラウザ単位） */
function saveUserState(stateObj) {}

function loadUserState() {
  return null;
}

function clearUserState() {
  PropertiesService.getUserProperties().deleteProperty("APP_STATE");
}

/** 「Copy of 」「 のコピー」などを除いた論理シート名（重複シートの統合用） */
function stripCopyOfPrefix_(name) {
  let n = String(name).trim();
  while (/^Copy of /i.test(n)) n = n.replace(/^Copy of /i, "");
  while (/ のコピー$/.test(n)) n = n.replace(/ のコピー$/, "");
  return n;
}

function dedupeRefOrderTemplateSheets_(ss) {
  if (!ss) return;
  const all = ss.getSheets();
  const candidates = [];
  for (let i = 0; i < all.length; i++) {
    const sh = all[i];
    const stripped = stripCopyOfPrefix_(sh.getName());
    if (stripped === REF_ORDER_TEMPLATE_SHEET_NAME) candidates.push(sh);
  }
  if (candidates.length === 0) return;
  const exact = ss.getSheetByName(REF_ORDER_TEMPLATE_SHEET_NAME);
  if (exact) {
    for (let j = 0; j < candidates.length; j++) {
      if (candidates[j].getSheetId() !== exact.getSheetId()) {
        try {
          ss.deleteSheet(candidates[j]);
        } catch (e) {
          Logger.log("dedupeRefOrderTemplateSheets_ delete: " + e);
        }
      }
    }
    return;
  }
  const keeper = candidates[0];
  for (let k = 1; k < candidates.length; k++) {
    try {
      ss.deleteSheet(candidates[k]);
    } catch (e) {
      Logger.log("dedupeRefOrderTemplateSheets_ delete dup: " + e);
    }
  }
  try {
    keeper.setName(REF_ORDER_TEMPLATE_SHEET_NAME);
  } catch (e) {
    Logger.log("dedupeRefOrderTemplateSheets_ rename: " + e);
  }
}

function orderSheetsMatchingClassPrefix_(ss, classPrefix) {
  const out = [];
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const sh = sheets[i];
    const norm = stripCopyOfPrefix_(sh.getName());
    if (!norm.endsWith("＠番号参照入力順")) continue;
    const p = norm.replace(/＠番号参照入力順$/, "");
    if (p === classPrefix) out.push(sh);
  }
  return out;
}

function scoreOrderSheetForMerge_(sheet) {
  try {
    const settings = getAdminSettings();
    const rosterCols = parseInt(settings.ROSTER_COLS) || 7;
    const data = sheet.getDataRange().getValues();
    let score = 0;
    for (let r = 0; r < data.length; r++) {
      for (let c = rosterCols; c < data[r].length; c++) {
        const v = data[r][c];
        if (v !== "" && v !== null && v !== undefined) score++;
      }
    }
    return score;
  } catch (e) {
    return 0;
  }
}

/**
 * 同一クラス（＠より前が同じ）の「＠番号参照入力順」シートを1枚にまとめる。
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null} 残すシート。無ければ null
 */
function dedupeClassOrderSheetsForPrefix_(ss, canonicalOrderName) {
  if (!ss || !canonicalOrderName || !canonicalOrderName.endsWith("＠番号参照入力順")) return null;
  const classPrefix = canonicalOrderName.replace(/＠番号参照入力順$/, "");
  const candidates = orderSheetsMatchingClassPrefix_(ss, classPrefix);
  if (candidates.length === 0) return null;
  if (candidates.length === 1) {
    const only = candidates[0];
    if (only.getName() !== canonicalOrderName && !ss.getSheetByName(canonicalOrderName)) {
      try {
        only.setName(canonicalOrderName);
      } catch (e) {
        Logger.log("dedupeClassOrderSheetsForPrefix_ rename single: " + e);
      }
    }
    return only;
  }
  let keeper = null;
  for (let i = 0; i < candidates.length; i++) {
    if (stripCopyOfPrefix_(candidates[i].getName()) === canonicalOrderName) {
      keeper = candidates[i];
      break;
    }
  }
  if (!keeper) {
    keeper = candidates[0];
    let best = scoreOrderSheetForMerge_(keeper);
    for (let j = 1; j < candidates.length; j++) {
      const sc = scoreOrderSheetForMerge_(candidates[j]);
      if (sc > best) {
        best = sc;
        keeper = candidates[j];
      }
    }
  }
  for (let k = 0; k < candidates.length; k++) {
    if (candidates[k].getSheetId() === keeper.getSheetId()) continue;
    try {
      ss.deleteSheet(candidates[k]);
    } catch (e) {
      Logger.log("dedupeClassOrderSheetsForPrefix_ delete: " + e);
    }
  }
  if (keeper.getName() !== canonicalOrderName) {
    try {
      if (!ss.getSheetByName(canonicalOrderName)) keeper.setName(canonicalOrderName);
    } catch (e) {
      Logger.log("dedupeClassOrderSheetsForPrefix_ rename keeper: " + e);
    }
  }
  return keeper;
}

function ensureRefOrderTemplateInWorkbook_(ss) {
  try {
    if (!ss) return;
    dedupeRefOrderTemplateSheets_(ss);
    if (ss.getSheetByName(REF_ORDER_TEMPLATE_SHEET_NAME)) return;
    const inputs = ss.getSheets().filter(s => /＠入力$/.test(s.getName()));
    if (!inputs.length) return;
    inputs[0].copyTo(ss).setName(REF_ORDER_TEMPLATE_SHEET_NAME);
  } catch (e) {
    Logger.log("ensureRefOrderTemplateInWorkbook_: " + e);
  }
}

function inputSheetNameToOrderSheetName_(inputSheetName) {
  const n = String(inputSheetName);
  if (n.indexOf("＠入力") === -1) return n + "＠番号参照入力順";
  return n.replace("＠入力", "＠番号参照入力順");
}

function escapeSheetNameForFormula_(sheetName) {
  return "'" + String(sheetName).replace(/'/g, "''") + "'";
}

function columnToLetter_(column) {
  let temp = "";
  let col = column;
  while (col > 0) {
    let rem = (col - 1) % 26;
    temp = String.fromCharCode(65 + rem) + temp;
    col = Math.floor((col - 1) / 26);
  }
  return temp;
}

function syncRefOrderSheetFromMain_(mainSheet, orderSheet, syncOptions) {
  syncOptions = syncOptions || {};
  const settings = getAdminSettings();
  const headerRows = parseInt(settings.HEADER_ROWS) || 5;
  const rosterCols = parseInt(settings.ROSTER_COLS) || 7;
  const mapName = parseInt(settings.COL_MAP_NAME);
  const colName = isNaN(mapName) ? 3 : mapName;
  const mainName = escapeSheetNameForFormula_(mainSheet.getName());
  const lastCol = Math.max(mainSheet.getLastColumn(), rosterCols + 1);
  const lastRow = Math.max(mainSheet.getLastRow(), headerRows + 1);

  for (let r = 1; r <= headerRows; r++) {
    for (let c = 1; c <= lastCol; c++) {
      const a1 = columnToLetter_(c) + r;
      orderSheet.getRange(r, c).setFormula("=" + mainName + "!" + a1);
    }
  }
  for (let r = headerRows + 1; r <= lastRow; r++) {
    for (let c = 1; c <= rosterCols; c++) {
      const a1 = columnToLetter_(c) + r;
      orderSheet.getRange(r, c).setFormula("=" + mainName + "!" + a1);
    }
  }
  if (syncOptions.clearTaskBody && lastCol > rosterCols && lastRow > headerRows) {
    orderSheet.getRange(headerRows + 1, rosterCols + 1, lastRow, lastCol).clearContent();
  }
  applyOrderSheetNameStyle_(orderSheet, headerRows, rosterCols, colName, lastRow);
}

function applyOrderSheetNameStyle_(orderSheet, headerRows, rosterCols, colNameZeroBased, lastRow) {
  const col = colNameZeroBased + 1;
  if (lastRow <= headerRows) return;
  const range = orderSheet.getRange(headerRows + 1, col, lastRow, col);
  range.setFontWeight("bold");
  range.setFontColor("#1155cc");
  range.setFontSize(12);
}

function ensureRefOrderSheet(bookType, classNameJSON) {
  const ss = getAppSpreadsheet(bookType);
  let target;
  try {
    target = JSON.parse(classNameJSON);
  } catch (e) {
    target = { sheetName: classNameJSON, group: "" };
  }
  const mainName = target.sheetName;
  const mainSheet = ss.getSheetByName(mainName);
  if (!mainSheet) throw new Error("入力シートがありません: " + mainName);

  ensureRefOrderTemplateInWorkbook_(ss);

  const orderName = inputSheetNameToOrderSheetName_(mainName);
  let orderSheet = dedupeClassOrderSheetsForPrefix_(ss, orderName);
  const isNewOrderSheet = !orderSheet;
  if (!orderSheet) {
    const template = ss.getSheetByName(REF_ORDER_TEMPLATE_SHEET_NAME);
    if (template) {
      orderSheet = template.copyTo(ss);
      orderSheet.setName(orderName);
    } else {
      orderSheet = mainSheet.copyTo(ss);
      orderSheet.setName(orderName);
    }
  }
  syncRefOrderSheetFromMain_(mainSheet, orderSheet, { clearTaskBody: isNewOrderSheet });
  return { ok: true, orderSheetName: orderName };
}

function writeProcessOrderToRefOrderSheet_(orderSheet, headerRows, taskCol, mainLastRow, processOrderByRow, partial) {
  if (!processOrderByRow || typeof processOrderByRow !== "object") return;
  const keys = Object.keys(processOrderByRow);
  if (!keys.length) return;
  if (partial) {
    keys.forEach(rowStr => {
      const row = parseInt(rowStr, 10);
      if (!isNaN(row) && row > headerRows && row <= mainLastRow) {
        orderSheet.getRange(row, taskCol).clearContent();
      }
    });
    keys.forEach(rowStr => {
      const row = parseInt(rowStr, 10);
      if (!isNaN(row) && row > headerRows && row <= mainLastRow) {
        orderSheet.getRange(row, taskCol).setValue(processOrderByRow[rowStr]);
      }
    });
  } else {
    if (mainLastRow > headerRows) {
      orderSheet.getRange(headerRows + 1, taskCol, mainLastRow, taskCol).clearContent();
    }
    keys.forEach(rowStr => {
      const row = parseInt(rowStr, 10);
      if (!isNaN(row) && row > headerRows && row <= mainLastRow) {
        orderSheet.getRange(row, taskCol).setValue(processOrderByRow[rowStr]);
      }
    });
  }
}

function getAppSpreadsheet(bookType) {
  const settings = getAdminSettings();
  let targetId;
  if (bookType === "quiz_pf") targetId = settings["QUIZ_PF_SS_ID"];
  else if (bookType === "quiz_score") targetId = settings["QUIZ_SCORE_SS_ID"];
  else targetId = settings["ASSIGNMENT_SS_ID"];
  return SpreadsheetApp.openById(targetId);
}

function getClassList(bookType) {
  const ss = getAppSpreadsheet(bookType);
  const sheets = ss.getSheets().filter(s => s.getName().endsWith('＠入力'));
  
  const settings = getAdminSettings();
  const headerRows = parseInt(settings.HEADER_ROWS) || 5;
  const mapClass = parseInt(settings.COL_MAP_CLASS) || 0;
  
  let classList = [];
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const lastRow = sheet.getLastRow();
    if (lastRow <= headerRows) {
      classList.push(JSON.stringify({sheetName: sheetName, group: "", displayName: sheetName}));
      return;
    }
    
    const dataRange = sheet.getRange(headerRows + 1, mapClass + 1, lastRow - headerRows, 1).getValues();
    let groups = new Set();
    dataRange.forEach(row => {
      if (row[0] !== "") groups.add(String(row[0]));
    });
    
    if (groups.size === 0 || (groups.size === 1 && groups.has(""))) {
      classList.push(JSON.stringify({sheetName: sheetName, group: "", displayName: sheetName.replace("＠入力", "")}));
    } else if (groups.size === 1) {
      const g = Array.from(groups)[0];
      classList.push(JSON.stringify({sheetName: sheetName, group: g, displayName: `${sheetName.replace("＠入力", "")} - ${g}`}));
    } else {
      Array.from(groups).sort().forEach(g => {
        classList.push(JSON.stringify({sheetName: sheetName, group: g, displayName: `${sheetName.replace("＠入力", "")} - ${g}`}));
      });
    }
  });
  
  return classList;
}

function getClassData(bookType, classNameJSON) {
  const ss = getAppSpreadsheet(bookType);
  let target;
  try {
    target = JSON.parse(classNameJSON);
  } catch(e) {
    target = { sheetName: classNameJSON, group: "" };
  }
  
  const sheet = ss.getSheetByName(target.sheetName);
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
  if (lastRow <= headerRows) return { students: [], tasks: [] };
  
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
  
  let students = [];
  for (let r = 0; r < range.length; r++) {
    const row = range[r];
    const groupVal = String(row[colClass]);
    if (!row[colName]) continue;
    
    if (target.group && groupVal !== target.group) continue;
    
    students.push({
      sheetRow: headerRows + 1 + r,
      group: groupVal,
      no: row[colNum],
      id: row[colId],
      name: row[colName]
    });
  }

  return { students: students, tasks: taskList };
}

/**
 * 選択中の課題列について、名簿の行順に一致するセル値を返す（2回目以降の入力で既存入力を維持するため）
 */
function getTaskColumnValues(bookType, classNameJSON, taskIndex) {
  const data = getClassData(bookType, classNameJSON);
  if (!data || !data.students.length) return [];

  const ss = getAppSpreadsheet(bookType);
  let target;
  try {
    target = JSON.parse(classNameJSON);
  } catch (e) {
    target = { sheetName: classNameJSON, group: "" };
  }
  const sheet = ss.getSheetByName(target.sheetName);
  if (!sheet) return data.students.map(() => "");

  const settings = getAdminSettings();
  const rosterCols = parseInt(settings.ROSTER_COLS) || 7;
  const col = rosterCols + 1 + parseInt(taskIndex, 10);
  if (isNaN(col) || col < rosterCols + 1) return data.students.map(() => "");

  if (col > sheet.getLastColumn()) {
    return data.students.map(() => "");
  }

  const rows = data.students.map(s => s.sheetRow);
  const minR = Math.min.apply(null, rows);
  const maxR = Math.max.apply(null, rows);
  const height = maxR - minR + 1;
  const block = sheet.getRange(minR, col, height, 1).getValues();
  const byRow = {};
  for (let i = 0; i < height; i++) {
    byRow[minR + i] = block[i][0];
  }

  return data.students.map(s => normalizeTaskCellForClient_(byRow[s.sheetRow]));
}

function normalizeTaskCellForClient_(v) {
  if (v === null || v === undefined || v === "") return "";
  if (Object.prototype.toString.call(v) === "[object Date]") {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "M/d");
  }
  return v;
}

function createNewTask(bookType, classNameJSON, dateStr, name) {
  const ss = getAppSpreadsheet(bookType);
  let target;
  try { target = JSON.parse(classNameJSON); } catch(e) { target = { sheetName: classNameJSON }; }
  const sheet = ss.getSheetByName(target.sheetName);
  
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
  
  logToConfig(target.sheetName + (target.group ? ` (${target.group})` : ""), "新規作成", `ブック: ${bookType}, 課題名: ${name}, 日付: ${dateStr}`);
  
  return true;
}

function submitAttendanceData(bookType, classNameJSON, taskIndex, taskName, resultsWithRow, processOrderByRow, options) {
  options = options || {};
  const partialValues = !!options.partialValues;
  const partialProcessOrder = !!options.partialProcessOrder;
  const writeProcessOrder = !!options.writeProcessOrder;
  const skipClearUserState = !!options.skipClearUserState;

  const ss = getAppSpreadsheet(bookType);
  let target;
  try { target = JSON.parse(classNameJSON); } catch(e) { target = { sheetName: classNameJSON }; }
  const sheet = ss.getSheetByName(target.sheetName);
  
  const settings = getAdminSettings();
  const headerRows = parseInt(settings.HEADER_ROWS) || 5;
  const rosterCols = parseInt(settings.ROSTER_COLS) || 7;
  
  const col = rosterCols + 1 + parseInt(taskIndex);
  
  if (col > sheet.getMaxColumns()) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), col - sheet.getMaxColumns());
  }
  
  if (taskName) {
    const taskNameCell = sheet.getRange(headerRows, col);
    if (!taskNameCell.getValue()) {
      taskNameCell.setValue(taskName);
    }
  }

  let oldValues = {};
  let minRow = Infinity;
  let maxRow = -Infinity;
  resultsWithRow.forEach(item => {
    if (item.sheetRow < minRow) minRow = item.sheetRow;
    if (item.sheetRow > maxRow) maxRow = item.sheetRow;
  });

  if (minRow !== Infinity && maxRow !== -Infinity && col <= sheet.getLastColumn()) {
    const numRows = maxRow - minRow + 1;
    const currentData = sheet.getRange(minRow, col, numRows, 1).getValues();
    for (let i = 0; i < currentData.length; i++) {
      oldValues[minRow + i] = currentData[i][0];
    }
  }

  let changeLogs = [];

  resultsWithRow.forEach(item => {
    if (item.val !== null && item.val !== undefined && item.val !== "") {
      const oldVal = oldValues[item.sheetRow];
      const oldValStr = (oldVal === undefined || oldVal === null || oldVal === "") ? "空欄" : oldVal;
      const newValStr = item.val;
      
      if (String(oldValStr) !== String(newValStr)) {
        changeLogs.push(`${item.name || '不明'}(${oldValStr}→${newValStr})`);
      }
      
      sheet.getRange(item.sheetRow, col).setValue(item.val);
    }
  });

  if (writeProcessOrder && processOrderByRow && typeof processOrderByRow === "object" && Object.keys(processOrderByRow).length > 0) {
    const orderName = inputSheetNameToOrderSheetName_(target.sheetName);
    const orderSheet = ss.getSheetByName(orderName);
    if (orderSheet) {
      const mainLastRow = sheet.getLastRow();
      writeProcessOrderToRefOrderSheet_(orderSheet, headerRows, col, mainLastRow, processOrderByRow, partialProcessOrder);
    }
  }
  
  const logDetail = `課題列: ${parseInt(taskIndex)+1}, 課題名: ${taskName || '既存'}, 変更: ${changeLogs.length > 0 ? changeLogs.join(", ") : "なし"}${partialValues ? " (部分確定)" : ""}`;
  logToConfig(target.sheetName + (target.group ? ` (${target.group})` : ""), bookType === 'quiz' ? "小テスト入力" : "課題入力", logDetail);
  
  return ss.getUrl();
}

function saveSingleProcessOrder(bookType, classNameJSON, taskIndex, sheetRow, processOrder) {
  const ss = getAppSpreadsheet(bookType);
  let target;
  try { target = JSON.parse(classNameJSON); } catch (e) { target = { sheetName: classNameJSON }; }
  const sheet = ss.getSheetByName(target.sheetName);
  if (!sheet) throw new Error("対象シートが見つかりません。");

  const orderName = inputSheetNameToOrderSheetName_(target.sheetName);
  const orderSheet = ss.getSheetByName(orderName);
  if (!orderSheet) throw new Error("番号参照入力順シートが見つかりません。");

  const settings = getAdminSettings();
  const rosterCols = parseInt(settings.ROSTER_COLS) || 7;
  const col = rosterCols + 1 + parseInt(taskIndex, 10);
  const row = parseInt(sheetRow, 10);
  const orderVal = Number(processOrder);

  if (isNaN(row) || row < 1) throw new Error("行番号が不正です。");
  if (isNaN(orderVal)) throw new Error("処理順番が不正です。");

  if (col > orderSheet.getMaxColumns()) {
    orderSheet.insertColumnsAfter(orderSheet.getMaxColumns(), col - orderSheet.getMaxColumns());
  }
  orderSheet.getRange(row, col).setValue(orderVal);
  return true;
}

function assertAdmin_() {
  const email = Session.getActiveUser().getEmail();
  if (getUserRole(email) !== "admin") {
    throw new Error("この操作は管理者のみ実行できます。");
  }
}

function createSecondaryCheckSheet(bookType, classNameJSON, taskIndex, taskName, sortDirection) {
  assertAdmin_();
  const ss = getAppSpreadsheet(bookType);
  let target;
  try { target = JSON.parse(classNameJSON); } catch(e) { target = { sheetName: classNameJSON, group: "" }; }
  const sheet = ss.getSheetByName(target.sheetName);
  if (!sheet) throw new Error("対象シートが見つかりません。");

  const orderName = inputSheetNameToOrderSheetName_(target.sheetName);
  const orderSheet = ss.getSheetByName(orderName);
  if (!orderSheet) throw new Error("番号参照入力順シートがありません。②番号参照を開始してシートを作成してください。");

  const settings = getAdminSettings();
  const headerRows = parseInt(settings.HEADER_ROWS) || 5;
  const rosterCols = parseInt(settings.ROSTER_COLS) || 7;
  const taskCol = rosterCols + 1 + parseInt(taskIndex);

  const mapClass = parseInt(settings.COL_MAP_CLASS);
  const mapNum = parseInt(settings.COL_MAP_NUMBER);
  const mapName = parseInt(settings.COL_MAP_NAME);
  const colClass = isNaN(mapClass) ? 0 : mapClass;
  const colNum = isNaN(mapNum) ? 1 : mapNum;
  const colName = isNaN(mapName) ? 3 : mapName;

  const lastRow = Math.max(sheet.getLastRow(), orderSheet.getLastRow());
  if (lastRow <= headerRows) throw new Error("名簿行がありません。");

  const orderVals = orderSheet.getRange(headerRows + 1, taskCol, lastRow, taskCol).getValues();
  const processOrderByRow = {};
  for (let i = 0; i < orderVals.length; i++) {
    const v = orderVals[i][0];
    const rowNum = headerRows + 1 + i;
    if (v !== "" && v !== null && !isNaN(Number(v))) {
      processOrderByRow[String(rowNum)] = Number(v);
    }
  }
  if (Object.keys(processOrderByRow).length === 0) {
    throw new Error("番号参照入力順シートに処理順番がありません。");
  }

  const rowsMap = {};
  if (taskCol <= sheet.getLastColumn()) {
    const vals = sheet.getRange(headerRows + 1, taskCol, lastRow, taskCol).getValues();
    for (let i = 0; i < vals.length; i++) {
      rowsMap[headerRows + 1 + i] = vals[i][0];
    }
  }

  const records = [];
  Object.keys(processOrderByRow).forEach(rowStr => {
    const rowNum = parseInt(rowStr, 10);
    if (isNaN(rowNum) || rowNum <= headerRows) return;
    const rowVals = sheet.getRange(rowNum, 1, 1, Math.max(colName, colNum, colClass) + 1).getValues()[0];
    records.push({
      processOrder: Number(processOrderByRow[rowStr]),
      rowNum: rowNum,
      cls: rowVals[colClass],
      no: rowVals[colNum],
      name: rowVals[colName],
      value: rowsMap[rowNum] !== undefined && rowsMap[rowNum] !== null ? rowsMap[rowNum] : ""
    });
  });

  records.sort((a, b) => {
    const av = isNaN(a.processOrder) ? Number.MAX_SAFE_INTEGER : a.processOrder;
    const bv = isNaN(b.processOrder) ? Number.MAX_SAFE_INTEGER : b.processOrder;
    if (av !== bv) return sortDirection === 'desc' ? bv - av : av - bv;
    const an = Number(a.no);
    const bn = Number(b.no);
    if (!isNaN(an) && !isNaN(bn)) return an - bn;
    return String(a.no).localeCompare(String(b.no), 'ja');
  });

  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMdd_HHmm");
  const baseName = (taskName || "入力") + "_教員二次点検";
  let newName = `${baseName}_${sortDirection === 'desc' ? '降順' : '昇順'}_${ts}`;
  if (newName.length > 90) newName = newName.substring(0, 90);
  const output = ss.insertSheet(newName);
  output.getRange(1, 1, 1, 6).setValues([["処理順番", "組", "番号", "氏名", "入力値", "元シート行"]]);
  if (records.length > 0) {
    const values = records.map(r => [r.processOrder, r.cls, r.no, r.name, r.value, r.rowNum]);
    output.getRange(2, 1, values.length, 6).setValues(values);
  }
  output.setFrozenRows(1);
  output.autoResizeColumns(1, 6);

  logToConfig(target.sheetName + (target.group ? ` (${target.group})` : ""), "二次点検票作成", `課題列:${parseInt(taskIndex)+1}, 課題名:${taskName || '既存'}, 並び:${sortDirection}`);
  return ss.getUrl();
}

function importRosterFromTSV(bookType, targetSheetName, parsedData, mapping) {
  const ss = getAppSpreadsheet(bookType);
  if (!ss) throw new Error("対象ブックが設定されていません。");
  
  let sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(targetSheetName);
  } else {
    sheet.clear();
  }
  
  createSampleSheet(sheet, targetSheetName, bookType);
  
  const settings = getAdminSettings();
  const mapName = parseInt(settings.COL_MAP_NAME) || 3;
  const mapId = parseInt(settings.COL_MAP_ID) || 2;
  const mapClass = parseInt(settings.COL_MAP_CLASS) || 0;
  const mapNum = parseInt(settings.COL_MAP_NUMBER) || 1;
  const rosterCols = parseInt(settings.ROSTER_COLS) || 7;
  
  const targetCols = {
    'name': mapName,
    'id': mapId,
    'group': mapClass,
    'no': mapNum,
    'gender': 4 
  };
  
  const startRow = 6;
  
  // 一旦サンプルデータをクリア（数式ごと上書きするため全体をクリアするより、必要な列だけ入れる）
  sheet.getRange(startRow, 1, 40, rosterCols).clearContent();
  
  if (parsedData.length === 0) return true;
  
  let outputData = [];
  for (let r = 0; r < parsedData.length; r++) {
    let rowOut = new Array(rosterCols).fill("");
    const rowIn = parsedData[r];
    
    for (let c = 0; c < mapping.length; c++) {
      const field = mapping[c];
      if (field !== 'ignore' && targetCols[field] !== undefined) {
        if (c < rowIn.length) {
          rowOut[targetCols[field]] = rowIn[c];
        }
      }
    }
    
    let rowIndex = startRow + r;
    rowOut[5] = `=IF(COUNTA($H$5:$Z$5)=0, 0, G${rowIndex}/COUNTA($H$5:$Z$5))`;
    rowOut[6] = `=COUNTIF(H${rowIndex}:Z${rowIndex},"提")+COUNTIF(H${rowIndex}:Z${rowIndex}, 1)`;
    
    outputData.push(rowOut);
  }
  
  sheet.getRange(startRow, 1, outputData.length, rosterCols).setValues(outputData);
  sheet.getRange(startRow, 6, outputData.length, 1).setNumberFormat("0.0%");
  
  ensureRefOrderTemplateInWorkbook_(ss);
  return true;
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