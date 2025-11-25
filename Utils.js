/**
 * Utils.gs
 * 汎用ユーティリティ関数群
 */

/**
 * ログシートに記録し、Loggerにも出力する
 * @param {string} message 
 */
function log(message) {
  Logger.log(message);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME_LOG);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME_LOG);
      sheet.appendRow(['タイムスタンプ', 'メッセージ']);
    }
    sheet.appendRow([new Date(), message]);
  } catch (e) {
    console.error('ログ書き込みエラー: ' + e.message);
  }
}

/**
 * PropertiesService関連のラッパー
 */
const Props = {
  getCredentials: function() {
    const props = PropertiesService.getUserProperties().getProperties();
    if (props.userid && props.password) {
      return { userid: props.userid, password: props.password };
    }
    return null;
  },

  setCredentials: function(userid, password) {
    if (!userid || !password) throw new Error('IDとパスワードが必要です。');
    PropertiesService.getUserProperties().setProperties({
      'userid': userid,
      'password': password
    });
  },

  getTaskListId: function() {
    return PropertiesService.getUserProperties().getProperty('taskListId');
  },

  setTaskListId: function(id) {
    PropertiesService.getUserProperties().setProperty('taskListId', id);
  }
};

/**
 * シート関連の共通処理
 */
const SheetUtils = {
  /**
   * 指定したシート名にデータを上書き保存する（ヘッダー維持）
   */
  writeToSheet: function(sheetName, dataRows) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

    // データクリア（2行目以降）
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }

    // ヘッダー
    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight('bold');
    
    // データ書き込み
    if (dataRows.length > 0) {
      sheet.getRange(2, 1, dataRows.length, dataRows[0].length).setValues(dataRows);
    }
    SpreadsheetApp.flush();
    log(`✅ ${dataRows.length}件を「${sheetName}」に書き込みました。`);
  }
};

// --- Settings.htmlからの呼び出し用関数 ---
function setCredentials(userid, password) {
  Props.setCredentials(userid, password);
}