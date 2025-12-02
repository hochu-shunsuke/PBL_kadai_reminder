/**
 * Utils.gs
 * システムの共通ヘルパー、設定管理、ログ、シート操作を担う。
 * * 依存: 
 * - Config.gs (定数)
 * - Main.gs (関数: runPostTasksSetup, 定数: SETUP_FUNCTION)
 * - Tasks API サービス
 */

// SETUP_FUNCTION は Main.gs で定義されています。（ここでは再宣言しない）


/**
 * HTMLから呼び出される認証情報保存関数
 */
function saveAuthDataFromHtml(userid, password) {
  if (!userid || !password) throw new Error('IDとパスワードが必要です。');
  Settings.saveAuth({ userid: String(userid), password: String(password) });
  return true; 
}

/**
 * HTMLから呼び出されるTasks設定保存関数
 */
function saveTasksDataFromHtml(settings) {
  const oldListId = Settings.getTaskListId(); 
  
  // 1. Tasksリストの検索・作成（IDの取得と保存）をまず実行
  const listId = setupTasksList(settings.taskListName); 
  Settings.setTaskListId(listId); 
  
  // 2. ★修正ロジック: Tasks IDが変わった場合、シートの全課題データをクリア
  if (oldListId !== listId) {
      log('🚨 TasksリストIDが変更されました。新しい環境で再スタートするため、シートの全課題データをクリアし、次回の実行で全て再取得します。');
      // AppLogic.gs で定義された関数を呼び出す
      clearAssignmentSheets(); 
  }
  
  // 3. その他の設定を保存
  Settings.saveTasks(settings);

  // 4. トリガー設定（Main.gsの関数呼び出し）
  runPostTasksSetup(settings);
  return true;
}

/**
 * HTML表示用の設定値取得
 */
function getTasksSettingsForHtml() {
  return {
    taskListName: Settings.getSetting('taskListName') || '大学課題',
    triggerHour: Settings.getSetting('triggerHour') || '6',
    cleanupDays: Settings.getSetting('cleanupDays') || '30'
  };
}


// --- 共通ヘルパー関数 ---

/**
 * ログ記録
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
    console.error('ログ記録エラー: ' + e.message);
  }
}

/**
 * 日付文字列をパースしてDateオブジェクトを返すヘルパー関数。
 */
function parseAssignmentDate(dateStr) {
  if (!dateStr) return null;
  const cleanStr = String(dateStr).trim().replace(/(\d{4})[\/年](\d{1,2})[\/月](\d{1,2})[\日]?/g, '$1/$2/$3');
  const date = new Date(cleanStr);
  return isNaN(date.getTime()) ? null : date;
}


// --- 設定管理オブジェクト ---

/**
 * 設定管理オブジェクト
 */
const Settings = {
  getSetting: function(key) {
    return PropertiesService.getUserProperties().getProperty(key);
  },
  
  saveAuth: function(data) {
    PropertiesService.getUserProperties().setProperties({
      'userid': data.userid,
      'password': data.password
    });
  },

  saveTasks: function(data) {
    PropertiesService.getUserProperties().setProperties({
      'taskListName': String(data.taskListName),
      'triggerHour': String(data.triggerHour),
      'cleanupDays': String(data.cleanupDays)
    });
  },

  getTaskListId: function() {
    return PropertiesService.getUserProperties().getProperty('taskListId');
  },

  setTaskListId: function(id) {
    PropertiesService.getUserProperties().setProperty('taskListId', id);
  },
  
  /**
   * TasksリストIDをPropertiesServiceから明示的に削除する
   */
  deleteTaskListId: function() {
    PropertiesService.getUserProperties().deleteProperty('taskListId');
  },
  
  /**
   * すべてのユーザープロパティと日次実行トリガーを削除する
   */
  resetAll: function() {
    PropertiesService.getUserProperties().deleteAllProperties();
    
    const triggers = ScriptApp.getProjectTriggers();
    for (const t of triggers) {
      if (t.getHandlerFunction() === SETUP_FUNCTION) {
          ScriptApp.deleteTrigger(t);
      }
    }
    log('すべての設定と自動実行トリガーを削除しました。');
  }
};


// --- Tasks連携ヘルパー ---

/**
 * Tasksリストの検索・作成
 */
function setupTasksList(listName) {
  const lists = Tasks.Tasklists.list().items;
  let targetId = null;
  
  // 1. 既存のリストを名前で検索
  if (lists) {
    for (const list of lists) {
      if (list.title === listName) {
        targetId = list.id;
        log(`既存のTasksリスト「${listName}」を再発見しました。`);
        break;
      }
    }
  }
  
  // 2. 見つからなければ新規作成
  if (!targetId) {
    const newList = Tasks.Tasklists.insert({ title: listName });
    targetId = newList.id;
    log(`Tasksリスト「${listName}」を新規作成しました。`);
  }

  return targetId; 
}


// --- シート操作ヘルパー ---

/**
 * シート書き込み共通処理
 * ★Tasks ID/フラグの復元ロジックを削除し、純粋に新しいデータでシートを上書きする
 */
const SheetUtils = {
  writeToSheet: function(sheetName, newAssignments) { // newAssignmentsはWebClass/Classroomから取得したデータ
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        log(`シート「${sheetName}」を新規作成しました。`);
    }

    // --- 既存の Tasks ID/Flag マージロジックを削除 ---
    
    // 1. シートをクリア
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const lastColumn = sheet.getLastColumn();
      if (lastColumn > 0) {
          sheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
      }
    }

    // 2. ヘッダーとデータを書き込み
    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight('bold');
    
    if (newAssignments.length > 0) {
      sheet.getRange(2, 1, newAssignments.length, newAssignments[0].length).setValues(newAssignments);
    }
    SpreadsheetApp.flush();
    log(`✅ ${newAssignments.length}件を「${sheetName}」へ書き込み完了 (Tasksフラグは空で上書き)`);
  }
};
