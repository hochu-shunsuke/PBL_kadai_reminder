/**
 * Utils.gs
 * システムの共通ヘルパー、設定管理、ログ、シート操作を担う。
 * * 依存: 
 * - Config.gs (定数)
 * - Main.gs (関数: runPostTasksSetup, 定数: SETUP_FUNCTION <--- Main.gsで定義されたものを利用)
 * - Tasks API サービス
 */


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
  // 1. Tasksリストの検索・作成（IDの取得と保存）をまず実行
  const listId = setupTasksList(settings.taskListName);
  Settings.setTaskListId(listId);

  // 2. その他の設定を保存
  Settings.saveTasks(settings);

  // 3. トリガー設定（Main.gsの関数呼び出し）
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
  getSetting: function (key) {
    return PropertiesService.getUserProperties().getProperty(key);
  },

  saveAuth: function (data) {
    PropertiesService.getUserProperties().setProperties({
      'userid': data.userid,
      'password': data.password
    });
  },

  saveTasks: function (data) {
    PropertiesService.getUserProperties().setProperties({
      'taskListName': String(data.taskListName),
      'triggerHour': String(data.triggerHour),
      'cleanupDays': String(data.cleanupDays)
    });
  },

  getTaskListId: function () {
    return PropertiesService.getUserProperties().getProperty('taskListId');
  },

  setTaskListId: function (id) {
    PropertiesService.getUserProperties().setProperty('taskListId', id);
  },

  /**
   * すべてのユーザープロパティと日次実行トリガーを削除する
   */
  resetAll: function () {
    PropertiesService.getUserProperties().deleteAllProperties();

    // SETUP_FUNCTION (dailySystemRun) は Main.gs で定義された定数を利用
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
 */
const SheetUtils = {
  writeToSheet: function (sheetName, dataRows) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      log(`シート「${sheetName}」を新規作成しました。`);
    }

    // --- 日付ソート処理 ---
    if (dataRows.length > 0) {
      dataRows.sort((a, b) => {
        const dateA = parseAssignmentDate(a[4]);
        const dateB = parseAssignmentDate(b[4]);

        const timeA = dateA ? dateA.getTime() : Infinity;
        const timeB = dateB ? dateB.getTime() : Infinity;

        return timeA - timeB;
      });
    }
    // ----------------------

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const lastColumn = sheet.getLastColumn();
      if (lastColumn > 0) {
        sheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
      }
    }

    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight('bold');

    if (dataRows.length > 0) {
      sheet.getRange(2, 1, dataRows.length, dataRows[0].length).setValues(dataRows);
    }
    SpreadsheetApp.flush();
    log(`✅ ${dataRows.length}件を「${sheetName}」へ書き込み完了`);
  }
};