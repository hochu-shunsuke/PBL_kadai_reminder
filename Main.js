/**
 * Main.gs
 * エントリーポイント
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('✨ 課題自動取得システム')
    .addItem('1. 認証情報を設定', 'showCredentialDialog')
    .addItem('2. Tasks連携設定', 'setupTasksList')
    .addSeparator()
    .addItem('3. 今すぐ実行（テスト）', 'dailySystemRun')
    .addToUi();
}

/**
 * 認証ダイアログ表示
 */
function showCredentialDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Setting')
    .setWidth(450).setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'WebClass認証情報の設定');
}

/**
 * Tasksリストのセットアップ
 */
function setupTasksList() {
  const ui = SpreadsheetApp.getUi();
  try {
    const lists = Tasks.Tasklists.list().items;
    let targetId = null;

    for (const list of lists) {
      if (list.title === TASK_LIST_NAME) {
        targetId = list.id;
        break;
      }
    }

    if (!targetId) {
      const newList = Tasks.Tasklists.insert({ title: TASK_LIST_NAME });
      targetId = newList.id;
    }

    Props.setTaskListId(targetId);
    ui.alert(`✅ 設定完了\nリスト「${TASK_LIST_NAME}」と連携しました。`);
  } catch (e) {
    ui.alert(`エラー: ${e.message}\nTasks APIが有効か確認してください。`);
  }
}

/**
 * 日次実行メイン関数
 */
function dailySystemRun() {
  log('---システム実行開始---');
  try {
    // 1. WebClass取得
    processWebClass();

    // 2. Classroom取得
    processClassroom();

    // 3. Tasks同期・登録・掃除
    processTasksSync();

    log('---システム実行完了---');
  } catch (e) {
    log(`🚨 致命的エラー中断: ${e.toString()}`);
  }
}