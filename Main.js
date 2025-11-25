/**
 * Main.gs
 * エントリーポイントとメニュー機能
 */

function onOpen() {
  // 起動時に設定シートを初期化（存在しない場合のみ）
  initializeSettingsSheet();

  SpreadsheetApp.getUi()
    .createMenu('✨ 課題自動取得システム')
    .addItem('1. 認証情報を設定', 'showCredentialDialog')
    .addItem('2. Tasks連携設定', 'setupTasksList')
    .addItem('3. **定期実行トリガーの自動設定**', 'setupDailyTrigger') // [改善案 4. 自動設定]
    .addSeparator()
    .addItem('4. 今すぐ実行（テスト）', 'dailySystemRun')
    .addToUi();
}

/**
 * 認証ダイアログ表示 (Setting.html を呼び出す)
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
    const taskListName = getSetting('TASKS_LIST_NAME'); // AppConfig.gs から読み込み
    const taskListId = getTaskListId(taskListName); // AppLogic.gs のヘルパー関数

    Props.setTaskListId(taskListId);
    ui.alert(`✅ 設定完了\nリスト「${taskListName}」と連携しました。`);
  } catch (e) {
    ui.alert(`🚨 エラー: ${e.message}\nTasks APIが有効か、または設定シートのTASKS_LIST_NAMEを確認してください。`);
  }
}

/**
 * 4. 定期実行トリガーを自動設定する関数
 */
function setupDailyTrigger() {
  const ui = SpreadsheetApp.getUi();
  const functionToRun = 'dailySystemRun';

  try {
    const triggerHour = getSetting('TRIGGER_HOUR'); // AppConfig.gs から読み込み

    // 既存のトリガーを全て削除（重複防止）
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === functionToRun) {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }

    // 新しい日次トリガーを設定
    ScriptApp.newTrigger(functionToRun)
      .timeBased()
      .everyDays(1)
      .atHour(triggerHour) // 設定シートの値を使用
      .create();

    ui.alert(`✅ 定期実行トリガーを設定しました。\n毎日午前${triggerHour}時〜${triggerHour + 1}時の間に自動実行されます。`);
  } catch (e) {
    ui.alert(`🚨 エラー: トリガー設定に失敗しました。\n設定シートの「TRIGGER_HOUR」が正しく設定されているか確認してください。\nエラー詳細: ${e.message}`);
  }
}

/**
 * 日次実行メイン関数
 */
function dailySystemRun() {
  log('=== システム実行開始 ===');
  try {
    processWebClass();
    processClassroom();
    processTasksSync();
    log('=== システム実行完了 ===');
  } catch (e) {
    log(`🚨 致命的エラー中断: ${e.toString()}\n認証情報やTasks連携設定を確認してください。`);
  }
}