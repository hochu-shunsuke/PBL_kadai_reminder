/**
 * Main.gs
 * エントリーポイントとメニュー機能
 */

const SETUP_FUNCTION = 'dailySystemRun';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('✨ 課題自動取得システム')
    .addItem('1. WebClass 認証情報を設定 (ID/PW)', 'showAuthDialog')
    .addItem('2. Tasks・自動実行設定を完了 (初回のみ)', 'showTasksSetupDialog')
    .addSeparator()
    .addItem('3. 今すぐ実行（テスト）', SETUP_FUNCTION)
    .addSeparator()
    .addItem('4. 設定をすべてリセット', 'resetAllSettings')
    .addToUi();
}

/**
 * 1. 認証情報設定ダイアログ表示
 */
function showAuthDialog() {
  // Settings.htmlを読み込みます
  const html = HtmlService.createHtmlOutputFromFile('Setting')
    .setWidth(450).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, 'WebClass 認証情報の設定');
}

/**
 * 2. Tasks・自動実行設定ダイアログ表示
 */
function showTasksSetupDialog() {
  // Setting_Tasks.htmlを読み込みます
  const html = HtmlService.createHtmlOutputFromFile('Setting_Tasks')
    .setWidth(500).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Tasks・自動実行設定');
}

/**
 * Tasks設定保存後に呼ばれる処理（トリガー設定など）
 * Setting_Tasks.html から呼び出されます
 */
function runPostTasksSetup(settings) {
  const ui = SpreadsheetApp.getUi();
  try {
    // 1. Tasksリストの連携・作成はUtils.gsのsaveTasksDataFromHtml内で完了しています

    // 2. トリガー設定
    setupDailyTrigger(Number(settings.triggerHour));

    ui.alert(`✅ 設定完了\nTasksリスト「${settings.taskListName}」と連携し、毎日${settings.triggerHour}時台の自動実行を設定しました。`);
  } catch (e) {
    log(`🚨 設定エラー: ${e.message}`);
    throw e; // HTML側にエラーを返す
  }
}

/**
 * 定期実行トリガーの設定（最適化済み）
 */
function setupDailyTrigger(hour) {
  const triggers = ScriptApp.getProjectTriggers();
  let existingTrigger = null;

  for (const t of triggers) {
    if (t.getHandlerFunction() === SETUP_FUNCTION) {
      existingTrigger = t;
      break;
    }
  }

  // 既に同じ時間のトリガーがあれば何もしない
  const currentHour = Settings.getSetting('triggerHour');
  if (existingTrigger && currentHour == hour) {
    log('✅ トリガー設定スキップ: 変更なし');
    return;
  }

  // 古いトリガー削除
  if (existingTrigger) {
    ScriptApp.deleteTrigger(existingTrigger);
  }

  // 新規作成
  ScriptApp.newTrigger(SETUP_FUNCTION)
    .timeBased().everyDays(1).atHour(hour).create();
  log(`✅ 毎日${hour}時のトリガーを設定しました。`);
}

/**
 * 4. 設定をすべてリセット (Utils.gsのresetAllを呼び出し)
 */
function resetAllSettings() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '🚨 設定リセットの確認',
    'WebClass認証情報、TasksリストID、自動実行トリガーをすべて削除します。よろしいですか？\n\n（この操作は元に戻せません）',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    try {
      Settings.resetAll(); // Utils.gsの関数を呼び出し
      ui.alert('✅ すべての設定とトリガーを削除しました。システムを再利用するには、再度メニュー1, 2を実行してください。');
    } catch (e) {
      ui.alert(`🚨 リセットエラー: ${e.message}`);
    }
  }
}

/**
 * 日次実行メイン関数
 */
function dailySystemRun() {
  log('--- システム実行開始 ---');
  try {
    processWebClass();
    processClassroom();
    processTasksSync();
    log('--- システム実行完了 ---');
  } catch (e) {
    log(`🚨 実行中断: ${e.message}`);
  }
}