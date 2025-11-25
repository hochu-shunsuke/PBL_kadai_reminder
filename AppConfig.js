/**
 * AppConfig.gs
 * [改善案 1. 設定シートによる抽象化] の実装
 * スプレッドシートの「設定」シートを管理し、設定値を抽象化します。
 */

const SETTING_SHEET_NAME = '設定';

/**
 * 設定シートが存在しない場合、作成し初期値を設定する
 */
function initializeSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SETTING_SHEET_NAME);
  if (sheet) return sheet;

  sheet = ss.insertSheet(SETTING_SHEET_NAME, 0); // 先頭に挿入

  // 設定項目の初期化
  const settingsData = [
    ['設定項目 (KEY)', '値', '説明'],
    ['TASKS_LIST_NAME', '大学課題', '課題を登録するGoogle Tasksのリスト名。'],
    ['TRIGGER_HOUR', '6', '自動実行を開始する時間帯（0-23）。例: 6は午前6時〜7時の間に実行。'],
    ['CLEANUP_DAYS', '30', '完了・削除・期限切れの課題をシートから完全に削除するまでの猶予日数。'], // [改善案 9. 削除閾値]
  ];

  sheet.getRange(1, 1, settingsData.length, settingsData[0].length).setValues(settingsData);
  sheet.getRange('A1:C1').setFontWeight('bold').setBackground('#ddd');

  sheet.setColumnWidth(1, 150).setColumnWidth(2, 200).setColumnWidth(3, 400);

  SpreadsheetApp.flush();
  return sheet;
}

/**
 * 指定した設定値を取得する
 * @param {string} key - 設定項目のキー (例: 'TASKS_LIST_NAME')
 * @returns {string|number} 設定値
 */
function getSetting(key) {
  const sheet = initializeSettingsSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      const value = data[i][1];
      // 空文字や数値以外は文字列として、数値として扱えるものは数値として返す
      return isNaN(Number(value)) || value === '' ? String(value).trim() : Number(value);
    }
  }

  log(`🚨 設定シートに項目 ${key} が見つかりません。`);
  throw new Error(`設定項目 ${key} が未設定です。設定シートを確認してください。`);
}