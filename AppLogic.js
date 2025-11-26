/**
 * AppLogic.gs
 * 課題データの取得、変換、Tasks連携のロジックを管理するコアファイル。
 * 依存: 
 * - Config.gs (定数)
 * - Utils.gs (Settings, SheetUtils, log, parseAssignmentDate)
 * - WebClassClient.gs (WebClassClientクラス)
 * - Parser.gs (WebClassParser)
 * - Tasks API サービス, Classroom API サービス
 */

/**
 * WebClassから課題を取得し、シートに書き込む
 */
function processWebClass() {
  log('--- WebClass課題取得開始 ---');
  const u = Settings.getSetting('userid');
  const p = Settings.getSetting('password');

  // 認証情報未設定の場合は中断
  if (!u || !p) {
    throw new Error('WebClass認証情報が未設定です。メニューから設定してください。');
  }

  const client = new WebClassClient();
  let dashUrl;
  try {
    dashUrl = client.login(u, p);
  } catch (e) {
    log(`🚨 ログイン失敗: ${e.message}`);
    throw new Error('WebClassへのログインに失敗しました。認証情報を確認してください。');
  }

  // ダッシュボードからコース一覧を取得
  const dashHtml = client.fetchWithSession(dashUrl);
  const courses = WebClassParser.parseDashboard(dashHtml);

  const rows = [];
  courses.forEach(c => {
    let cName = c.name.replace(/^\s*\d+\s*/, '').replace(/\s*\(.*\)\s*$/, '').trim();
    try {
      const html = client.fetchWithSession(c.url);
      const asses = WebClassParser.parseCourseContents(html);

      asses.forEach(a => {
        // [ソース, 授業名, 課題タイトル, 開始日時, 終了日時, 課題リンク, Tasks ID, 登録済みフラグ]
        rows.push(['WebClass', cName, a.title, a.start, a.end, a.shareLink, '', '']);
      });
    } catch (e) {
      log(`⚠️ ${cName} の課題取得中にエラー: ${e.message}`);
    }
    Utilities.sleep(500);
  });

  SheetUtils.writeToSheet(SHEET_NAME_WEBCLASS, rows);
  log('--- WebClass課題取得完了 ---');
}

/**
 * Google Classroomから課題を取得し、シートに書き込む
 */
function processClassroom() {
  log('--- Classroom課題取得開始 ---');
  try {
    const courses = Classroom.Courses.list({ courseStates: ['ACTIVE'] }).courses;
    const rows = [];
    if (courses) {
      courses.forEach(c => {
        const works = Classroom.Courses.CourseWork.list(c.id, { courseWorkStates: ['PUBLISHED'] }).courseWork;
        if (!works) return;

        works.forEach(w => {
          if (!w.dueDate) return;

          const d = w.dueDate;
          const t = w.dueTime || { hours: 0, minutes: 0 };

          const dt = new Date(d.year, d.month - 1, d.day, t.hours || 0, t.minutes || 0);
          const dueStr = Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');

          rows.push(['Classroom', c.name, w.title, '', dueStr, w.alternateLink, '', '']);
        });
      });
    }
    SheetUtils.writeToSheet(SHEET_NAME_CLASSROOM, rows);
    log('--- Classroom課題取得完了 ---');
  } catch (e) {
    log(`🚨 Classroom取得エラー: ${e.message}`);
  }
}

/**
 * スプレッドシートとTasksの同期処理
 */
function processTasksSync() {
  const listId = Settings.getTaskListId();
  if (!listId) {
    log('⚠️ TasksリストIDが未設定のため、同期・登録処理をスキップしました。');
    return;
  }

  log('--- Tasks同期処理開始 ---');
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const allRows = [];
  const sheetDataMap = new Map();

  // 1. WebClassとClassroomの全データを読み込み、統合する
  [SHEET_NAME_WEBCLASS, SHEET_NAME_CLASSROOM].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) return;

    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADER.length);
    const data = range.getValues();

    sheetDataMap.set(name, { rows: data, range: range, updated: false });

    data.forEach((row, originalIndex) => {
      allRows.push([...row, name, originalIndex]);
    });
  });

  if (allRows.length === 0) {
    log('同期対象の課題が見つかりませんでした。');
    _cleanup(ss);
    return;
  }

  // 2. 統合した全課題を、締切の遅い順にソートする (Tasksへの登録順を決定)
  allRows.sort((a, b) => {
    const dateA = parseAssignmentDate(a[4]);
    const dateB = parseAssignmentDate(b[4]);

    const timeA = dateA ? dateA.getTime() : Infinity;
    const timeB = dateB ? dateB.getTime() : Infinity;

    // ★重要修正点: timeB - timeA にすることで、締切が遅い順（降順）になる
    return timeB - timeA;
  });

  // 3. 締切の遅い順にTasksへの同期・登録処理を実行
  allRows.forEach(fullRow => {
    const [src, course, title, start, due, link, taskId, flag, sheetName, originalIndex] = fullRow;

    const sheetContext = sheetDataMap.get(sheetName);
    const originalRow = sheetContext.rows[originalIndex];

    // --- 課題の完了状態をTasksからシートへ同期（originalRowを操作） ---
    if (originalRow[6] && !['COMPLETED', 'DELETED'].includes(originalRow[7])) {
      try {
        const taskStatus = Tasks.Tasks.get(listId, originalRow[6]).status;
        if (taskStatus === 'completed') {
          originalRow[7] = 'COMPLETED'; sheetContext.updated = true;
        }
      } catch (e) {
        if (e.message.includes('NotFound')) {
          originalRow[7] = 'DELETED'; sheetContext.updated = true;
          log(`Tasksから削除された課題を検出: ${title}`);
        }
      }
    }

    // --- 新規課題をTasksに登録（originalRowを操作） ---
    if (!originalRow[6] && !['COMPLETED', 'DELETED', 'EXPIRED', 'SKIPPED_NODATE'].includes(originalRow[7])) {

      let dueObj = parseAssignmentDate(due);

      if (!dueObj) {
        originalRow[7] = 'SKIPPED_NODATE'; sheetContext.updated = true;
        return;
      }

      // 既に期限が過ぎているかチェック (1日余裕)
      if (dueObj.getTime() < new Date().getTime() - (24 * 3600 * 1000)) {
        originalRow[7] = 'EXPIRED'; sheetContext.updated = true;
        log(`期限切れの課題を検出: ${title}`);
        return;
      }

      try {
        const diff = (dueObj.getTime() - new Date().getTime()) / 86400000;
        const urgent = diff <= 3;
        const dueDisp = Utilities.formatDate(dueObj, Session.getScriptTimeZone(), 'MM/dd(E) HH:mm');

        let taskDue = new Date(dueObj);

        const task = {
          title: `${urgent ? '🔥 ' : ''}[${course}] ${title} (${dueDisp}まで)`,
          due: taskDue.toISOString(),
          notes: `リンク:\n${link}\n\n期限: ${dueDisp}\nソース: ${src}`
        };

        const t = Tasks.Tasks.insert(task, listId);

        originalRow[6] = t.id;
        originalRow[7] = 'REGISTERED';
        sheetContext.updated = true;
        log(`Tasks登録: ${task.title}`);
      } catch (e) {
        log(`Tasks登録失敗: ${title} - ${e.message}`);
      }
    }
  });

  // 4. 更新されたデータを元のシートに書き戻す
  sheetDataMap.forEach((context, name) => {
    if (context.updated) {
      SheetUtils.writeToSheet(name, context.rows);
    }
  });

  _cleanup(ss);
  log('--- Tasks同期処理完了 ---');
}

/**
 * 期限切れ、完了済み、削除済みタスクをシートから削除（整理）
 */
function _cleanup(ss) {
  const days = Number(Settings.getSetting('cleanupDays') || 30);
  const thresh = days * 86400000;
  const now = new Date().getTime();

  log(`--- シートクリーンアップ開始 (猶予期間: ${days}日) ---`);

  [SHEET_NAME_WEBCLASS, SHEET_NAME_CLASSROOM].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) return;

    const rows = sheet.getDataRange().getValues();

    for (let i = rows.length - 1; i >= 1; i--) {
      const row = rows[i];
      const [, , , , due, , taskId, flag] = row;

      let dObj = parseAssignmentDate(due);

      let shouldDelete = false;

      if (['COMPLETED', 'DELETED', 'EXPIRED', 'SKIPPED_NODATE'].includes(flag)) {

        if (flag === 'SKIPPED_NODATE' || !dObj) {
          shouldDelete = true;
        } else {
          if (now - dObj.getTime() > thresh) shouldDelete = true;
        }
      }

      if (!taskId && dObj && (now - dObj.getTime()) > thresh) {
        shouldDelete = true;
      }

      if (shouldDelete) {
        sheet.deleteRow(i + 1);
      }
    }
  });
  log('--- シートクリーンアップ完了 ---');
}