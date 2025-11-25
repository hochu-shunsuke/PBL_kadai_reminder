/**
 * AppLogic.gs
 * WebClass, Classroom, Tasksの各処理ロジック (旧 Coge.gs)
 */

// --- 1. WebClass取得 ---
function processWebClass() {
  log('--- WebClass処理開始 ---');
  const creds = Props.getCredentials();
  if (!creds) throw new Error('WebClass認証情報が未設定です。');

  const client = new WebClassClient();
  const dashboardUrl = client.login(creds.userid, creds.password);

  const dashboardHtml = client.fetchWithSession(dashboardUrl);
  const courses = WebClassParser.parseDashboard(dashboardHtml);

  const allRows = [];
  courses.forEach((course, i) => {
    let courseName = course.name.replace(/^\s*\d+\s*/, '').replace(/\s*\(.*\)\s*$/, '').trim();

    try {
      const html = client.fetchWithSession(course.url);
      const assignments = WebClassParser.parseCourseContents(html);

      assignments.forEach(a => {
        allRows.push([
          'WebClass', courseName, a.title, a.start, a.end, a.shareLink, '', ''
        ]);
      });
    } catch (e) {
      log(`🚨 ${courseName} の取得失敗: ${e.message}`);
    }
    Utilities.sleep(500);
  });

  SheetUtils.writeToSheet(SHEET_NAME_WEBCLASS, allRows);
  log('--- WebClass処理完了 ---');
}

// --- 2. Classroom取得 ---
function processClassroom() {
  log('--- Classroom処理開始 ---');
  try {
    const courses = Classroom.Courses.list({ courseStates: ['ACTIVE'] }).courses;

    const allRows = [];
    courses.forEach(course => {
      const works = Classroom.Courses.CourseWork.list(course.id, { courseWorkStates: ['PUBLISHED'] }).courseWork;
      if (!works) return;

      works.forEach(work => {
        if (!work.dueDate) return;

        const d = work.dueDate;
        const t = work.dueTime || { hours: 0, minutes: 0 };
        const dateObj = new Date(d.year, d.month - 1, d.day, t.hours || 0, t.minutes || 0);
        const dueStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');

        allRows.push(['Classroom', course.name, work.title, '', dueStr, work.alternateLink, '', '']);
      });
    });

    SheetUtils.writeToSheet(SHEET_NAME_CLASSROOM, allRows);
  } catch (e) {
    log(`🚨 Classroom取得エラー: ${e.message}`);
  }
  log('--- Classroom処理完了 ---');
}

// --- 3. Tasks同期・登録 ---
function processTasksSync() {
  const taskListId = getTaskListIdProperty();
  if (!taskListId) {
    log('TasksリストIDが未設定です。同期をスキップします。');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = [SHEET_NAME_WEBCLASS, SHEET_NAME_CLASSROOM];

  sheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) return;

    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADER.length);
    const data = range.getValues();
    let isUpdated = false;

    data.forEach((row, i) => {
      const [src, course, title, start, due, link, taskId, flag] = row;

      // A. 完了同期 
      if (taskId && flag !== 'COMPLETED' && flag !== 'DELETED') {
        try {
          const t = Tasks.Tasks.get(taskListId, taskId);
          if (t.status === 'completed') {
            data[i][7] = 'COMPLETED';
            isUpdated = true;
          }
        } catch (e) {
          if (e.message.includes('NotFound')) {
            data[i][7] = 'DELETED'; // Tasks側で削除された
            isUpdated = true;
          }
        }
      }

      // B. 新規登録 [改善案 5, 6 の実装]
      if (!taskId && !['COMPLETED', 'DELETED', 'EXPIRED'].includes(flag)) {

        let dueDateObj = null;
        const rawDue = String(due).trim();

        if (rawDue) {
          try {
            // 様々な日付形式に対応してパースを試みる
            dueDateObj = new Date(rawDue.replace(/(\d{4})[\/年](\d{1,2})[\/月](\d{1,2})[\日]?/g, '$1/$2/$3'));
            if (isNaN(dueDateObj.getTime())) dueDateObj = null;
          } catch (e) { dueDateObj = null; }
        }

        // 期限切れチェック
        if (dueDateObj && dueDateObj.getTime() < new Date().getTime()) {
          data[i][7] = 'EXPIRED';
          isUpdated = true;
          return;
        }

        try {
          const newTask = {};
          let dueDisplay = '期限なし';

          if (dueDateObj) {
            // --- [改善案 5] Tasksタイトル形式の最適化 ---
            const timeUntilDue = (dueDateObj.getTime() - new Date().getTime()) / (1000 * 3600 * 24);
            const isUrgent = timeUntilDue <= 3 && timeUntilDue >= 0;

            dueDisplay = Utilities.formatDate(dueDateObj, Session.getScriptTimeZone(), 'MM/dd(E) HH:mm');
            newTask.title = `${isUrgent ? '🔥 ' : ''}[${course}] ${title} (${dueDisplay}まで)`;

            // --- [改善案 6] 期限設定精度の向上 (時刻がない場合は23:59に設定) ---
            let taskDueDate = new Date(dueDateObj.getTime());
            // 時刻情報が含まれていないかチェック
            if (!rawDue.match(/(\d{1,2}:\d{2})/) && !rawDue.match(/(\d{1,2}時\d{2}分)/)) {
              taskDueDate.setHours(23, 59, 0, 0); // 23:59:00に設定
            }
            newTask.due = taskDueDate.toISOString(); // Tasks APIはRFC3339形式（UTC）を要求

          } else {
            newTask.title = `[${course}] ${title}`;
          }

          newTask.notes = `リンク:\n${link}\n\n期限: ${dueDisplay}\nソース: ${src}`;

          const created = Tasks.Tasks.insert(newTask, taskListId);
          data[i][6] = created.id;
          data[i][7] = 'REGISTERED';
          isUpdated = true;
          log(`Tasks登録: ${newTask.title}`);
        } catch (e) {
          log(`🚨 Tasks登録失敗: ${title} - ${e.message}`);
        }
      }
    });

    if (isUpdated) {
      range.setValues(data);
    }
  });

  _cleanupOldRows(ss, sheets);
}

/**
 * 課題の削除ロジック [改善案 9. 削除閾値の適用]
 */
function _cleanupOldRows(ss, targetSheetNames) {
  const today = new Date().getTime();

  let cleanupDays = 30;
  try {
    cleanupDays = getSetting('CLEANUP_DAYS'); // AppConfig.gs から読み込み
  } catch (e) {
    log(`⚠️ 設定シートからCLEANUP_DAYSを取得できませんでした。デフォルトの${cleanupDays}日を使用します。`);
  }
  const deleteThresholdMs = cleanupDays * 24 * 60 * 60 * 1000;

  targetSheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) return;

    const rows = sheet.getDataRange().getValues();

    // 後ろからループして削除
    for (let i = rows.length - 1; i >= 1; i--) {
      const [src, course, title, start, due, link, taskId, flag] = rows[i];

      let shouldDelete = false;
      const rawDue = String(due).trim();

      let dueDateObj = null;
      if (rawDue) {
        try {
          dueDateObj = new Date(rawDue.replace(/(\d{4})[\/年](\d{1,2})[\/月](\d{1,2})[\日]?/g, '$1/$2/$3'));
          if (isNaN(dueDateObj.getTime())) dueDateObj = null;
        } catch (e) { dueDateObj = null; }
      }

      // 1. 完了・削除・期限切れ済みの場合
      if (['COMPLETED', 'DELETED', 'EXPIRED'].includes(flag)) {
        // 期限があり、かつ期限切れから閾値日数以上経過
        if (dueDateObj && (today - dueDateObj.getTime()) > deleteThresholdMs) {
          shouldDelete = true;
        }
        // 期限はないがTasksから削除されたものは即座に削除（ゴミデータ回避）
        else if (flag === 'DELETED' && !dueDateObj) {
          shouldDelete = true;
        }
      }

      // 2. 未連携で、期限からCLEANUP_DAYS以上経過（古いゴミデータ）
      if (!taskId && dueDateObj && (today - dueDateObj.getTime()) > deleteThresholdMs) {
        shouldDelete = true;
      }

      if (shouldDelete) {
        sheet.deleteRow(i + 1);
      }
    }
  });
}

// TasksリストIDの検索・作成ヘルパー
function getTaskListId(taskListName) {
  const lists = Tasks.Tasklists.list().items;
  let targetId = null;

  for (const list of lists) {
    if (list.title === taskListName) {
      targetId = list.id;
      break;
    }
  }

  if (!targetId) {
    const newList = Tasks.Tasklists.insert({ title: taskListName });
    targetId = newList.id;
  }
  return targetId;
}