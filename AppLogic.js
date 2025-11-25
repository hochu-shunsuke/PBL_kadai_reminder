/**
 * AppLogic.gs
 * WebClass, Classroom, Tasksの各処理ロジック
 */

/**
 * 1. WebClassの課題を取得してシートに書き込む
 */
function processWebClass() {
  log('--- WebClass処理開始 ---');
  const creds = Props.getCredentials();
  if (!creds) throw new Error('WebClass認証情報が未設定です。');

  const client = new WebClassClient();
  const dashboardUrl = client.login(creds.userid, creds.password);
  
  // ダッシュボード取得
  const dashboardHtml = client.fetchWithSession(dashboardUrl);
  const courses = WebClassParser.parseDashboard(dashboardHtml);
  log(`取得したコース数: ${courses.length}`);

  const allRows = [];
  courses.forEach((course, i) => {
    // 科目名整形
    let courseName = course.name.replace(/^\s*\d+\s*/, '').replace(/\s*\(.*\)\s*$/, '').trim();
    log(`[${i+1}/${courses.length}] ${courseName} をスキャン中...`);

    try {
      const html = client.fetchWithSession(course.url);
      const assignments = WebClassParser.parseCourseContents(html);
      
      assignments.forEach(a => {
        allRows.push([
          'WebClass',
          courseName,
          a.title,
          a.start,
          a.end,
          a.shareLink,
          '', '' // Tasks用プレースホルダ
        ]);
      });
    } catch (e) {
      log(`🚨 ${courseName} の取得失敗: ${e.message}`);
    }
    Utilities.sleep(500); // サーバー負荷軽減
  });

  SheetUtils.writeToSheet(SHEET_NAME_WEBCLASS, allRows);
  log('--- WebClass処理完了 ---');
}

/**
 * 2. Classroomの課題を取得してシートに書き込む
 */
function processClassroom() {
  log('--- Classroom処理開始 ---');
  try {
    const courses = Classroom.Courses.list({ courseStates: ['ACTIVE'] }).courses;
    if (!courses || courses.length === 0) {
      log('アクティブなコースがありません。');
      return;
    }

    const allRows = [];
    courses.forEach(course => {
      const works = Classroom.Courses.CourseWork.list(course.id, { courseWorkStates: ['PUBLISHED'] }).courseWork;
      if (!works) return;

      works.forEach(work => {
        if (!work.dueDate) return;
        
        // 日付整形
        const d = work.dueDate;
        const t = work.dueTime || { hours: 0, minutes: 0 };
        const dateObj = new Date(d.year, d.month - 1, d.day, t.hours || 0, t.minutes || 0);
        const dueStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');

        allRows.push([
          'Classroom',
          course.name,
          work.title,
          '', // 開始日なし
          dueStr,
          work.alternateLink,
          '', ''
        ]);
      });
    });

    SheetUtils.writeToSheet(SHEET_NAME_CLASSROOM, allRows);
  } catch (e) {
    log(`🚨 Classroom取得エラー: ${e.message}`);
  }
  log('--- Classroom処理完了 ---');
}

/**
 * 3. Tasks同期と登録（メイン処理）
 */
function processTasksSync() {
  const taskListId = Props.getTaskListId();
  if (!taskListId) {
    log('TasksリストIDが設定されていません。同期をスキップします。');
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = [SHEET_NAME_WEBCLASS, SHEET_NAME_CLASSROOM];

  sheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) return;

    // 範囲取得: ヘッダー除くデータ部分
    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADER.length);
    const data = range.getValues();
    let isUpdated = false;

    data.forEach((row, i) => {
      // Row Map: [Source, Course, Title, Start, Due, Link, TaskID, Flag]
      // Index:   0       1       2      3      4    5     6       7
      const [src, course, title, start, due, link, taskId, flag] = row;

      // A. 完了同期: Tasks側で完了していたらシートもCOMPLETEDに
      if (taskId && flag !== 'COMPLETED' && flag !== 'DELETED') {
        try {
          const t = Tasks.Tasks.get(taskListId, taskId);
          if (t.status === 'completed') {
            data[i][7] = 'COMPLETED';
            isUpdated = true;
          }
        } catch (e) {
          if (e.message.includes('NotFound')) {
            data[i][7] = 'DELETED';
            isUpdated = true;
          }
        }
      }

      // B. 新規登録: IDがなく、有効期限内ならTasksへ登録
      if (!taskId && !['COMPLETED', 'DELETED', 'EXPIRED'].includes(flag)) {
        // 日付パース
        let dueDateObj = null;
        if (due instanceof Date) dueDateObj = due;
        else if (typeof due === 'string' && due.trim()) {
          dueDateObj = new Date(due.replace(/\//g, '-'));
        }

        // 期限切れチェック
        if (dueDateObj && dueDateObj.getTime() < new Date().getTime()) {
          data[i][7] = 'EXPIRED';
          isUpdated = true;
          return;
        }

        // Tasks登録
        try {
          const newTask = {
            title: `[${course}] ${title}`,
            notes: `リンク:\n${link}\n\n期限: ${due}\nソース: ${src}`,
          };
          if (dueDateObj) {
            // Tasks APIのdueはRFC3339 timestamp (日付のみ使用する場合が多い)
            newTask.due = new Date(Date.UTC(dueDateObj.getFullYear(), dueDateObj.getMonth(), dueDateObj.getDate())).toISOString();
          }

          const created = Tasks.Tasks.insert(newTask, taskListId);
          data[i][6] = created.id;
          data[i][7] = 'REGISTERED';
          isUpdated = true;
          log(`Tasks登録: ${newTask.title}`);
        } catch (e) {
          log(`Tasks登録失敗: ${title} - ${e.message}`);
        }
      }
    });

    if (isUpdated) {
      range.setValues(data);
    }
  });
  
  // クリーンアップ（古い行の削除）
  _cleanupOldRows(ss, sheets);
}

/**
 * 古い行を削除する内部関数
 */
function _cleanupOldRows(ss, targetSheetNames) {
  const today = new Date().getTime();
  
  targetSheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) return;
    
    const rows = sheet.getDataRange().getValues(); // 全データ
    // 後ろからループして削除
    for (let i = rows.length - 1; i >= 1; i--) {
      const [src, course, title, start, due, link, taskId, flag] = rows[i];
      
      let shouldDelete = false;
      // 1. 完了・削除・期限切れ済み
      if (['COMPLETED', 'DELETED', 'EXPIRED'].includes(flag)) shouldDelete = true;
      
      // 2. 未連携だが期限から7日以上経過
      if (!taskId && due) {
        const d = new Date(due.replace(/\//g, '-'));
        if ((today - d.getTime()) > 7 * 24 * 60 * 60 * 1000) shouldDelete = true;
      }

      if (shouldDelete) {
        sheet.deleteRow(i + 1);
      }
    }
  });
}