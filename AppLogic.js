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
 * Tasksリスト再設定時、または強制リセット時に課題シートのデータをクリアする
 * (ヘッダー行とTasks ID/フラグだけでなく、課題全体をクリアし、次回全て再取得させる)
 */
function clearAssignmentSheets() {
  log('--- 課題シートの全データクリア開始 (Tasksリスト再設定のため) ---');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  [SHEET_NAME_WEBCLASS, SHEET_NAME_CLASSROOM].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) return;
    
    // ヘッダー行 (1行目) を残して、2行目以降の全データをクリア
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    // データが存在する場合のみクリア実行
    if (lastRow > 1 && lastCol > 0) {
        sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
        log(`✅ シート「${name}」の全課題データをクリアしました。`);
    }
  });
  log('--- 課題シートの全データクリア完了 ---');
}


/**
 * WebClassから課題を取得し、シートに書き込む
 */
function processWebClass() {
  log('--- WebClass課題取得開始 ---');
  const u = Settings.getSetting('userid');
  const p = Settings.getSetting('password');
  
  if (!u || !p) {
    throw new Error('WebClass認証情報が未設定です。メニューから設定してください。');
  }

  const client = new WebClassClient();
  let dashUrl;
  try {
    dashUrl = client.login(u, p);
  } catch(e) {
    log(`🚨 ログイン失敗: ${e.message}`);
    throw new Error('WebClassへのログインに失敗しました。認証情報を確認してください。');
  }

  const dashHtml = client.fetchWithSession(dashUrl);
  const courses = WebClassParser.parseDashboard(dashHtml);

  const rows = [];
  courses.forEach(c => {
    let cName = c.name.replace(/^\s*\d+\s*/, '').replace(/\s*\(.*\)\s*$/, '').trim();
    try {
      const html = client.fetchWithSession(c.url);
      const asses = WebClassParser.parseCourseContents(html);
      
      asses.forEach(a => {
        // Tasks ID(6) と フラグ(7) は空でセット
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
    if(courses) {
      courses.forEach(c => {
        const works = Classroom.Courses.CourseWork.list(c.id, { courseWorkStates: ['PUBLISHED'] }).courseWork;
        if (!works) return;
        
        works.forEach(w => {
          if (!w.dueDate) return; 

          const d = w.dueDate; 
          const t = w.dueTime || {hours:0,minutes:0};
          
          const dt = new Date(d.year, d.month-1, d.day, t.hours||0, t.minutes||0);
          const dueStr = Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
          
          // Tasks ID(6) と フラグ(7) は空でセット
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

  // TasksリストIDの有効性チェック
  try {
    Tasks.Tasklists.get(listId); 
  } catch (e) {
    if (e.message.includes('Not Found') || e.message.includes('not found')) {
      log(`🚨 TasksリストID「${listId}」が見つかりませんでした。リストが削除された可能性があります。`);
      Settings.deleteTaskListId(); 
      log('✅ 無効なTasksリストIDを削除しました。Tasks連携を再開するには、メニューの「2. Tasks・自動実行設定を完了」から再設定してください。');
      return; 
    }
    log(`🚨 Tasks APIエラーにより同期中断: ${e.message}`);
    throw new Error(`Tasks APIエラーにより同期中断: ${e.message}`);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheetDataMap = new Map(); 
  
  // 1. 各シートの最新データを読み込み、Tasks ID/Flagを復元するためのマップを作成
  [SHEET_NAME_WEBCLASS, SHEET_NAME_CLASSROOM].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) return;
    
    // シートの全課題データ（取得直後の、Tasks ID/Flagが空の可能性が高いデータ）を取得
    const data = sheet.getRange(2, 1, sheet.getLastRow()-1, HEADER.length).getValues();
    
    sheetDataMap.set(name, { rows: data, sheet: sheet, updated: false });
  });

  // 2. 既存の Tasks IDとFlag を保持するためのマップを作成（全シートを統合）
  const previousDataMap = new Map();
  [SHEET_NAME_WEBCLASS, SHEET_NAME_CLASSROOM].forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (!sheet || sheet.getLastRow() <= 1) return;
      const data = sheet.getRange(2, 1, sheet.getLastRow()-1, HEADER.length).getValues();
      data.forEach(row => {
          const link = row[5]; 
          const tasksId = row[6]; 
          const registeredFlag = row[7]; 
          if (link && (tasksId || registeredFlag)) {
              previousDataMap.set(link, [tasksId, registeredFlag]);
          }
      });
  });

  const allRows = []; 
  // 3. 最新データにTasks ID/Flagをマージし、allRowsに統合
  sheetDataMap.forEach((context, sheetName) => {
      context.rows.forEach((row, originalIndex) => {
          const link = row[5];
          if (previousDataMap.has(link)) {
              const [tasksId, registeredFlag] = previousDataMap.get(link);
              row[6] = tasksId;
              row[7] = registeredFlag;
          }
          allRows.push([...row, sheetName, originalIndex]); // 統合データに追加
      });
  });


  if (allRows.length === 0) {
    log('同期対象の課題が見つかりませんでした。');
    _cleanup(ss); 
    return;
  }
  
  // 4. 統合した全課題を、締切の遅い順にソートする (Tasksへの登録順を決定)
  allRows.sort((a, b) => {
    const dateA = parseAssignmentDate(a[4]); 
    const dateB = parseAssignmentDate(b[4]);

    const timeA = dateA ? dateA.getTime() : Infinity;
    const timeB = dateB ? dateB.getTime() : Infinity;

    return timeB - timeA; 
  });

  // 5. 締切の遅い順にTasksへの同期・登録処理を実行
  allRows.forEach(fullRow => {
    const [src, course, title, start, due, link, taskId, flag, sheetName, originalIndex] = fullRow;
    
    const sheetContext = sheetDataMap.get(sheetName);
    const originalRow = sheetContext.rows[originalIndex]; // これはマージ後のデータ

    
    // --- 課題の完了状態をTasksからシートへ同期（originalRowを操作） ---
    if (originalRow[6] && !['COMPLETED','DELETED'].includes(originalRow[7])) {
      try {
        const taskStatus = Tasks.Tasks.get(listId, originalRow[6]).status;
        if (taskStatus === 'completed') {
          originalRow[7] = 'COMPLETED'; sheetContext.updated = true;
        }
      } catch(e) { 
        if(e.message.includes('NotFound')) { 
          originalRow[7] = 'DELETED'; sheetContext.updated = true; 
          log(`Tasksから削除された課題を検出: ${title}`);
        }
      }
    }

    // --- 新規課題をTasksに登録（originalRowを操作） ---
    // Tasks IDが空（まだ登録されていない）場合にのみ登録を試みる
    if (!originalRow[6] && !['COMPLETED','DELETED','EXPIRED', 'SKIPPED_NODATE'].includes(originalRow[7])) {
      
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
      } catch(e) { 
        log(`Tasks登録失敗: ${title} - ${e.message}`); 
      }
    }
  });

  // 6. 更新されたデータをソートし、元のシートに書き戻す
  sheetDataMap.forEach((context, name) => {
    if (context.updated) {
      // 課題を期限の早い順にソート（シート表示用）
      context.rows.sort((a, b) => {
        const dateA = parseAssignmentDate(a[4]);
        const dateB = parseAssignmentDate(b[4]);

        const timeA = dateA ? dateA.getTime() : Infinity;
        const timeB = dateB ? dateB.getTime() : Infinity;

        return timeA - timeB;
      });
      
      // シートに書き戻す
      context.sheet.getRange(2, 1, context.rows.length, context.rows[0].length).setValues(context.rows);
      SpreadsheetApp.flush();
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
      const [,,,, due,, taskId, flag] = row;
      
      let dObj = parseAssignmentDate(due); 

      let shouldDelete = false;

      if (['COMPLETED','DELETED','EXPIRED', 'SKIPPED_NODATE'].includes(flag)) {
        
        if (flag === 'SKIPPED_NODATE' || !dObj) {
            shouldDelete = true;
        } else {
            if ((now - dObj.getTime()) > thresh) shouldDelete = true; 
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
