/**
 * ToDoリストを作成するメイン関数
 */
function createToDoList() {
  const today = new Date()

  // 本日の予定を取得
  const events = getScheduleEvents(today);

  // 予定をスプレッドシートに書き込む
  writeEventsToSpreadsheet(events);
  
  // 書式設定を適用
  formatSpreadsheet(events.length);
}

/**
 * 予定をスプレッドシートに書き込む（データの書き込みのみ）
 * 
 * @param {Array} events 予定の配列
 */
function writeEventsToSpreadsheet(events) {
  // スプレッドシートを取得
  const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!sheetId) {
    throw new Error('SHEET_IDが設定されていません。PropertiesServiceで設定してください。');
  }
  const spreadsheet = SpreadsheetApp.openById(sheetId);
  const sheet = spreadsheet.getActiveSheet();

  // 既存のスプレッドシートの内容を削除
  sheet.clear();

  // 予定が存在する場合のみ処理を実行
  if (events.length > 0) {
    // シートのタイトルを追加
    sheet.getRange(1, 1).setValue("【本日の予定】");
    
    // B2セルに「予定」というヘッダーを追加
    sheet.getRange(2, 2).setValue("予定");
    // C2セルに「完了」というヘッダーを追加
    sheet.getRange(2, 3).setValue("完了");
    
    // 各予定をチェックボックス形式で追加
    events.forEach((event, index) => {
      // B3セル以降に新規予定を追加
      sheet.getRange(index + 3, 2).setValue(event.getTitle());
      // C3セル以降にチェックボックスを追加
      sheet.getRange(index + 3, 3).insertCheckboxes();
    });
  }
}

/**
 * スプレッドシートの書式設定を行う
 * 
 * @param {number} eventCount 予定の数
 */
function formatSpreadsheet(eventCount) {
  // スプレッドシートを取得
  const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!sheetId) {
    throw new Error('SHEET_IDが設定されていません。PropertiesServiceで設定してください。');
  }
  const spreadsheet = SpreadsheetApp.openById(sheetId);
  const sheet = spreadsheet.getActiveSheet();
  
  // 予定が存在する場合のみ書式設定を実行
  if (eventCount > 0) {
    // タイトルの書式設定
    const titleRange = sheet.getRange(1, 1);
    titleRange.setFontSize(16)
             .setFontWeight('bold')
             .setBackground('#4285f4')
             .setFontColor('#ffffff')
             .setHorizontalAlignment('center');
    
    // ヘッダー行の書式設定
    const headerRange = sheet.getRange(2, 2, 1, 2);
    headerRange.setBackground('#e8f0fe')
               .setFontWeight('bold')
               .setHorizontalAlignment('center')
               .setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
    
    // データ行の書式設定
    const dataRange = sheet.getRange(3, 2, eventCount, 2);
    dataRange.setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
    
    // 交互に背景色を設定（縞模様）
    for (let i = 0; i < eventCount; i++) {
      const rowRange = sheet.getRange(i + 3, 2, 1, 2);
      if (i % 2 === 0) {
        rowRange.setBackground('#f8f9fa');
      } else {
        rowRange.setBackground('#ffffff');
      }
    }
    
    // 列幅の調整
    sheet.setColumnWidth(1, 200);    // A列：タイトル（自動調整）
    sheet.setColumnWidth(2, 400); // B列：予定内容
    sheet.setColumnWidth(3, 80);  // C列：チェックボックス
    
    // 行の高さを調整
    sheet.setRowHeight(1, 40); // タイトル行
    sheet.setRowHeight(2, 30); // ヘッダー行
    for (let i = 3; i < eventCount + 3; i++) {
      sheet.setRowHeight(i, 25); // データ行
    }
  }
}

/**
 * 指定した日付の予定を取得する
 * 
 * @param {Date} date 取得する日付
 * @return {Array} 予定の配列
 */
function getScheduleEvents(date) {
  // 日付の開始時刻と終了時刻を設定
  const startTime = new Date(date);
  startTime.setHours(0, 0, 0, 0);
  const endTime = new Date(date);
  endTime.setHours(23, 59, 59, 999);

  // カレンダーから予定を取得
  const calendarId = PropertiesService.getScriptProperties().getProperty('CALENDAR_ID');
  if (!calendarId) {
    throw new Error('CALENDAR_IDが設定されていません。PropertiesServiceで設定してください。');
  }
  const calendar = CalendarApp.getCalendarById(calendarId);
  return calendar.getEvents(startTime, endTime);
}