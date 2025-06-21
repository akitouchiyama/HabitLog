/**
 * ToDoリストを作成するメイン関数
 */
function createToDoList() {
  const today = new Date()

  // 本日の予定を取得
  const events = getScheduleEvents(today);

  // 予定をドキュメントに書き込む
  writeEventsToDocument(events);
}

/**
 * 予定をドキュメントに書き込む
 * 
 * @param {Array} events 予定の配列
 */
function writeEventsToDocument(events) {
  // ドキュメントを取得
  const documentId = PropertiesService.getScriptProperties().getProperty('DOCUMENT_ID');
  if (!documentId) {
    throw new Error('DOCUMENT_IDが設定されていません。PropertiesServiceで設定してください。');
  }
  const doc = DocumentApp.openById(documentId);
  const body = doc.getBody();

  // 既存のドキュメントの内容を削除
  body.clear();

  // 予定が存在する場合のみ処理を実行
  if (events.length > 0) {
    // 予定セクションのヘッダーを追加
    body.appendParagraph("【本日の予定】").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    
    // 各予定をチェックボックス形式で追加
    events.forEach(event => {
      // 新規予定を追加
      body.appendListItem(event.getTitle());
    });
    
    // 空行を追加
    body.appendParagraph("");
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