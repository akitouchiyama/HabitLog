/**
 * ToDoリストを作成するメイン関数
 */
function createToDoList() {
  const today = new Date()

  // 本日の予定を取得
  const events = getScheduleEvents(today);
  console.log("本日の予定:", events);

  // 予定をドキュメントに書き込む
  writeEventsToDocument(events);
}

/**
 * 予定をドキュメントに書き込む
 * 
 * @param {Array} events 予定の配列
 */
function writeEventsToDocument(events) {
  // 現在のドキュメントを取得
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  // 予定が存在する場合のみ処理を実行
  if (events.length > 0) {
    // 既存の予定を取得
    const existingEvents = getExistingEvents(body);
    
    // 予定セクションのヘッダーを追加
    body.appendParagraph("【本日の予定】").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    
    // 各予定をチェックボックス形式で追加
    events.forEach(event => {
      const title = event.getTitle();
      const startTime = event.getStartTime();
      const endTime = event.getEndTime();
      
      // 時間をフォーマット
      const timeStr = Utilities.formatDate(startTime, 'JST', 'HH:mm') + 
                     ' - ' + 
                     Utilities.formatDate(endTime, 'JST', 'HH:mm');
      
      // 予定の文字列を作成
      const eventText = `${timeStr} ${title}`;
      
      // 既存の予定に存在する場合、チェックボックスをリセット
      if (existingEvents.includes(eventText)) {
        resetExistingCheckbox(body, eventText);
      } else {
        // 新規予定を追加
        const paragraph = body.appendParagraph();
        const checkbox = paragraph.addCheckBox();
        checkbox.setChecked(false);
        paragraph.appendText(` ${eventText}`);
      }
    });
    
    // 空行を追加
    body.appendParagraph("");
  }
}

/**
 * 既存の予定のチェックボックスをリセットする
 * 
 * @param {GoogleAppsScript.Document.Body} body ドキュメントの本文
 * @param {string} eventText 予定の文字列
 */
function resetExistingCheckbox(body, eventText) {
  const paragraphs = body.getParagraphs();
  
  paragraphs.forEach(paragraph => {
    const text = paragraph.getText();
    // チェックボックスが☑の段落のみを処理
    if (text.includes('☑') && text.includes(eventText)) {
      // チェックボックスを☐に変更
      const newText = text.replace('☑', '☐');
      paragraph.setText(newText);
    }
  });
}

/**
 * ドキュメント内の既存の予定を取得する
 * 
 * @param {GoogleAppsScript.Document.Body} body ドキュメントの本文
 * @return {Array} 既存の予定の配列
 */
function getExistingEvents(body) {
  const existingEvents = [];
  const paragraphs = body.getParagraphs();
  
  paragraphs.forEach(paragraph => {
    const text = paragraph.getText();
    // チェックボックスを含む段落のみを処理
    if (text.includes('☐') || text.includes('☑')) {
      // チェックボックスと時間の部分を除去して予定のタイトルのみを取得
      const eventText = text.replace(/[☐☑]?\s*\d{2}:\d{2}\s*-\s*\d{2}:\d{2}\s*/, '').trim();
      existingEvents.push(eventText);
    }
  });
  
  return existingEvents;
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
  const calendar = CalendarApp.getDefaultCalendar();
  const events = calendar.getEvents(startTime, endTime);

  // 「:schedule」を含む予定のみをフィルタリング
  return events.filter(event => {
    const title = event.getTitle();
    return title.includes(':schedule');
  });
}
