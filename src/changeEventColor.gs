/**
 * ドキュメントの編集をトリガーに実行される関数
 * 
 * @param {Object} e イベントオブジェクト
 */
function onEdit(e) {
  try {
    // 編集されたドキュメントを取得
    const doc = DocumentApp.getActiveDocument();
    if (!doc) {
      console.log('ドキュメントが見つかりません');
      return;
    }

    // チェックボックスの変更を処理
    processCheckboxChanges(doc);
  } catch (error) {
    console.error(`エラーが発生しました: ${error.message}`);
  }
}

/**
 * チェックボックスの変更を処理する
 * 
 * @param {GoogleAppsScript.Document.Document} doc ドキュメント
 */
function processCheckboxChanges(doc) {
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();
  const today = new Date();
  
  // 各段落を処理
  paragraphs.forEach(paragraph => {
    try {
      const text = paragraph.getText();
      
      // チェックボックスが含まれる段落のみを処理
      if (text.includes('☑')) {
        // 時間と予定名を抽出
        const match = text.match(/(\d{2}:\d{2}\s*-\s*\d{2}:\d{2})\s*(.+)/);
        if (match) {
          const timeStr = match[1];
          const title = match[2].trim();
          
          // カレンダーのイベントを更新
          updateEventColor(today, timeStr, title);
        }
      }
    } catch (paragraphError) {
      console.error(`段落の処理でエラー: ${paragraphError.message}`);
    }
  });
}

/**
 * カレンダーのイベントの色を更新する
 * 
 * @param {Date} date 対象の日付
 * @param {string} timeStr 時間文字列（"HH:mm - HH:mm"形式）
 * @param {string} title イベントのタイトル
 */
function updateEventColor(date, timeStr, title) {
  try {
    // 時間範囲を解析
    const [startStr, endStr] = timeStr.split('-').map(t => t.trim());
    
    // 開始時刻と終了時刻を設定
    const [startHour, startMin] = startStr.split(':').map(Number);
    const [endHour, endMin] = endStr.split(':').map(Number);
    
    const startTime = new Date(date);
    startTime.setHours(startHour, startMin, 0, 0);
    const endTime = new Date(date);
    endTime.setHours(endHour, endMin, 0, 0);
    
    // カレンダーから該当するイベントを検索
    const calendar = CalendarApp.getDefaultCalendar();
    const events = calendar.getEvents(startTime, endTime);
    
    // タイトルが一致するイベントの色を変更
    events.forEach(event => {
      if (event.getTitle() === title) {
        // セージ色（SAGE）に設定
        event.setColor(CalendarApp.EventColor.SAGE);
        console.log(`イベントの色を変更しました: ${title}`);
      }
    });
  } catch (error) {
    console.error(`イベントの色変更でエラー: ${error.message}`);
  }
}