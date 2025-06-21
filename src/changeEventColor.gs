/**
 * チェック済み予定の色を変更する関数
 */
function updateCheckedEventColors() {
      try {
    // ドキュメントを取得
    const documentId = PropertiesService.getScriptProperties().getProperty('DOCUMENT_ID');
    if (!documentId) {
      throw new Error('DOCUMENT_IDが設定されていません。PropertiesServiceで設定してください。');
    }
    const doc = DocumentApp.openById(documentId);
    if (!doc) {
      console.log('ドキュメントが見つかりません');
      return;
    }

    // 完了済み予定の色変更を処理
    processCompletedEvents(doc);
  } catch (error) {
    console.error(`エラーが発生しました: ${error.message}`);
  }
}

/**
 * 完了済み予定の色変更を処理する
 * 
 * @param {GoogleAppsScript.Document.Document} doc ドキュメント
 */
function processCompletedEvents(doc) {
  const body = doc.getBody();
  const today = new Date();
  
  // OK/NG付きのイベントを含む段落のみを取得
  const paragraphs = body.getParagraphs().filter(paragraph => {
    const text = paragraph.getText();
    return text.includes('OK ') || text.includes('NG ');
  });

  if (paragraphs.length === 0) {
    console.log('OK/NG付きの予定はありません');
    return;
  }

  // 正規表現を一度だけコンパイル
  const statusEventRegex = /(OK|NG)\s+(.+)/;
  
  // 処理結果を記録
  const results = {
    success: 0,
    failed: 0,
    notFound: 0,  // イベントが見つからなかった数
    errors: []
  };

  paragraphs.forEach(paragraph => {
    try {
      const text = paragraph.getText();
      const match = text.match(statusEventRegex);
      
      if (!match) {
        throw new Error(`予定の形式が不正です: ${text}`);
      }

      const [_, status, eventTitle] = match;
      
      // ステータスに応じて色を決定
      const color = status === 'OK' ? CalendarApp.EventColor.GREEN : CalendarApp.EventColor.RED;
      
      // カレンダーのイベントを更新
      const updateResult = updateEventByTitle(today, eventTitle.trim(), color);
      
      if (updateResult === true) {
        results.success++;
      } else if (updateResult === false) {
        results.notFound++;
        results.errors.push({
          text: text,
          error: 'カレンダーにイベントが見つかりませんでした'
        });
      }
      
    } catch (error) {
      results.failed++;
      results.errors.push({
        text: paragraph.getText(),
        error: error.message
      });
      console.error(`段落の処理でエラー: ${error.message}`);
    }
  });

  // 処理結果をログに出力
  console.log(`処理完了: 成功=${results.success}, イベント未発見=${results.notFound}, 失敗=${results.failed}`);
  if (results.errors.length > 0) {
    console.log('エラー詳細:', results.errors);
  }
}

/**
 * イベントタイトルでカレンダーのイベントの色を更新する
 * 
 * @param {Date} date 対象の日付
 * @param {string} title イベントのタイトル
 * @param {GoogleAppsScript.Calendar.EventColor} color 設定する色
 * @returns {boolean} 更新が成功したかどうか
 */
function updateEventByTitle(date, title, color) {
  try {
    // 対象日の開始時刻と終了時刻を設定（一日全体）
    const startTime = new Date(date);
    startTime.setHours(0, 0, 0, 0);
    const endTime = new Date(date);
    endTime.setHours(23, 59, 59, 999);
    
    // カレンダーから該当するイベントを検索
    const calendarId = PropertiesService.getScriptProperties().getProperty('CALENDAR_ID');
    if (!calendarId) {
      throw new Error('CALENDAR_IDが設定されていません。PropertiesServiceで設定してください。');
    }
    const calendar = CalendarApp.getCalendarById(calendarId);
    const events = calendar.getEvents(startTime, endTime);
    
    // タイトルが一致するイベントを探す
    const matchingEvent = events.find(event => event.getTitle() === title);
    
    if (!matchingEvent) {
      console.log(`イベントが見つかりませんでした: ${title}`);
      return false;
    }

    // イベントの色を変更
    matchingEvent.setColor(color);
    const colorName = color === CalendarApp.EventColor.SAGE ? '緑色' : '赤色';
    console.log(`イベントの色を${colorName}に変更しました: ${title}`);
    return true;

  } catch (error) {
    console.error(`イベントの色変更でエラー: ${error.message}`, {
      date: date.toISOString(),
      title
    });
    return false;
  }
}

