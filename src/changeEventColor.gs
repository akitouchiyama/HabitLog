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
  const today = new Date();
  
  // チェックボックスを含む段落のみを取得
  const paragraphs = body.getParagraphs().filter(paragraph => 
    paragraph.getText().includes('☑')
  );

  if (paragraphs.length === 0) {
    console.log('チェック済みの予定はありません');
    return;
  }

  // 正規表現を一度だけコンパイル
  const timeTitleRegex = /(\d{2}:\d{2}\s*-\s*\d{2}:\d{2})\s*(.+)/;
  
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
      const match = text.match(timeTitleRegex);
      
      if (!match) {
        throw new Error(`予定の形式が不正です: ${text}`);
      }

      const [_, timeStr, title] = match;
      
      // カレンダーのイベントを更新
      const updateResult = updateEventColor(today, timeStr, title.trim());
      
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
 * 時間文字列を解析して開始時刻と終了時刻を取得する
 * 
 * @param {string} timeStr 時間文字列（"HH:mm - HH:mm"形式）
 * @returns {Object} 開始時刻と終了時刻のオブジェクト
 */
function parseTimeString(timeStr) {
  const [startStr, endStr] = timeStr.split('-').map(t => t.trim());
  const [startHour, startMin] = startStr.split(':').map(Number);
  const [endHour, endMin] = endStr.split(':').map(Number);
  
  return {
    startHour,
    startMin,
    endHour,
    endMin
  };
}

/**
 * カレンダーのイベントの色を更新する
 * 
 * @param {Date} date 対象の日付
 * @param {string} timeStr 時間文字列（"HH:mm - HH:mm"形式）
 * @param {string} title イベントのタイトル
 * @returns {boolean} 更新が成功したかどうか
 */
function updateEventColor(date, timeStr, title) {
  try {
    // 時間範囲を解析
    const { startHour, startMin, endHour, endMin } = parseTimeString(timeStr);
    
    // 開始時刻と終了時刻を設定
    const startTime = new Date(date);
    startTime.setHours(startHour, startMin, 0, 0);
    const endTime = new Date(date);
    endTime.setHours(endHour, endMin, 0, 0);
    
    // カレンダーから該当するイベントを検索
    const calendar = CalendarApp.getDefaultCalendar();
    const events = calendar.getEvents(startTime, endTime);
    
    // タイトルが一致するイベントを探す
    const matchingEvent = events.find(event => event.getTitle() === title);
    
    if (!matchingEvent) {
      console.log(`イベントが見つかりませんでした: ${title} (${timeStr})`);
      return false;
    }

    // イベントの色を変更
    matchingEvent.setColor(CalendarApp.EventColor.SAGE);
    console.log(`イベントの色を変更しました: ${title} (${timeStr})`);
    return true;

  } catch (error) {
    console.error(`イベントの色変更でエラー: ${error.message}`, {
      date: date.toISOString(),
      timeStr,
      title
    });
    return false;
  }
}