/**
 * チェック済み予定の色を変更する関数
 */
function updateCheckedEventColors() {
  try {
    // スプレッドシートを取得
    const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    if (!sheetId) {
      throw new Error('SHEET_IDが設定されていません。PropertiesServiceで設定してください。');
    }
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    if (!spreadsheet) {
      console.log('スプレッドシートが見つかりません');
      return;
    }
    const sheet = spreadsheet.getActiveSheet();

    // 完了済み予定の色変更を処理
    processCompletedEvents(sheet);
  } catch (error) {
    console.error(`エラーが発生しました: ${error.message}`);
  }
}

/**
 * 完了済み予定の色変更を処理する
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet スプレッドシート
 */
function processCompletedEvents(sheet) {
  const today = new Date();
  
  // B3列以降の予定データの範囲を取得（最大20行まで確認）
  const maxRows = 20;
  const eventRange = sheet.getRange(3, 2, maxRows, 2); // B3:C(3+maxRows)
  const values = eventRange.getValues();
  
  // 実際にデータが入っている行をフィルタリング（元のインデックスを保持）
  const dataRows = values.map((row, originalIndex) => ({ row, originalIndex }))
    .filter(item => {
      const eventTitle = item.row[0]; // B列の値（予定）
      return eventTitle && eventTitle.toString().trim() !== '';
    });
  
  if (dataRows.length === 0) {
    console.log('予定データがありません');
    return;
  }

  // 処理結果を記録
  const results = {
    success: 0,
    failed: 0,
    notFound: 0,  // イベントが見つからなかった数
    unchecked: 0, // チェックされていない数
    errors: []
  };

  dataRows.forEach((item, index) => {
    try {
      const row = item.row;
      const originalIndex = item.originalIndex;
      const eventTitle = row[0]; // B列の値（予定）
      const isChecked = row[1];  // C列の値（チェックボックス）
      
      // チェックボックスがTRUEでない場合はスキップ
      if (isChecked !== true) {
        results.unchecked++;
        return;
      }
      
      // チェックされた予定の色を緑色に変更
      const color = CalendarApp.EventColor.GREEN;
      
      // カレンダーのイベントを更新
      const updateResult = updateEventByTitle(today, eventTitle.toString().trim(), color);
      
      if (updateResult === true) {
        results.success++;
      } else if (updateResult === false) {
        results.notFound++;
        results.errors.push({
          text: eventTitle.toString(),
          error: 'カレンダーにイベントが見つかりませんでした'
        });
      }
      
    } catch (error) {
      results.failed++;
      results.errors.push({
        text: row && row[0] ? row[0].toString() : '不明',
        error: error.message
      });
      console.error(`行${originalIndex + 3}の処理でエラー: ${error.message}`);
    }
  });

  // 処理結果をログに出力
  console.log(`処理完了: 成功=${results.success}, イベント未発見=${results.notFound}, 未チェック=${results.unchecked}, 失敗=${results.failed}`);
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
    
    // タイトルが一致するイベントを探す（部分一致で検索）
    const matchingEvent = events.find(event => 
      event.getTitle().trim() === title.trim()
    );
    
    if (!matchingEvent) {
      console.log(`イベントが見つかりませんでした: ${title}`);
      return false;
    }

    // イベントの色を変更
    matchingEvent.setColor(color);
    const colorName = color === CalendarApp.EventColor.GREEN ? '緑色' : '赤色';
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

