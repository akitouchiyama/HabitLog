// 設定値
const CONFIG = {
  MAX_ROWS: 20,
  COMPLETION_TITLE: '🥳',
  COMPLETION_COLOR: CalendarApp.EventColor.BLUE,
  CHECKED_COLOR: CalendarApp.EventColor.GREEN
};

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
 * カレンダーを取得する共通関数
 * 
 * @returns {GoogleAppsScript.Calendar.Calendar} カレンダー
 */
function getCalendar() {
  const calendarId = PropertiesService.getScriptProperties().getProperty('CALENDAR_ID');
  if (!calendarId) {
    throw new Error('CALENDAR_IDが設定されていません。PropertiesServiceで設定してください。');
  }
  return CalendarApp.getCalendarById(calendarId);
}

/**
 * 一日の範囲の開始時刻と終了時刻を取得する共通関数
 * 
 * @param {Date} date 対象の日付
 * @returns {Object} {startTime, endTime}
 */
function getDayRange(date) {
  const startTime = new Date(date);
  startTime.setHours(0, 0, 0, 0);
  const endTime = new Date(date);
  endTime.setHours(23, 59, 59, 999);
  return { startTime, endTime };
}

/**
 * すべてのタスクがチェックされている場合に完了予定を作成する
 * 
 * @param {Array} dataRows データ行の配列
 * @param {Date} date 対象の日付
 * @param {Array} existingEvents 既存のイベント配列
 */
function createCompletionEventIfAllChecked(dataRows, date, existingEvents) {
  try {
    if (dataRows.length === 0) {
      console.log('タスクがないため、完了予定の作成をスキップします');
      return;
    }

    // すべてのタスクがチェックされているかを確認
    const allChecked = dataRows.every(item => {
      const isChecked = item.row[1]; // C列の値（チェックボックス）
      return isChecked === true;
    });

    if (!allChecked) {
      console.log('未完了のタスクがあるため、完了予定の作成をスキップします');
      return;
    }

    // 既に完了予定があるかチェック
    const completionEventExists = existingEvents.some(event => 
      event.getTitle().trim() === CONFIG.COMPLETION_TITLE
    );

    if (completionEventExists) {
      console.log(`「${CONFIG.COMPLETION_TITLE}」の予定は既に存在します`);
      return;
    }

    // カレンダーを取得して終日予定を作成
    const calendar = getCalendar();
    const completionEvent = calendar.createAllDayEvent(CONFIG.COMPLETION_TITLE, date);
    completionEvent.setColor(CONFIG.COMPLETION_COLOR);
    
    console.log(`すべてのタスクが完了したため、「${CONFIG.COMPLETION_TITLE}」の予定を作成しました`);

  } catch (error) {
    console.error(`完了予定の作成でエラー: ${error.message}`);
  }
}

/**
 * 完了済み予定の色変更を処理する
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet スプレッドシート
 */
function processCompletedEvents(sheet) {
  const today = new Date();
  
  // スプレッドシートからデータを取得
  const dataRows = getTaskDataFromSheet(sheet);
  if (dataRows.length === 0) {
    console.log('予定データがありません');
    return;
  }

  // カレンダーイベントを取得
  const calendar = getCalendar();
  const { startTime, endTime } = getDayRange(today);
  const calendarEvents = calendar.getEvents(startTime, endTime);

  // チェック済み予定の色を変更
  const results = updateCheckedEventsColor(dataRows, calendarEvents);

  // 処理結果をログに出力
  console.log(`処理完了: 成功=${results.success}, イベント未発見=${results.notFound}, 未チェック=${results.unchecked}, 失敗=${results.failed}`);
  if (results.errors.length > 0) {
    console.log('エラー詳細:', results.errors);
  }

  // すべてのタスクがチェックされている場合の完了予定を作成
  createCompletionEventIfAllChecked(dataRows, today, calendarEvents);
}

/**
 * スプレッドシートからタスクデータを取得する
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet スプレッドシート
 * @returns {Array} データ行の配列
 */
function getTaskDataFromSheet(sheet) {
  // B3列以降の予定データの範囲を取得
  const eventRange = sheet.getRange(3, 2, CONFIG.MAX_ROWS, 2); // B3:C(3+MAX_ROWS)
  const values = eventRange.getValues();
  
  // 実際にデータが入っている行をフィルタリング（元のインデックスを保持）
  return values.map((row, originalIndex) => ({ row, originalIndex }))
    .filter(item => {
      const eventTitle = item.row[0]; // B列の値（予定）
      return eventTitle && eventTitle.toString().trim() !== '';
    });
}

/**
 * チェック済み予定の色を更新する
 * 
 * @param {Array} dataRows データ行の配列
 * @param {Array} calendarEvents カレンダーイベントの配列
 * @returns {Object} 処理結果
 */
function updateCheckedEventsColor(dataRows, calendarEvents) {
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
      
      // カレンダーイベントからタイトルが一致するものを探す
      const matchingEvent = calendarEvents.find(event => 
        event.getTitle().trim() === eventTitle.toString().trim()
      );
      
      if (!matchingEvent) {
        results.notFound++;
        results.errors.push({
          text: eventTitle.toString(),
          error: 'カレンダーにイベントが見つかりませんでした'
        });
        return;
      }

      // イベントの色を変更
      matchingEvent.setColor(CONFIG.CHECKED_COLOR);
      const colorName = CONFIG.CHECKED_COLOR === CalendarApp.EventColor.GREEN ? '緑色' : '赤色';
      console.log(`イベントの色を${colorName}に変更しました: ${eventTitle}`);
      results.success++;
      
    } catch (error) {
      results.failed++;
      results.errors.push({
        text: row && row[0] ? row[0].toString() : '不明',
        error: error.message
      });
      console.error(`行${originalIndex + 3}の処理でエラー: ${error.message}`);
    }
  });

  return results;
}