/**
 * 環境変数を設定するためのヘルパー関数
 * 実際の値に置き換えてください
 * 
 * @param {string} sheetId スプレッドシートのID
 * @param {string} calendarId GoogleカレンダーのID
 */
function setupEnvironmentVariables(sheetId = '1XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX', calendarId = 'your-calendar@gmail.com') {
  // パラメータのチェック
  if (!sheetId || !calendarId) {
    console.error('エラー: sheetIdとcalendarIdの両方を指定してください');
    console.log('使用例: setupEnvironmentVariables("スプレッドシートID", "カレンダーID")');
    return;
  }
  
  if (typeof sheetId !== 'string' || typeof calendarId !== 'string') {
    console.error('エラー: sheetIdとcalendarIdは文字列で指定してください');
    return;
  }
  
  try {
    // PropertiesServiceに環境変数を設定
    const properties = PropertiesService.getScriptProperties();
    properties.setProperties({
      'SHEET_ID': sheetId,
      'CALENDAR_ID': calendarId
    });
    
    console.log('環境変数の設定が完了しました:');
    console.log('SHEET_ID:', sheetId);
    console.log('CALENDAR_ID:', calendarId);
    
  } catch (error) {
    console.error('環境変数の設定でエラーが発生しました:', error.message);
  }
}

/**
 * スプレッドシートIDのみを設定する関数
 * 実際のIDに置き換えてください
 *
 * @param {string} sheetId スプレッドシートのID
 */
function setSheetId(sheetId = '1XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX') {
  if (!sheetId || typeof sheetId !== 'string') {
    console.error('エラー: スプレッドシートIDを文字列で指定してください');
    console.log('使用例: setSheetId("スプレッドシートID")');
    return;
  }
  
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty('SHEET_ID', sheetId);
    console.log('スプレッドシートIDを設定しました:', sheetId);
  } catch (error) {
    console.error('スプレッドシートIDの設定でエラーが発生しました:', error.message);
  }
}

/**
 * カレンダーIDのみを設定する関数
 * 実際のIDに置き換えてください
 *
 * @param {string} calendarId GoogleカレンダーのID
 */
function setCalendarId(calendarId = 'your-calendar@gmail.com') {
  if (!calendarId || typeof calendarId !== 'string') {
    console.error('エラー: カレンダーIDを文字列で指定してください');
    console.log('使用例: setCalendarId("カレンダーID")');
    return;
  }
  
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty('CALENDAR_ID', calendarId);
    console.log('カレンダーIDを設定しました:', calendarId);
  } catch (error) {
    console.error('カレンダーIDの設定でエラーが発生しました:', error.message);
  }
}

/**
 * 設定された環境変数を確認する関数
 */
function checkEnvironmentVariables() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const sheetId = properties.getProperty('SHEET_ID');
    const calendarId = properties.getProperty('CALENDAR_ID');
    
    console.log('現在の環境変数設定:');
    console.log('SHEET_ID:', sheetId || '未設定');
    console.log('CALENDAR_ID:', calendarId || '未設定');
    
    if (!sheetId || !calendarId) {
      console.warn('警告: 環境変数が正しく設定されていません。setupEnvironmentVariables()を実行してください。');
    } else {
      console.log('✅ 環境変数は正しく設定されています。');
    }
    
  } catch (error) {
    console.error('環境変数の確認でエラーが発生しました:', error.message);
  }
}

/**
 * 環境変数をクリアする関数（必要に応じて使用）
 */
function clearEnvironmentVariables() {
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.deleteProperty('SHEET_ID');
    properties.deleteProperty('CALENDAR_ID');
    
    console.log('環境変数をクリアしました。');
    
  } catch (error) {
    console.error('環境変数のクリアでエラーが発生しました:', error.message);
  }
} 