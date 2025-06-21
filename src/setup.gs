/**
 * 環境変数を設定するためのヘルパー関数
 * 
 * @param {string} documentId GoogleドキュメントのID
 * @param {string} calendarId GoogleカレンダーのID
 * 
 * 使用例（実際のIDに置き換えてください）:
 * setupEnvironmentVariables('1XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX', 'your-calendar@gmail.com')
 */
function setupEnvironmentVariables(documentId, calendarId) {
  // パラメータのチェック
  if (!documentId || !calendarId) {
    console.error('エラー: documentIdとcalendarIdの両方を指定してください');
    console.log('使用例: setupEnvironmentVariables("ドキュメントID", "カレンダーID")');
    return;
  }
  
  if (typeof documentId !== 'string' || typeof calendarId !== 'string') {
    console.error('エラー: documentIdとcalendarIdは文字列で指定してください');
    return;
  }
  
  try {
    // PropertiesServiceに環境変数を設定
    const properties = PropertiesService.getScriptProperties();
    properties.setProperties({
      'DOCUMENT_ID': documentId,
      'CALENDAR_ID': calendarId
    });
    
    console.log('環境変数の設定が完了しました:');
    console.log('DOCUMENT_ID:', documentId);
    console.log('CALENDAR_ID:', calendarId);
    
  } catch (error) {
    console.error('環境変数の設定でエラーが発生しました:', error.message);
  }
}

/**
 * ドキュメントIDのみを設定する関数
 * 
 * @param {string} documentId GoogleドキュメントのID
 * 
 * 使用例（実際のIDに置き換えてください）:
 * setDocumentId('1XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX')
 */
function setDocumentId(documentId) {
  if (!documentId || typeof documentId !== 'string') {
    console.error('エラー: ドキュメントIDを文字列で指定してください');
    console.log('使用例: setDocumentId("ドキュメントID")');
    return;
  }
  
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty('DOCUMENT_ID', documentId);
    console.log('ドキュメントIDを設定しました:', documentId);
  } catch (error) {
    console.error('ドキュメントIDの設定でエラーが発生しました:', error.message);
  }
}

/**
 * カレンダーIDのみを設定する関数
 * 
 * @param {string} calendarId GoogleカレンダーのID
 * 
 * 使用例（実際のIDに置き換えてください）:
 * setCalendarId('your-calendar@gmail.com')
 */
function setCalendarId(calendarId) {
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
    const documentId = properties.getProperty('DOCUMENT_ID');
    const calendarId = properties.getProperty('CALENDAR_ID');
    
    console.log('現在の環境変数設定:');
    console.log('DOCUMENT_ID:', documentId || '未設定');
    console.log('CALENDAR_ID:', calendarId || '未設定');
    
    if (!documentId || !calendarId) {
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
    properties.deleteProperty('DOCUMENT_ID');
    properties.deleteProperty('CALENDAR_ID');
    
    console.log('環境変数をクリアしました。');
    
  } catch (error) {
    console.error('環境変数のクリアでエラーが発生しました:', error.message);
  }
} 