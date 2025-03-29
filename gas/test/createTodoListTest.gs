/**
 * createToDoListメソッドの基本機能のテスト
 */
function testCreateToDoList_base() {
  // テスト用の予定を作成
  const calendar = CalendarApp.getDefaultCalendar();
  const today = new Date();
  calendar.createEvent(
    "テスト予定:schedule",
    new Date(today.setHours(10, 0)),
    new Date(today.setHours(11, 0))
  );

  // 実行
  createToDoList();
  
  // ドキュメントの内容を確認
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const text = body.getText();
  
  // 検証
  console.log("予定が正しく追加されているか：", text.includes("テスト予定:schedule"));
  console.log("時間が正しく表示されているか：", text.includes("10:00 - 11:00"));
  console.log("チェックボックスが追加されているか：", text.includes("☐"));
}

/**
 * 予定の重複チェック機能のテスト
 */
function testCreateToDoList_duplicate_check() {
  // テスト用の予定を作成
  const calendar = CalendarApp.getDefaultCalendar();
  const today = new Date();
  calendar.createEvent(
    "重複テスト:schedule",
    new Date(today.setHours(10, 0)),
    new Date(today.setHours(11, 0))
  );

  // 2回実行
  createToDoList();
  createToDoList();
  
  // ドキュメントの内容を確認
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();
  
  // 重複カウント
  let count = 0;
  paragraphs.forEach(p => {
    if (p.getText().includes("重複テスト:schedule")) count++;
  });
  
  // 検証
  console.log("予定が重複していないか：", count === 1);
}

/**
 * チェックボックスをリセットする機能のテスト
 */
function testCreateToDoList_checkbox_reset() {
  // テスト用の予定を作成
  const calendar = CalendarApp.getDefaultCalendar();
  const today = new Date();
  calendar.createEvent(
    "リセットテスト:schedule",
    new Date(today.setHours(10, 0)),
    new Date(today.setHours(11, 0))
  );

  // 1回目の実行
  createToDoList();
  
  // 2回目の実行
  createToDoList();
  
  // ドキュメントの内容を確認
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const text = body.getText();
  
  // 検証
  console.log("チェックボックスがリセットされているか：", 
    text.includes("☐") && !text.includes("☑"));
}
