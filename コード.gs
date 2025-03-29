// TODOリストフォルダ
const FOLDER_ID = "1DywCksLqibr2hNmtAQdCHFVJPJzoTwAJ";

/**
 * ToDoリストを作成するメイン関数。
 */
function createToDoList() {
  const today = new Date()

  const folder = DriveApp.getFolderById(FOLDER_ID);

  // 各フォルダ(年、月フォルダ)がなければ新規作成
  const year = today.getFullYear();
  const yearFolder = findOrCreateFolder(folder, year + "年");
  const month = today.getMonth() + 1;
  const monthFolder = findOrCreateFolder(yearFolder, month + "月");

  // 先月のフォルダを取得
  const lastMonth = month == 1 ? 12 : month - 1;
  var lastMonthFolder;
  if(lastMonth == 12) {
    // 先月が12月だった場合、去年のフォルダから探す
    const lastYear = year - 1;    
    const lastYearFolder = findOrCreateFolder(folder, lastYear + "年");
    lastMonthFolder = findOrCreateFolder(lastYearFolder, lastMonth + "月");
  }else {
    lastMonthFolder = findOrCreateFolder(yearFolder, lastMonth + "月");
  }

  // 今月のフォルダ、もしくは先月のフォルダから最新のファイルを取得
  var latestFile = findLatestFile(monthFolder) || findLatestFile(lastMonthFolder);
  if (latestFile) {
    // もしファイルが存在すればそのファイルをコピー
    const title = Utilities.formatDate(today, 'JST', 'MMdd');
    const newfile = latestFile.makeCopy(title);
    // コピーしたファイルを今月のフォルダに移動
    newfile.moveTo(monthFolder);
  } else {
    throw "コピー元のファイルが見つかりませんでした。";
  }
}

/**
 * フォルダの検索、及び作成処理
 * 
 * @param {parentFolder} 親フォルダ
 * @param {folderName} 検索・作成対象のフォルダ名
 */
function findOrCreateFolder(parentFolder, folderName) {
  // 親フォルダ内で指定したフォルダを探す
  const folder = parentFolder.getFoldersByName(folderName);
  
  // フォルダが存在すればそのフォルダを返す
  if (folder.hasNext()) {
    return folder.next();
  }

  // フォルダが存在しなければ、新しく作成する
  return parentFolder.createFolder(folderName);
}

/**
 * 指定したフォルダ内の最新のファイルを探す。
 * 
 * @param {Folder} folder 検索するフォルダ。
 * @return {File} 最新のファイル。見つからない場合はnullを返す。
 */
function findLatestFile(folder) {
  // フォルダ内の全てのファイルを取得
  const files = folder.getFiles();
  // もしファイルが存在すれば最新のファイルを返す
  if (files.hasNext()) {
    return files.next();
  }
  // フォルダ内にファイルがなければnullを返す
  return null;
}