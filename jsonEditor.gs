
//  JSON辞書ファイル設定
const DICT_FILE_ID = 'XXXXXXXXXXXXXXXXXXXXXXXX'; 
const BK_FOLDER_ID = 'XXXXXXXXXXXXXXXXXXXXXXXX';

// キャッシュ
let cachedDict = null;

//JSON辞書読み込み
function loadDict() {
  if (cachedDict) return cachedDict;

  const file = DriveApp.getFileById(DICT_FILE_ID);
  const jsonText = file.getBlob().getDataAsString('UTF-8');
  cachedDict = JSON.parse(jsonText);
  return cachedDict;
}

//JSONエディタを開く（モーダル）
function openJsonEditor() {
  const file = DriveApp.getFileById(DICT_FILE_ID);
  const jsonText = file.getBlob().getDataAsString('UTF-8');

  const template = HtmlService.createTemplateFromFile('json_editor');
  template.jsonText = JSON.stringify(JSON.parse(jsonText), null, 2);

  const html = template.evaluate()
    .setWidth(800)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, '変換辞書エディタ');
}

//JSON保存処理
function saveJson(newJsonText) {
  let parsed;
  try {
    parsed = JSON.parse(newJsonText);
  } catch (e) {
    return { success: false, message: 'JSONの構文エラー：' + e };
  }

  // バックアップ
  try {
    createBackupFile();
  } catch (e) {
    // バックアップ失敗 → ユーザーに確認
    const ui = SpreadsheetApp.getUi();
    const res = ui.alert(
      'バックアップに失敗しました',
      'バックアップなしでファイルを上書き保存しますか？\n\nエラー内容：' + e,
      ui.ButtonSet.OK_CANCEL
    );

    if (res !== ui.Button.OK) {
      return { success: false, message: '保存を中止しました。' };
    }
  }

  // 保存処理
  try {
    const file = DriveApp.getFileById(DICT_FILE_ID);
    file.setContent(JSON.stringify(parsed, null, 2));
  } catch (e) {
    return { success: false, message: 'JSON保存に失敗：' + e };
  }

  cachedDict = parsed;
  return { success: true, message: '保存しました。' };
}


//  バックアップ作成
function createBackupFile() {
  const file = DriveApp.getFileById(DICT_FILE_ID);
  const bkFolder = DriveApp.getFolderById(BK_FOLDER_ID);

  const today = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd");

  let baseName = `dictionary_${today}.json`;
  let finalName = baseName;

  let counter = 1;
  while (fileExistsInFolder(bkFolder, finalName)) {
    const suffix = String(counter).padStart(2, '0');
    finalName = `dictionary_${today}_${suffix}.json`;
    counter++;
  }

  file.makeCopy(finalName, bkFolder);
}

function fileExistsInFolder(folder, name) {
  const files = folder.getFilesByName(name);
  return files.hasNext();
}



