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


//メニュー
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('スマート補正')
    .addItem('現在の競技を統一表現にして更新する', 'runNowSwimtype')
    .addSeparator()
    .addItem('次の競技を統一表現にする', 'runNextSwimtype')
    .addSeparator()
    .addItem('辞書エディタを開く', 'openJsonEditor')
    .addToUi();
}


//----------------------------------------------------------------------------

//メイン関数
function runNextSwimtype1() { executeWithAlert('水泳管理', 'E', '次の競技'); }
function runNowSwimtype()   { executeWithAlert('水泳管理', 'D', '現在の競技'); }

function executeWithAlert(sheetName, colLetter, label) {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(`${label}の補正`, `統一表現に変換しますか？`, ui.ButtonSet.OK_CANCEL);
  if (res !== ui.Button.OK) return;

  runCorrectionForRange(sheetName, colLetter);

  SpreadsheetApp.getUi().alert('処理が完了しました。');
}

//-----------------------------------------------------------------


//  メイン処理
function runCorrectionForRange(sheetName, colLetter) {
  const dict = loadDict();
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName);

  //選手名を配列へ
  const bValues = sheet.getRange('B:B').getValues(); 
  let lastRow = 0;
  for (let i = bValues.length - 1; i >= 0; i--) {
    if (bValues[i][0]) { lastRow = i + 1; break; }
  }

  const startRow = 5;
  if (lastRow < startRow) {
    SpreadsheetApp.getUi().alert('処理対象のデータがありません。');
    return;
  }

  const range = sheet.getRange(`${colLetter}${startRow}:${colLetter}${lastRow}`);
  const originalValues = range.getValues().flat();

  if (sheetName === "水泳管理") {
    backupBeforeCorrection(sheet, colLetter, originalValues);
  }

  const results = originalValues.map(v => {
    const text = (v || '').toString().trim();
    if (!text) return [''];
    return [normalizeSwimtypeName(text, dict)];
  });

   range.setValues(results);

  if (sheetName === "水泳管理") {
    slideSwimtypes(ss, startRow, lastRow);
  }

}

//ワイルドカード変換処理
function normalizeSwimtypeName(input, dict) {
  const text = (input || "").trim();
  if (!text) return "";

  const hits = new Set();

  for (const Swimtype in dangerousMap) {
    const patterns = dangerousMap[Swimtype];
    if (patterns.some(p => remaining.startsWith(p))) {
      hits.add(Swimtype);
    }
  }

  //辞書チェック（skipDict=false のときのみ）
  if (!skipDict) {
    for (const Swimtype in dict) {
      if (!Object.prototype.hasOwnProperty.call(dict, Swimtype)) continue;

      // 完全一致
      if (remaining.includes(Swimtype)) {
        hits.add(Swimtype);
        continue;
      }

      // パターン一致
      const patterns = dict[Swimtype];
      if (Array.isArray(patterns) && patterns.some(p => p && remaining.includes(p))) {
        hits.add(Swimtype);
      }
    }
  }

  //結果の決定
  let result = "";
  const hitArray = Array.from(hits);

  if (hitArray.length === 0) {
    result = remaining;
  } else if (hitArray.length === 1) {
    result = hitArray[0];
  } else {
    result = hitArray.join("/");
  }

  return result.trim();
}

