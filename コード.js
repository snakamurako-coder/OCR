/**
 * 【重要】スクリプトプロパティ
 * 1. GCP_API_KEY : (手動設定)
 * 2. SHEET_ID    : (自動)
 * 3. FOLDER_ID   : (自動)
 */

const SHEET_NAME = 'Roster';

function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('採点・文字起こしツール v6 (列構成変更版)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getProperty(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

/**
 * 初期セットアップ
 * ヘッダー構成を修正: ID, Name, Absent, File ID, OCR Result 1
 */
function setupResources() {
  const props = PropertiesService.getScriptProperties();
  let sheetId = props.getProperty('SHEET_ID');
  let folderId = props.getProperty('FOLDER_ID');
  const apiKey = props.getProperty('GCP_API_KEY');
  
  let messages = [];
  let sheetUrl = "", folderUrl = "";

  // 1. シート作成・確認
  if (sheetId) {
    try { sheetUrl = SpreadsheetApp.openById(sheetId).getUrl(); } catch (e) { messages.push("シートID無効"); }
  } else {
    const ss = SpreadsheetApp.create("OCR管理シート_自動生成");
    const sheet = ss.getSheets()[0];
    sheet.setName(SHEET_NAME);
    // ★変更: File IDをD列に移動
    sheet.getRange("A1:E1").setValues([["ID", "Name", "Absent", "File ID", "OCR Result 1"]]);
    sheet.getRange("A1:E1").setBackground("#efefef").setFontWeight("bold");
    sheet.setFrozenRows(1);
    props.setProperty('SHEET_ID', ss.getId());
    sheetUrl = ss.getUrl();
    messages.push("管理シート作成完了 (新カラム構成)");
  }

  // 2. フォルダ作成・確認
  if (folderId) {
    try { folderUrl = DriveApp.getFolderById(folderId).getUrl(); } catch (e) { messages.push("フォルダID無効"); }
  } else {
    const parents = DriveApp.getFileById(ScriptApp.getScriptId()).getParents();
    const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    const newFolder = parent.createFolder("OCR_Images_自動生成");
    props.setProperty('FOLDER_ID', newFolder.getId());
    folderUrl = newFolder.getUrl();
    messages.push("画像フォルダ作成完了");
  }

  return { success: true, sheetUrl: sheetUrl, folderUrl: folderUrl, isApiReady: !!apiKey, logs: messages };
}

/**
 * 機能1: リネーム処理
 * File ID の参照・書き込み先を D列(index 3) に変更
 */
function renameFilesProcess() {
  const folderId = getProperty('FOLDER_ID');
  const sheetId = getProperty('SHEET_ID');
  if (!folderId || !sheetId) throw new Error("セットアップ未完了");

  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const fileList = [];
  while (files.hasNext()) { fileList.push(files.next()); }
  fileList.sort((a, b) => a.getName().localeCompare(b.getName(), undefined, {numeric: true, sensitivity: 'base'}));

  const ss = SpreadsheetApp.openById(sheetId);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.getSheets()[0];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return "名簿データなし";

  // A列(ID)〜D列(File ID)まで取得
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();

  // 枚数チェック
  const expected = data.filter(r => r[2] == "").length; 
  const actual = fileList.length;
  let processLog = [];

  if (actual !== expected) {
    processLog.push(`⚠️【警告】枚数不一致: 出席${expected}名 vs 画像${actual}枚`);
  } else {
    processLog.push(`✅ 枚数OK: ${expected}名 = ${actual}枚`);
  }
  processLog.push("--------------------------------------------------");

  let fileIndex = 0;
  let updateData = []; 

  for (let i = 0; i < data.length; i++) {
    // 0:ID, 1:Name, 2:Absent, 3:FileID
    const [id, name, absent, currentFileId] = data[i];
    
    if (!id && !name) { updateData.push([currentFileId]); continue; }
    if (absent) {
      processLog.push(`[スキップ] ${name} (欠席)`);
      updateData.push([currentFileId]);
      continue;
    }

    if (fileIndex < fileList.length) {
      const file = fileList[fileIndex];
      const orig = file.getName();
      const ext = orig.includes('.') ? orig.substring(orig.lastIndexOf('.')) : '';
      const newName = `${id}_${name}${ext}`; 

      try {
        if (orig !== newName) {
           file.setName(newName);
           processLog.push(`[リネーム] ${orig} -> ${newName}`);
        } else {
           processLog.push(`[維持] ${orig}`);
        }
        updateData.push([file.getId()]);
        fileIndex++; 
      } catch (e) {
        processLog.push(`[エラー] ${orig}: ${e.message}`);
        updateData.push([currentFileId]);
      }
    } else {
      processLog.push(`❌ [不足] ${name}`);
      updateData.push([currentFileId]);
    }
  }
  
  // D列(4列目)に書き込み
  sheet.getRange(2, 4, updateData.length, 1).setValues(updateData);
  return processLog.length > 0 ? processLog.join('\n') : "変更なし";
}

/**
 * 機能2: OCR処理
 * 結果を E列(5列目) 以降に横展開して書き込み
 */
function ocrProcess(targetCoordsList) {
  const apiKey = getProperty('GCP_API_KEY');
  const sheetId = getProperty('SHEET_ID');
  if (!apiKey) throw new Error("GCP_API_KEY未設定");

  const ss = SpreadsheetApp.openById(sheetId);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.getSheets()[0];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return "データなし";

  // A〜D列を取得 (File IDは D列=index 3)
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  let processLog = [];

  // ヘッダーの自動拡張チェック
  const resultCount = targetCoordsList ? targetCoordsList.length : 1;
  ensureHeaders(sheet, resultCount);

  for (let i = 0; i < data.length; i++) {
    const name = data[i][1];
    const fileId = data[i][3]; // D列

    // ファイルIDがある場合のみ実行（上書き防止判定は簡易的に省略、必要ならE列チェックを追加）
    if (fileId) {
      try {
        // 配列で受け取る ['Res1', 'Res2', ...]
        const texts = callVisionApi(fileId, apiKey, targetCoordsList);
        
        processLog.push(`[OCR完了] 行${i + 2}: ${name}`);
        
        // E列(5列目)から横方向に書き込み
        if (texts.length > 0) {
          sheet.getRange(i + 2, 5, 1, texts.length).setValues([texts]);
        }
      } catch (e) {
        processLog.push(`[OCRエラー] 行${i + 2}: ${e.message}`);
      }
    }
  }
  return processLog.length > 0 ? processLog.join('\n') : "未処理なし";
}

/**
 * ヘッダー自動拡張
 * 現在の列数を確認し、OCR Result N が足りなければ追加する
 */
function ensureHeaders(sheet, count) {
  const startCol = 5; // E列
  const lastCol = sheet.getLastColumn();
  const needed = (startCol - 1) + count; // D列まで + 必要な数

  if (lastCol < needed) {
    const headers = [];
    for (let i = lastCol - (startCol - 1) + 1; i <= count; i++) {
      headers.push(`OCR Result ${i}`);
    }
    if (headers.length > 0) {
      sheet.getRange(1, lastCol + 1, 1, headers.length)
           .setValues([headers])
           .setBackground("#efefef")
           .setFontWeight("bold");
    }
  }
}

/**
 * Vision API
 * 戻り値を「文字列の配列」に変更
 */
function callVisionApi(fileId, apiKey, targetCoordsList) {
  const url = `https://vision.googleapis.com/v1/images:annotate?key=${apiKey}`;
  const file = DriveApp.getFileById(fileId);
  const base64Image = Utilities.base64Encode(file.getBlob().getBytes());

  const payload = {
    requests: [{
      image: { content: base64Image },
      features: [{ type: "DOCUMENT_TEXT_DETECTION", maxResults: 1 }],
      imageContext: { languageHints: ["en-t-i0-handwrit"] }
    }]
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  const json = JSON.parse(response.getContentText());

  if (json.error) throw new Error(json.error.message);
  
  // 何もなければ空文字を返す
  const annotation = json.responses[0].fullTextAnnotation;
  if (!annotation) {
    return targetCoordsList && targetCoordsList.length > 0 
           ? new Array(targetCoordsList.length).fill("(文字なし)")
           : ["(文字なし)"];
  }

  // 領域指定なし -> 全文を1要素の配列で返す
  if (!targetCoordsList || targetCoordsList.length === 0) {
    return [annotation.text];
  }

  // 複数領域フィルタリング
  let buckets = new Array(targetCoordsList.length).fill().map(() => []);

  annotation.pages.forEach(p => p.blocks.forEach(b => b.paragraphs.forEach(pr => pr.words.forEach(w => {
    const box = w.boundingBox.vertices;
    for (let i = 0; i < targetCoordsList.length; i++) {
      if (isCenterInRect(box, targetCoordsList[i])) {
        buckets[i].push(w.symbols.map(s => s.text).join(''));
        break; 
      }
    }
  }))));

  // 結果を配列にして返す
  return buckets.map(bucket => bucket.length > 0 ? bucket.join(' ') : "(空白)");
}

function isCenterInRect(vertices, rect) {
  const xs = vertices.map(v => v.x || 0), ys = vertices.map(v => v.y || 0);
  const cx = (Math.min(...xs) + Math.max(...xs)) / 2;
  const cy = (Math.min(...ys) + Math.max(...ys)) / 2;
  return (cx >= rect.x && cx <= rect.x + rect.w && cy >= rect.y && cy <= rect.y + rect.h);
}

// --- UI用: 画像取得 ---
function getSampleImageForUi(index) {
  const folderId = getProperty('FOLDER_ID');
  if (!folderId) return { success: false, message: "フォルダID未設定" };
  
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const fileList = [];
    while (files.hasNext()) { fileList.push(files.next()); }

    fileList.sort((a, b) => a.getName().localeCompare(b.getName(), undefined, {numeric: true, sensitivity: 'base'}));

    if (fileList.length === 0) return { success: false, message: "フォルダが空です" };
    
    if (index < 0) index = 0;
    if (index >= fileList.length) index = fileList.length - 1;

    const file = fileList[index];
    if (file.getMimeType().startsWith('image/')) {
      return { 
        success: true, 
        name: file.getName(), 
        mimeType: file.getMimeType(), 
        base64: Utilities.base64Encode(file.getBlob().getBytes()),
        currentIndex: index,
        totalCount: fileList.length
      };
    } else {
       return { success: false, message: "画像ファイルではありません" };
    }
  } catch(e) {
    return { success: false, message: "画像取得エラー: " + e.message };
  }
}