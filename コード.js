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
    sheetId = ss.getId();
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

  const countStatus = getRosterImageCountStatus();
  return { success: true, sheetUrl: sheetUrl, folderUrl: folderUrl, sheetId: sheetId, isApiReady: !!apiKey, logs: messages, countStatus: countStatus };
}

function isActiveRosterEntry(id, name, absent) {
  return !!(id || name) && !absent;
}

function getFolderFileList(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const fileList = [];
  while (files.hasNext()) { fileList.push(files.next()); }
  fileList.sort((a, b) => a.getName().localeCompare(b.getName(), undefined, { numeric: true, sensitivity: 'base' }));
  return fileList;
}

function getRosterSheetData(sheetId) {
  const ss = SpreadsheetApp.openById(sheetId);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.getSheets()[0];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { sheet: sheet, data: [], lastRow: lastRow };
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  return { sheet: sheet, data: data, lastRow: lastRow };
}

/**
 * UI用: 名簿（Absent除外）とフォルダ内画像枚数の照合
 */
function getRosterImageCountStatus() {
  const folderId = getProperty('FOLDER_ID');
  const sheetId = getProperty('SHEET_ID');
  if (!folderId || !sheetId) {
    return { success: false, message: "セットアップ未完了" };
  }

  try {
    const fileList = getFolderFileList(folderId);
    const roster = getRosterSheetData(sheetId);
    if (roster.lastRow < 2) {
      return { success: false, message: "名簿データなし" };
    }

    const expected = roster.data.filter(function(r) {
      return isActiveRosterEntry(r[0], r[1], r[2]);
    }).length;
    const actual = fileList.length;
    const match = expected === actual;

    return {
      success: true,
      expected: expected,
      actual: actual,
      match: match,
      message: match
        ? "✅ 枚数一致: 受験者 " + expected + " 名 = 画像 " + actual + " 枚"
        : "⚠️ 枚数不一致: 受験者 " + expected + " 名 vs 画像 " + actual + " 枚 — このまま紐付けすると名簿順と画像の対応がずれる可能性があります"
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * 機能1: リネーム / 名簿紐付け処理
 * mode: 'rename' = ファイル名変更 + D列にFile ID
 *       'link'   = 紐付けのみ（ファイル名は維持）+ D列にFile ID
 *       'skip'   = 何もしない
 */
function renameFilesProcess(mode) {
  if (!mode) mode = 'rename';
  if (mode === 'skip') {
    return "処理をスキップしました。（ファイル名・名簿紐付けは変更していません）";
  }
  const doRename = mode !== 'link';

  const folderId = getProperty('FOLDER_ID');
  const sheetId = getProperty('SHEET_ID');
  if (!folderId || !sheetId) throw new Error("セットアップ未完了");

  const fileList = getFolderFileList(folderId);
  const roster = getRosterSheetData(sheetId);
  const sheet = roster.sheet;
  const data = roster.data;

  if (roster.lastRow < 2) return "名簿データなし";

  const expected = data.filter(function(r) {
    return isActiveRosterEntry(r[0], r[1], r[2]);
  }).length;
  const actual = fileList.length;
  let processLog = [doRename ? "モード: リネーム" : "モード: 紐付けのみ"];

  if (actual !== expected) {
    processLog.push("⚠️【警告】枚数不一致: 受験者" + expected + "名 vs 画像" + actual + "枚");
    processLog.push("   → 名簿順と画像の対応がずれる可能性があります。枚数を確認してください。");
  } else {
    processLog.push("✅ 枚数OK: 受験者" + expected + "名 = 画像" + actual + "枚");
  }
  processLog.push("--------------------------------------------------");

  let fileIndex = 0;
  let updateData = []; 

  for (let i = 0; i < data.length; i++) {
    // 0:ID, 1:Name, 2:Absent, 3:FileID
    const [id, name, absent, currentFileId] = data[i];
    
    if (!isActiveRosterEntry(id, name, absent)) {
      updateData.push([currentFileId]);
      if (absent && (id || name)) {
        processLog.push("[スキップ] " + name + " (欠席)");
      }
      continue;
    }

    if (fileIndex < fileList.length) {
      const file = fileList[fileIndex];
      const orig = file.getName();

      try {
        if (doRename) {
          const ext = orig.includes('.') ? orig.substring(orig.lastIndexOf('.')) : '';
          const newName = `${id}_${name}${ext}`;
          if (orig !== newName) {
            file.setName(newName);
            processLog.push(`[リネーム] ${orig} -> ${newName}`);
          } else {
            processLog.push(`[維持] ${orig}`);
          }
        } else {
          processLog.push(`[紐付け] ${orig} → ${name} (ID:${id})`);
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
 * 機能2: OCR処理（結果をTSV形式で返却、シートには書き込まない）
 */
function ocrProcess(targetCoordsList) {
  const apiKey = getProperty('GCP_API_KEY');
  const sheetId = getProperty('SHEET_ID');
  if (!apiKey) throw new Error("GCP_API_KEY未設定");
  if (!sheetId) throw new Error("セットアップ未完了");

  const ss = SpreadsheetApp.openById(sheetId);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.getSheets()[0];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("データなし");

  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const resultCount = targetCoordsList && targetCoordsList.length > 0 ? targetCoordsList.length : 1;
  const headers = buildOcrHeaders(resultCount);
  const rows = [];
  const processLog = [];

  for (let i = 0; i < data.length; i++) {
    const [id, name, absent, fileId] = data[i];
    const rowNum = i + 2;

    if (!fileId) continue;

    const base = [id, name, absent, fileId];
    try {
      const texts = callVisionApi(fileId, apiKey, targetCoordsList);
      rows.push(base.concat(texts));
      processLog.push(`[OCR完了] 行${rowNum}: ${name}`);
    } catch (e) {
      const errCells = new Array(resultCount).fill(`(エラー: ${e.message})`);
      rows.push(base.concat(errCells));
      processLog.push(`[OCRエラー] 行${rowNum}: ${e.message}`);
    }
  }

  if (rows.length === 0) throw new Error("未処理なし（File ID が設定された行がありません）");

  return {
    success: true,
    headers: headers,
    rows: rows,
    tsv: rowsToTsv(headers, rows),
    log: processLog.join('\n'),
    defaultSheetId: sheetId
  };
}

/**
 * OCR結果をスプレッドシートにエクスポート
 * exportMode: 'existing' = 指定IDのスプレッドシート / 'new' = 新規作成
 */
function exportOcrResults(exportMode, targetSheetId, headers, rows) {
  if (!headers || !rows || rows.length === 0) throw new Error("エクスポートするデータがありません");

  let ss, sheet, sheetUrl, message;

  if (exportMode === 'new') {
    const name = `OCR結果_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm')}`;
    ss = SpreadsheetApp.create(name);
    sheet = ss.getSheets()[0];
    sheet.setName('OCR Results');
    message = `新規スプレッドシート「${name}」を作成しました。`;
  } else {
    const id = extractSpreadsheetId(targetSheetId);
    if (!id) throw new Error("スプレッドシートIDが未入力です");
    try {
      ss = SpreadsheetApp.openById(id);
    } catch (e) {
      throw new Error("スプレッドシートを開けません: " + e.message);
    }
    sheet = ss.getSheetByName('OCR Results');
    if (!sheet) {
      sheet = ss.insertSheet('OCR Results');
    } else {
      sheet.clear();
    }
    message = `スプレッドシートにエクスポートしました（シート: OCR Results）。`;
  }

  writeOcrTableToSheet(sheet, headers, rows);
  sheetUrl = ss.getUrl();

  return { success: true, message: message, sheetUrl: sheetUrl, sheetId: ss.getId() };
}

function buildOcrHeaders(resultCount) {
  const headers = ['ID', 'Name', 'Absent', 'File ID'];
  for (let i = 1; i <= resultCount; i++) headers.push(`OCR Result ${i}`);
  return headers;
}

function rowsToTsv(headers, rows) {
  const lines = [headers.map(escapeTsvCell).join('\t')];
  rows.forEach(row => lines.push(row.map(escapeTsvCell).join('\t')));
  return lines.join('\n');
}

function escapeTsvCell(val) {
  const s = String(val == null ? '' : val);
  if (/[\t\n\r"]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

function extractSpreadsheetId(input) {
  if (!input) return '';
  const m = String(input).match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : String(input).trim();
}

function writeOcrTableToSheet(sheet, headers, rows) {
  const table = [headers].concat(rows);
  sheet.clear();
  sheet.getRange(1, 1, table.length, headers.length).setValues(table);
  sheet.getRange(1, 1, 1, headers.length)
       .setBackground('#efefef')
       .setFontWeight('bold');
  sheet.setFrozenRows(1);
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

  if (!json.responses || json.responses.length === 0) {
    return targetCoordsList && targetCoordsList.length > 0
           ? new Array(targetCoordsList.length).fill("(文字なし)")
           : ["(文字なし)"];
  }
  
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
  if (!annotation.pages || annotation.pages.length === 0) {
    return new Array(targetCoordsList.length).fill("(文字なし)");
  }

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
    const fileList = getFolderFileList(folderId);

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
