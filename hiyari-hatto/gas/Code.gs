/**
 * ヒヤリハット報告アプリ - Google Apps Script
 *
 * 【セットアップ手順】
 * 1. Google スプレッドシートを新規作成
 * 2. 拡張機能 → Apps Script を開く
 * 3. このコードを貼り付けて保存
 * 4. 関数「initSheet」を実行（ヘッダー行を自動作成）
 * 5. デプロイ → 新しいデプロイ → ウェブアプリ
 *    - 実行するユーザー：自分
 *    - アクセスできるユーザー：全員
 * 6. 表示されたURLをコピーして index.html の GAS_URL に貼り付け
 */

// ===== 初期セットアップ =====

/**
 * シートの初期化（ヘッダー行を作成）
 * 最初に1回だけ実行してください
 */
function initSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 「報告データ」シートを作成（なければ）
  var sheet = ss.getSheetByName('報告データ');
  if (!sheet) {
    sheet = ss.insertSheet('報告データ');
  }

  // ヘッダー行
  var headers = ['ID', '発生日時', '分類', '内容', '深刻度', '報告者', '報告日時'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダーの書式
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e53e3e');
  headerRange.setFontColor('#ffffff');

  // 列幅の調整
  sheet.setColumnWidth(1, 140);  // ID
  sheet.setColumnWidth(2, 160);  // 発生日時
  sheet.setColumnWidth(3, 80);   // 分類
  sheet.setColumnWidth(4, 400);  // 内容
  sheet.setColumnWidth(5, 80);   // 深刻度
  sheet.setColumnWidth(6, 100);  // 報告者
  sheet.setColumnWidth(7, 160);  // 報告日時

  // 1行目を固定
  sheet.setFrozenRows(1);

  Logger.log('シートを初期化しました');
}

// ===== Web API =====

/**
 * GET リクエスト：報告データの取得
 */
function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('報告データ');
    if (!sheet) {
      return createJsonResponse({ success: false, error: 'シートが見つかりません' });
    }

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return createJsonResponse({ success: true, reports: [] });
    }

    var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    var reports = data.map(function(row) {
      return {
        id: String(row[0]),
        datetime: formatDateForJson(row[1]),
        category: row[2],
        content: row[3],
        severity: row[4],
        reporter: row[5],
        createdAt: formatDateForJson(row[6])
      };
    });

    // 新しい順にソート
    reports.reverse();

    return createJsonResponse({ success: true, reports: reports });
  } catch (err) {
    return createJsonResponse({ success: false, error: err.message });
  }
}

/**
 * POST リクエスト：報告の追加・削除
 */
function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    var action = params.action || 'add';

    if (action === 'add') {
      return addReport(params);
    } else if (action === 'delete') {
      return deleteReport(params.id);
    } else {
      return createJsonResponse({ success: false, error: '不明なアクション: ' + action });
    }
  } catch (err) {
    return createJsonResponse({ success: false, error: err.message });
  }
}

/**
 * 報告を追加
 */
function addReport(params) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('報告データ');
  if (!sheet) {
    return createJsonResponse({ success: false, error: 'シートが見つかりません。initSheet を実行してください。' });
  }

  var id = String(params.id || Date.now());
  var row = [
    id,
    params.datetime || '',
    params.category || '',
    params.content || '',
    params.severity || '',
    params.reporter || '未記入',
    new Date()
  ];

  sheet.appendRow(row);

  // 深刻度に応じて行の色を変更
  var lastRow = sheet.getLastRow();
  var rowRange = sheet.getRange(lastRow, 1, 1, 7);

  if (params.severity === '重大') {
    rowRange.setBackground('#fff5f5');
  } else if (params.severity === '要注意') {
    rowRange.setBackground('#fffff0');
  }

  return createJsonResponse({ success: true, id: id });
}

/**
 * 報告を削除
 */
function deleteReport(targetId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('報告データ');
  if (!sheet) {
    return createJsonResponse({ success: false, error: 'シートが見つかりません' });
  }

  var lastRow = sheet.getLastRow();
  for (var i = lastRow; i >= 2; i--) {
    var cellValue = String(sheet.getRange(i, 1).getValue());
    if (cellValue === String(targetId)) {
      sheet.deleteRow(i);
      return createJsonResponse({ success: true });
    }
  }

  return createJsonResponse({ success: false, error: '該当する報告が見つかりません' });
}

// ===== ユーティリティ =====

/**
 * JSON レスポンスを作成（CORS 対応）
 */
function createJsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 日付をJSON用にフォーマット
 */
function formatDateForJson(value) {
  if (!value) return '';
  if (value instanceof Date) {
    // YYYY-MM-DDTHH:mm 形式
    var y = value.getFullYear();
    var m = String(value.getMonth() + 1).padStart(2, '0');
    var d = String(value.getDate()).padStart(2, '0');
    var h = String(value.getHours()).padStart(2, '0');
    var min = String(value.getMinutes()).padStart(2, '0');
    return y + '-' + m + '-' + d + 'T' + h + ':' + min;
  }
  return String(value);
}
