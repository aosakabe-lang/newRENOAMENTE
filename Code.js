// ==========================================
// 設定項目（※ご自身のスプレッドシートIDを指定）
// ==========================================
const SPREADSHEET_ID = '1pKouDDVqUeb6Y58ywBaa8Zj7diEoJGHjlURkuIJEto0'; 
const SHEET_NAME = '案件管理データ'; 

// 1. Webアプリの初期表示
function doGet() {
  const html = HtmlService.createTemplateFromFile('index').evaluate();
  html.setTitle('リノアメンテナンスツール');
  html.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return html;
}

// 2. データの保存処理 (下書き / 確定)
function saveReportData(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(["レポートID", "ステータス", "更新日時", "電圧区分", "施主様名", "発電所名", "データJSON"]);
    }

    const reportId = formData.reportId || Utilities.getUuid();
    formData.reportId = reportId; // IDを更新
    
    const rowData = [
      reportId,
      formData.status, // "下書き" または "確定"
      new Date(),
      formData.voltageType || '',
      formData.clientName || '',
      formData.plantName || '',
      JSON.stringify(formData)
    ];

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === reportId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex > -1) {
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    return { success: true, reportId: reportId };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// 3. 過去データ・下書きの一覧を取得
function getReportList() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const list = [];
    
    // ヘッダーを除いて逆順（新しい順）で取得
    for (let i = data.length - 1; i >= 1; i--) {
      list.push({
        reportId: data[i][0],
        status: data[i][1],
        updatedAt: Utilities.formatDate(new Date(data[i][2]), "JST", "yyyy/MM/dd HH:mm"),
        voltageType: data[i][3],
        clientName: data[i][4],
        plantName: data[i][5]
      });
    }
    return list;
  } catch (e) {
    throw new Error("一覧取得エラー: " + e.message);
  }
}

// 4. 指定したIDのデータ詳細（JSON）を取得して再復元用データとして返す
function getReportDetail(reportId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === reportId) {
        return JSON.parse(data[i][6]); // データJSONをオブジェクトに戻して返す
      }
    }
    return null;
  } catch (e) {
    throw new Error("データ取得エラー: " + e.message);
  }
}