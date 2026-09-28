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

// 点検写真をGoogle Driveへ保存し、プレビュー用URLを返す
function uploadImageToDrive(fileData) {
  try {
    if (!fileData || typeof fileData.bytes !== 'string' || !fileData.bytes) {
      throw new Error('画像データが空です。画像を選び直してください。');
    }
    if (!/^image\/(jpeg|png|gif|webp|heic|heif)$/i.test(fileData.mimeType || '')) {
      throw new Error('JPEG、PNG、GIF、WebP、HEIC形式の画像を選択してください。');
    }
    if (Math.ceil(fileData.bytes.length * 3 / 4) > 10 * 1024 * 1024) {
      throw new Error('画像は10MB以下にしてください。');
    }

    const safeName = String(fileData.filename || 'inspection-photo')
      .replace(/[^a-zA-Z0-9._-]/g, '_');
    const blob = Utilities.newBlob(
      Utilities.base64Decode(fileData.bytes),
      fileData.mimeType,
      safeName
    );
    const file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return {
      success: true,
      fileId: file.getId(),
      fileUrl: 'https://drive.google.com/uc?export=view&id=' + file.getId()
    };
  } catch (e) {
    return { success: false, error: e.message || String(e) };
  }
}

// 保存済み案件に登録された写真だけを、プレビュー用データとして返す
function getReportPhotoPreview(reportId, photoUrl) {
  try {
    const report = getReportDetail(reportId);
    if (!report) throw new Error('対象の報告書が見つかりません。');

    const savedUrls = [
      ...Object.values(report.photos || {}),
      ...(report.pcsData || []).map(pcs => pcs.photo || ''),
      ...(report.panelFarPhotos || [])
    ];
    if (!savedUrls.includes(photoUrl)) throw new Error('この報告書に登録されていない写真です。');

    const match = String(photoUrl).match(/(?:[?&]id=|\/d\/)([A-Za-z0-9_-]+)/);
    if (!match) throw new Error('写真URLの形式が正しくありません。');

    const blob = DriveApp.getFileById(match[1]).getBlob();
    const bytes = blob.getBytes();
    if (bytes.length > 10 * 1024 * 1024) {
      throw new Error('画像が大きいため一覧内に表示できません。画像をクリックしてDriveで開いてください。');
    }

    return {
      success: true,
      mimeType: blob.getContentType(),
      base64: Utilities.base64Encode(bytes)
    };
  } catch (e) {
    return { success: false, error: e.message || String(e) };
  }
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