const SPREADSHEET_ID = '1vdhgjk6lzSUeP1O-CC73LPAFSpLJlT7E3nKx7At7JMA';
const SHEET_NAME = '工作坊報名資料';

const HEADERS = [
  '提交時間',
  '姓名',
  '機構名',
  '職稱',
  '負責職類',
  '是否擔任教學訓練計畫主持人',
  '預計參與方式',
  'Email',
  '聯繫電話',
  '來源頁面',
];

function doGet() {
  const sheet = getRegistrationSheet_();

  return createJsonResponse_({
    ok: true,
    message: '教學訓練計畫主持人工作坊報名 API 已啟用',
    sheetName: sheet.getName(),
  });
}

function doPost(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};
    const sheet = getRegistrationSheet_();
    const row = [
      new Date(),
      sanitize_(params.name),
      sanitize_(params.organization),
      sanitize_(params.jobTitle),
      sanitize_(params.profession),
      sanitize_(params.isHost),
      sanitize_(params.participationType),
      sanitize_(params.email),
      sanitize_(params.phone),
      sanitize_(params.sourcePage),
    ];

    validateRequired_(row);
    sheet.appendRow(row);

    return createJsonResponse_({
      ok: true,
      message: '報名資料已寫入 Google Sheets',
    });
  } catch (error) {
    return createJsonResponse_({
      ok: false,
      message: error.message,
    });
  }
}

function getRegistrationSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
  ensureHeaders_(sheet);

  return sheet;
}

function ensureHeaders_(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  const currentHeaders = headerRange.getValues()[0];
  const hasHeaders = currentHeaders.some((value) => String(value).trim() !== '');

  if (!hasHeaders) {
    headerRange.setValues([HEADERS]);
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, HEADERS.length);
  }
}

function validateRequired_(row) {
  const requiredColumnIndexes = [1, 2, 3, 4, 5, 6, 7, 8];
  const missing = requiredColumnIndexes.filter((index) => !row[index]);

  if (missing.length > 0) {
    throw new Error('缺少必填欄位，請確認表單資料完整。');
  }
}

function sanitize_(value) {
  return String(value || '').trim();
}

function createJsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
