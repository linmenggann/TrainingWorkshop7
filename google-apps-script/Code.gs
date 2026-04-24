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

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const sheet = getRegistrationSheet_();
  const payload = params.action === 'dashboard'
    ? buildDashboardPayload_(sheet)
    : {
        ok: true,
        message: '教學訓練計畫主持人工作坊報名 API 已啟用',
        sheetName: sheet.getName(),
      };

  if (params.callback) {
    return createJsonpResponse_(params.callback, payload);
  }

  return createJsonResponse_(payload);
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

function buildDashboardPayload_(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  const dataRows = values.length > 1 ? values.slice(1).filter((row) => row.some((value) => value !== '')) : [];
  const rows = dataRows.map((row) => ({
    submittedAt: formatDate_(row[0]),
    name: sanitize_(row[1]),
    organization: sanitize_(row[2]),
    jobTitle: sanitize_(row[3]),
    profession: sanitize_(row[4]),
    isHost: sanitize_(row[5]),
    participationType: sanitize_(row[6]),
    email: sanitize_(row[7]),
    phone: sanitize_(row[8]),
    sourcePage: sanitize_(row[9]),
  }));
  const summary = rows.reduce((acc, row) => {
    acc.total += 1;
    incrementCount_(acc.byProfession, row.profession || '未填寫');
    incrementCount_(acc.byParticipationType, row.participationType || '未填寫');
    incrementCount_(acc.byHostStatus, row.isHost || '未填寫');
    incrementCount_(acc.byOrganization, row.organization || '未填寫');
    return acc;
  }, {
    total: 0,
    byProfession: {},
    byParticipationType: {},
    byHostStatus: {},
    byOrganization: {},
  });

  return {
    ok: true,
    sheetName: sheet.getName(),
    lastUpdated: formatDate_(new Date()),
    summary,
    rows: rows.reverse(),
  };
}

function incrementCount_(target, key) {
  target[key] = (target[key] || 0) + 1;
}

function formatDate_(value) {
  if (!value) {
    return '';
  }

  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) {
    return sanitize_(value);
  }

  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
}

function createJsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function createJsonpResponse_(callback, payload) {
  const safeCallback = String(callback).replace(/[^\w.$]/g, '');
  return ContentService
    .createTextOutput(`${safeCallback}(${JSON.stringify(payload)});`)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
