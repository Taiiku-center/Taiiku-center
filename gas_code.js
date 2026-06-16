// ============================================================
//  授業申込カレンダー — Google Apps Script バックエンド
//  すべてGETパラメータで処理（CORS回避）
// ============================================================

const SHEET_NAME = '申込データ';
const EXPECTED_HEADERS = ['ID', '生徒名', '学年', '日付', '曜日', '開始時刻', '終了時刻', '状態', '申請日時', '更新日時'];

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, EXPECTED_HEADERS.length, 120);
  }
  // ヘッダー行を常に正しい順序に強制修正
  const headerRange = sheet.getRange(1, 1, 1, EXPECTED_HEADERS.length);
  const current = headerRange.getValues()[0];
  if (EXPECTED_HEADERS.some((h, i) => String(current[i]) !== h)) {
    headerRange.setValues([EXPECTED_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, EXPECTED_HEADERS.length, 120);
  }
  return sheet;
}

function doGet(e) {
  const action = e.parameter.action;
  let result;

  try {
    if (action === 'getAll') {
      result = getAllData();
    } else if (action === 'getByStudent') {
      result = getByStudent(e.parameter.student);
    } else if (action === 'submit') {
      const slots = JSON.parse(decodeURIComponent(e.parameter.slots));
      result = submitSlots(slots);
    } else if (action === 'updateStatus') {
      result = updateStatus(e.parameter.id, e.parameter.status);
    } else if (action === 'delete') {
      result = deleteRow(e.parameter.id);
    } else if (action === 'fixSheet') {
      getSheet(); // ヘッダー修正のみ
      result = { ok: true, message: 'ヘッダーを修正しました' };
    } else {
      result = { error: 'unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }

  const cb = e.parameter.callback;
  if (cb) {
    return ContentService
      .createTextOutput(`${cb}(${JSON.stringify(result)})`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// 値を文字列に正規化（Dateオブジェクト対策）
function normalizeValue(h, v) {
  if (v instanceof Date) {
    if (h === '開始時刻' || h === '終了時刻') {
      return Utilities.formatDate(v, 'Asia/Tokyo', 'HH:mm');
    }
    if (h === '日付' || h === '申請日時' || h === '更新日時') {
      return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    }
    return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  }
  return String(v === null || v === undefined ? '' : v);
}

// 全データ取得
function getAllData() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { rows: [] };
  const headers = data[0];
  const rows = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, j) => obj[h] = normalizeValue(h, row[j]));
    return obj;
  });
  return { rows };
}

// 名前のスペースを除去して正規化
function normalizeName(s) {
  return String(s || '').replace(/[\s　]/g, '');
}

// 生徒別データ取得（スペースの有無に関わらず一致）
function getByStudent(student) {
  const all = getAllData();
  const key = normalizeName(student);
  return { rows: all.rows.filter(r => normalizeName(r['生徒名']) === key) };
}

// 申請（複数スロット）
function submitSlots(slots) {
  const sheet = getSheet();
  const now = new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' });
  const added = [];

  slots.forEach(slot => {
    const id = Utilities.getUuid();
    sheet.appendRow([
      id,
      normalizeName(slot.student),
      slot.grade || '',
      slot.date,
      slot.dow,
      slot.start,
      slot.end,
      '申請中',
      now,
      now
    ]);
    added.push(id);
  });

  return { ok: true, added };
}

// ステータス更新
function updateStatus(id, status) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' });

  const headers = data[0];
  const statusCol = headers.indexOf('状態') + 1;
  const updatedCol = headers.indexOf('更新日時') + 1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i + 1, statusCol).setValue(status);
      sheet.getRange(i + 1, updatedCol).setValue(now);
      return { ok: true };
    }
  }
  return { error: 'not found' };
}

// 行削除
function deleteRow(id) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { error: 'not found' };
}
