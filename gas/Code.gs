/**
 * 百名湯 制覇マップ — 入湯記録の読み書きAPI
 * 置き場所：入湯記録用スプレッドシートの「拡張機能 → Apps Script」
 * 初回だけ setup() を実行し、ログに出る合言葉を控える。
 */
const SHEET_NAME = '入湯記録';
const HEADER = ['記録ID', '温泉ID', '温泉名', '入湯年', 'メモ', '登録日時'];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) { sh.appendRow(HEADER); sh.setFrozenRows(1); }
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('API_KEY');
  if (!key) { key = Utilities.getUuid().replace(/-/g, '').slice(0, 16); props.setProperty('API_KEY', key); }
  Logger.log('合言葉: ' + key);
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const key = PropertiesService.getScriptProperties().getProperty('API_KEY');
    if (!key || req.key !== key) return out({ ok: false, error: '合言葉が違います' });
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const lock = LockService.getScriptLock(); lock.waitLock(10000);
    try {
      if (req.action === 'add') {
        const year = parseInt(req.year, 10);
        if (!(year >= 1950 && year <= 2100) || !/^[A-Z]\d{3}$/.test(req.onsenId)) return out({ ok: false, error: '入力値が不正です' });
        sh.appendRow([Utilities.getUuid(), req.onsenId, String(req.onsenName || '').slice(0, 50), year, String(req.memo || '').slice(0, 200), new Date()]);
      } else if (req.action === 'delete') {
        const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(r => r[0]);
        const i = ids.indexOf(req.id);
        if (i >= 0) sh.deleteRow(i + 2);
      }
      return out({ ok: true, visits: list(sh) });
    } finally { lock.releaseLock(); }
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function list(sh) {
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues()
    .filter(r => r[0])
    .map(r => ({ id: String(r[0]), onsenId: String(r[1]), onsenName: r[2], year: Number(r[3]), memo: r[4] }));
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
