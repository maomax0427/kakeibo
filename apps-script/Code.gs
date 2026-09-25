/**
 * わたしの家計簿 — データ保存用 Apps Script
 * スプレッドシート「家計簿データ」の 拡張機能 → Apps Script に貼り付けて、
 * デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *   次のユーザーとして実行：自分
 *   アクセスできるユーザー：全員
 * で公開してください。発行された URL がアプリの接続先になります（他人に教えないこと）。
 */
const TX_SHEET = 'transactions';
const CFG_SHEET = 'config';
const TX_HEAD = ['id', 'date', 'type', 'amount', 'categoryId', 'memo', 'createdAt', 'recurringId'];

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function txSheet_() {
  let sh = ss_().getSheetByName(TX_SHEET);
  if (!sh) {
    sh = ss_().insertSheet(TX_SHEET);
    sh.getRange(1, 1, 1, TX_HEAD.length).setValues([TX_HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:B').setNumberFormat('@');
    sh.getRange('F:F').setNumberFormat('@');
    sh.getRange('H:H').setNumberFormat('@');
  }
  return sh;
}

function cfgSheet_() {
  let sh = ss_().getSheetByName(CFG_SHEET);
  if (!sh) {
    sh = ss_().insertSheet(CFG_SHEET);
    sh.getRange('A1').setValue('config (JSON)').setFontWeight('bold');
    sh.getRange('A2').setNumberFormat('@');
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function dateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v);
}

function load_() {
  const sh = txSheet_();
  const values = sh.getDataRange().getValues().slice(1);
  const transactions = values.filter(r => r[0]).map(r => {
    const t = {
      id: String(r[0]), date: dateStr_(r[1]), type: String(r[2]), amount: Number(r[3]),
      categoryId: String(r[4]), memo: String(r[5] || ''), createdAt: Number(r[6]) || 0,
    };
    if (r[7]) t.recurringId = String(r[7]);
    return t;
  });
  const raw = cfgSheet_().getRange('A2').getValue();
  let config = null;
  try { config = raw ? JSON.parse(raw) : null; } catch (e) { config = null; }
  return { ok: true, config, transactions };
}

function upsert_(txs) {
  const sh = txSheet_();
  const ids = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues().map(r => String(r[0]));
  const appends = [];
  for (const t of txs) {
    const row = [t.id, t.date, t.type, Number(t.amount), t.categoryId, t.memo || '', t.createdAt || Date.now(), t.recurringId || ''];
    const i = ids.indexOf(String(t.id));
    if (i > 0) sh.getRange(i + 1, 1, 1, row.length).setValues([row]);
    else appends.push(row);
  }
  if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, TX_HEAD.length).setValues(appends);
}

function delete_(delIds) {
  const sh = txSheet_();
  const ids = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues().map(r => String(r[0]));
  const rows = [];
  ids.forEach((id, i) => { if (i > 0 && delIds.indexOf(id) >= 0) rows.push(i + 1); });
  rows.sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
}

function doGet() { return json_(load_()); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = JSON.parse(e.postData.contents || '{}');
    switch (req.action) {
      case 'load': return json_(load_());
      case 'saveConfig': cfgSheet_().getRange('A2').setValue(JSON.stringify(req.config)); break;
      case 'upsert': upsert_(req.txs || []); break;
      case 'delete': delete_((req.ids || []).map(String)); break;
      case 'replaceAll': {
        const sh = txSheet_();
        if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, TX_HEAD.length).clearContent();
        upsert_(req.txs || []);
        break;
      }
      default: return json_({ ok: false, error: 'unknown action' });
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** 最初に一度だけ実行（シートの準備と、権限の承認のため） */
function setup() {
  txSheet_();
  cfgSheet_();
  const s1 = ss_().getSheetByName('シート1') || ss_().getSheetByName('Sheet1');
  if (s1 && ss_().getSheets().length > 1) ss_().deleteSheet(s1);
}
