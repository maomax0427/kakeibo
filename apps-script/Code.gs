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
      case 'quickAdd': {
        const r = quickAdd_(req);
        if (req.format === 'text') return ContentService.createTextOutput(r.message);
        return json_(r);
      }
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

/* ---------- 共有メニューから追加（iPhoneのショートカット用） ----------
 * PayPayの支払い画面のスクリーンショットなどから読み取った文字(text)を受け取り、
 * 金額・お店・日付を取り出して支出として記録する。
 * req = {action:'quickAdd', text, category?(カテゴリ名), source?('PayPay'), format?('text')}
 */
const SKIP_WORDS_ = ['支払い', '支払', '完了', '残高', 'PayPay', 'ペイペイ', 'ポイント', '決済', '番号', '取引', '詳細', '日時', '金額',
  '合計', '獲得', '付与', '予定', '確認', 'ありがとう', '閉じる', '戻る', '共有', '明細', '利用', 'クーポン', '特典', '還元', '送る', '受け取る',
  'お支払い', 'ご利用', 'マネー', 'ライト', 'クレジット', 'カード', '詳しく', 'もっと', 'すべて', '履歴', 'ホーム', 'アカウント', 'スキャン'];

function parsePayText_(text, today) {
  const raw = String(text || '').replace(/\r/g, '');
  const norm = raw.replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0)).replace(/[，]/g, ',').replace(/[￥]/g, '¥');
  const lines = norm.split('\n').map(s => s.trim()).filter(Boolean);

  // 金額: 「¥1,234」「1,234円」。ポイント・残高の行は除き、支払い金額らしい行を優先する
  const cands = [];
  lines.forEach((ln, i) => {
    const re = /(?:¥\s*([\d,]+))|(?:([\d,]+)\s*円)/g; let m;
    while ((m = re.exec(ln))) {
      const v = Number((m[1] || m[2]).replace(/,/g, ''));
      if (!v || v > 10000000) continue;
      const ctx = ln + ' ' + (lines[i - 1] || '');
      let score = 0;
      if (/残高|ポイント|獲得|付与|還元|P\b|上限|あと/.test(ctx)) score -= 5;
      if (/支払|お支払い|合計|金額|決済/.test(ctx)) score += 3;
      if (/^[-−]?\s*(¥\s*)?[\d,]+\s*円?$/.test(ln.replace(/\s/g, ''))) score += 2; // 金額だけの大きな行
      cands.push({v, score, i});
    }
  });
  cands.sort((a, b) => b.score - a.score || a.i - b.i);
  const amount = cands.length ? cands[0].v : 0;

  // 日付: 2026年11月1日 / 2026/11/01 / 11月1日 / 11/01 12:34
  let date = today, m;
  if ((m = norm.match(/(20\d\d)\s*[年\/.-]\s*(\d{1,2})\s*[月\/.-]\s*(\d{1,2})/))) date = `${m[1]}-${pad2_(m[2])}-${pad2_(m[3])}`;
  else if ((m = norm.match(/(?:^|[^\d])(\d{1,2})\s*[月\/]\s*(\d{1,2})\s*日?\s*(?:\(|（)?/))) {
    const y = Number(today.slice(0, 4)); let d = `${y}-${pad2_(m[1])}-${pad2_(m[2])}`;
    if (d > today) d = `${y - 1}-${pad2_(m[1])}-${pad2_(m[2])}`;
    if (Number(m[1]) >= 1 && Number(m[1]) <= 12 && Number(m[2]) >= 1 && Number(m[2]) <= 31) date = d;
  }

  // お店: 「支払い先」「お店」の次の行、なければ金額・日付・決まり文句以外で最初の行
  let shop = '';
  const labelIdx = lines.findIndex(l => /^(支払い先|お支払い先|店舗|お店|加盟店|支払先)/.test(l));
  if (labelIdx >= 0) shop = lines[labelIdx].replace(/^(支払い先|お支払い先|店舗|お店|加盟店|支払先)[:：\s]*/, '') || lines[labelIdx + 1] || '';
  if (!shop) {
    shop = lines.find(l => l.length >= 2 && l.length <= 30 &&
      !/[¥円]|\d{1,2}:\d{2}|\d{4}[\/年]|\d{1,2}\s*[月\/]\s*\d{1,2}|^\d[\d,\s]*$|^[-−+]/.test(l) &&
      !SKIP_WORDS_.some(w => l.includes(w)) && !/^[\x00-\x2f\x3a-\x40]+$/.test(l)) || '';
  }
  return {amount, date, shop: shop.slice(0, 40)};
}
function pad2_(n) { return String(n).padStart(2, '0'); }

function guessCategory_(shop, config, txs) {
  const cats = (config && config.categories) || [];
  const exp = cats.filter(c => c.type === 'expense' && !c.archived);
  if (shop) {
    const same = txs.filter(t => t.type === 'expense' && t.memo && (t.memo === shop || shop.includes(t.memo) || t.memo.includes(shop)))
      .sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0))[0];
    if (same && exp.some(c => c.id === same.categoryId)) return same.categoryId;
  }
  const rules = [
    ['conbini', /セブン|ファミリーマート|ファミマ|ローソン|ミニストップ|デイリー|ニューデイズ|NewDays|コンビニ/i],
    ['food', /マクドナルド|マック|スタバ|スターバックス|吉野家|すき家|松屋|なか卯|ラーメン|らーめん|カフェ|珈琲|コーヒー|寿司|すし|バーガー|食堂|ゴンチャ|ドトール|タリーズ|サイゼ|ガスト|丸亀|日高屋|餃子|弁当|ベーカリー|パン|レストラン|居酒屋|焼肉|定食|そば|うどん|カレー/i],
    ['parking', /パーキング|駐車|タイムズ|三井のリパーク|リパーク/i],
    ['medical', /薬局|ドラッグ|病院|クリニック|医院|マツモトキヨシ|ウエルシア|スギ薬局|ツルハ/i],
  ];
  for (const [id, re] of rules) if (re.test(shop) && exp.some(c => c.id === id)) return id;
  return exp.some(c => c.id === 'uncat') ? 'uncat' : (exp[0] && exp[0].id);
}

function quickAdd_(req) {
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const p = parsePayText_(req.text, today);
  if (!p.amount) return {ok: false, message: '金額を読み取れませんでした。支払い金額が写っている画面で、もう一度お試しください。'};
  const data = load_();
  const config = data.config || {categories: []};
  // 「未分類」カテゴリがなければ作る（あとでアプリで直せるように）
  if (!(config.categories || []).some(c => c.id === 'uncat')) {
    config.categories = [...(config.categories || []), {id: 'uncat', name: '未分類', type: 'expense', group: 'variable', color: '#78909C', budget: 0}];
    cfgSheet_().getRange('A2').setValue(JSON.stringify(config));
  }
  let categoryId = null;
  if (req.category) { const c = config.categories.find(c => c.name === String(req.category).trim() && c.type === 'expense'); if (c) categoryId = c.id; }
  if (!categoryId) categoryId = guessCategory_(p.shop, config, data.transactions);
  const source = req.source || 'PayPay';
  const memo = p.shop || source;
  // 同じ画面を2回共有しても二重にならないよう、内容からIDを作る
  const key = [p.date, p.amount, memo].join('|');
  const id = 'qa_' + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, key).map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('').slice(0, 12);
  const dup = data.transactions.some(t => t.id === id);
  upsert_([{id, date: p.date, type: 'expense', amount: p.amount, categoryId, memo, createdAt: Date.now()}]);
  const cat = config.categories.find(c => c.id === categoryId);
  const md = `${Number(p.date.slice(5, 7))}/${Number(p.date.slice(8))}`;
  return {ok: true, id, ...p, categoryId,
    message: `${dup ? '（記録済みを更新）' : '記録しました'}\n${md} ${memo} ¥${p.amount.toLocaleString('ja-JP')}（${cat ? cat.name : '未分類'}）`};
}
