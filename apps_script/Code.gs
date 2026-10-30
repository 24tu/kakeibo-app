/**
 * 家計簿アプリのデータ保存係（Googleスプレッドシートに貼り付けて使う）
 *
 * 使い方:
 *   1. スプレッドシートの「拡張機能」→「Apps Script」を開き、このコードを全部貼り付けて保存
 *   2. 上の関数選択で「setup」を選んで「実行」→ 許可を求められたら許可する
 *   3. 実行ログに出る「合言葉」をメモする（スマホのアプリに一度だけ入力する）
 *   4. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *        次のユーザーとして実行: 自分 / アクセスできるユーザー: 全員
 *      → 表示された「ウェブアプリのURL」をメモする
 *
 * 合言葉がないと、データの読み書きはできません。
 */

const SHEETS = {
  cash:   ["id", "date", "month", "shop", "amt", "cat", "memo", "createdAt"],
  cards:  ["month", "d", "shop", "amt", "cat", "card", "kind"],
  budget: ["cat", "amount"],
};

/** 最初に1回だけ実行する：シートを作って合言葉を決める */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const [name, header] of Object.entries(SHEETS)) {
    const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
    sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty("TOKEN");
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, "").slice(0, 16);
    props.setProperty("TOKEN", token);
  }
  Logger.log("合言葉: " + token);
}

/** アプリからの依頼を受け取る */
function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const token = PropertiesService.getScriptProperties().getProperty("TOKEN");
    if (!token || req.token !== token) return reply({ ok: false, error: "bad_token" });

    switch (req.action) {
      case "load":       return reply({ ok: true, ...loadAll() });
      case "addCash":    return reply(withLock(() => addCash(req.rec)));
      case "deleteCash": return reply(withLock(() => deleteCash(req.id)));
      case "setBudget":  return reply(withLock(() => setBudget(req.amounts)));
      case "setCards":   return reply(withLock(() => setCards(req.month, req.items)));
      default:           return reply({ ok: false, error: "unknown_action" });
    }
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  }
}

function doGet() {
  return reply({ ok: true, message: "家計簿アプリのデータ保存係は動いています" });
}

// ---------- 中身 ----------

function sheet(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }

function rows(name) {
  const values = sheet(name).getDataRange().getValues();
  const header = values.shift();
  return values.filter(r => r.some(v => v !== "")).map(r => {
    const o = {};
    header.forEach((h, i) => o[h] = r[i] instanceof Date ? fmtDate(r[i]) : r[i]);
    return o;
  });
}

function fmtDate(d) { return Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd"); }

function loadAll() {
  const budget = {};
  rows("budget").forEach(r => { if (Number(r.amount) > 0) budget[r.cat] = Number(r.amount); });
  return {
    cash: rows("cash").map(r => ({ ...r, amt: Number(r.amt), date: String(r.date), month: String(r.month) })),
    cards: rows("cards").map(r => ({ ...r, amt: Number(r.amt), month: String(r.month), d: String(r.d) })),
    budget,
  };
}

function addCash(rec) {
  const id = rec.id || Utilities.getUuid();
  // 同じ記録が2回届いても1件にする（電波が悪いときの再送対策）
  if (rows("cash").some(r => r.id === id)) return { ok: true, id };
  const date = String(rec.date);
  // 日付や月が自動で日付型に変わらないよう、先頭に ' を付けて文字として保存する
  sheet("cash").appendRow([id, "'" + date, "'" + date.slice(0, 7), rec.shop || "", Number(rec.amt),
    rec.cat || "その他", rec.memo || "", new Date().toISOString()]);
  return { ok: true, id };
}

function deleteCash(id) {
  const sh = sheet("cash");
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = ids.length - 1; i >= 1; i--) {
    if (ids[i][0] === id) sh.deleteRow(i + 1);
  }
  return { ok: true };
}

function setBudget(amounts) {
  const sh = sheet("budget");
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  const data = Object.entries(amounts || {}).filter(([, v]) => Number(v) > 0).map(([k, v]) => [k, Number(v)]);
  if (data.length) sh.getRange(2, 1, data.length, 2).setValues(data);
  return { ok: true };
}

/** その月のカード明細を入れ替える（PCの kakeibo.py から送られてくる） */
function setCards(month, items) {
  const sh = sheet("cards");
  const values = sh.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][0]) === month) sh.deleteRow(i + 1);
  }
  const data = (items || []).map(it => ["'" + month, "'" + (it.d || ""), it.shop, Number(it.amt), it.cat, it.card, it.kind]);
  if (data.length) sh.getRange(sh.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
  return { ok: true, count: data.length };
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
