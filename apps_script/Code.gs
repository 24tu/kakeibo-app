/**
 * 家計簿アプリのデータ保存係（Googleスプレッドシートに貼り付けて使う）  第3版
 *
 * はじめて使うとき:
 *   1. スプレッドシートの「拡張機能」→「Apps Script」を開き、このコードを全部貼り付けて保存
 *   2. 上の関数選択で「setup」を選んで「実行」→ 許可を求められたら許可する
 *   3. 実行ログに出る「合言葉」をメモする（スマホのアプリに一度だけ入力する）
 *   4. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *        次のユーザーとして実行: 自分 / アクセスできるユーザー: 全員
 *      → 表示された「ウェブアプリのURL」をメモする
 *
 * コードを新しくしたとき（URLは変わりません）:
 *   1. コードを全部貼り替えて保存 → 「setup」を実行
 *   2. 「デプロイ」→「デプロイを管理」→ 鉛筆マーク → バージョン「新バージョン」→「デプロイ」
 *
 * 合言葉がないと、データの読み書きはできません。
 */

const VERSION = 3;

const SHEETS = {
  // 手入力の記録（現金・PayPay・カードなど）。type は expense（支出）/ income（収入）/ nospend（支出なしの日）
  tx:       ["id", "date", "month", "shop", "amt", "cat", "method", "memo", "createdAt", "updatedAt", "type"],
  // 固定費（毎月◯日に◯円）。type は expense（支出）/ income（給料などの収入）
  fixed:    ["id", "name", "day", "amt", "cat", "method", "start", "end", "memo", "type"],
  // PCから送られるカード明細（stmt = どのカードの何月分の明細か）
  cards:    ["stmt", "month", "d", "shop", "amt", "cat", "card", "via", "kind"],
  budget:   ["cat", "amount"],
  settings: ["key", "value"],
};
// アプリから書きかえてよいシートと、文字のまま保存したい列（日付が自動変換されないように）
const EDITABLE = ["tx", "fixed"];
const TEXT_COLS = ["date", "month", "start", "end", "d", "stmt"];

/** 最初に1回、またはコードを新しくしたときに実行する */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const [name, header] of Object.entries(SHEETS)) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const current = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
    if (current.join() !== header.join()) {
      if (name === "cards") sh.clear();   // カード明細は形が変わったので空にする（PCから送り直す）
      sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight("bold");
    }
    sh.setFrozenRows(1);
  }

  // 第1版の「cash」シートの記録を「tx」に移す
  const old = ss.getSheetByName("cash");
  if (old && old.getLastRow() > 1) {
    const values = old.getDataRange().getValues();
    const head = values.shift();
    values.forEach(r => {
      const o = {}; head.forEach((h, i) => o[h] = r[i] instanceof Date ? fmtDate(r[i]) : r[i]);
      upsert("tx", { ...o, method: "現金" });
    });
    old.setName("cash_old");
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
      case "load":       return reply({ ok: true, version: VERSION, ...loadAll() });
      case "upsert":     return reply(withLock(() => { check(req.sheet); return { ok: true, id: upsert(req.sheet, req.rec) }; }));
      case "remove":     return reply(withLock(() => { check(req.sheet); removeRow(req.sheet, req.id); return { ok: true }; }));
      case "setBudget":  return reply(withLock(() => setBudget(req.amounts)));
      case "setMethods": return reply(withLock(() => setSetting("methods", JSON.stringify(req.methods || []))));
      case "setCards":   return reply(withLock(() => setCards(req.stmt, req.items)));
      default:           return reply({ ok: false, error: "unknown_action" });
    }
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  }
}

function doGet() {
  return reply({ ok: true, version: VERSION, message: "家計簿アプリのデータ保存係は動いています" });
}

// ---------- 中身 ----------

function sheet(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
function check(name) { if (!EDITABLE.includes(name)) throw new Error("bad_sheet"); }
function fmtDate(d) { return Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd"); }

function rows(name) {
  const values = sheet(name).getDataRange().getValues();
  const header = values.shift();
  return values.filter(r => r.some(v => v !== "")).map(r => {
    const o = {};
    header.forEach((h, i) => o[h] = r[i] instanceof Date ? fmtDate(r[i]) : r[i]);
    return o;
  });
}

function loadAll() {
  const budget = {};
  rows("budget").forEach(r => { if (Number(r.amount) > 0) budget[r.cat] = Number(r.amount); });
  const settings = {};
  rows("settings").forEach(r => settings[r.key] = r.value);
  let methods = null;
  try { methods = settings.methods ? JSON.parse(settings.methods) : null; } catch (e) {}
  const num = r => ({ ...r, amt: Number(r.amt) });
  return {
    tx: rows("tx").map(num),
    fixed: rows("fixed").map(r => ({ ...num(r), day: Number(r.day) })),
    cards: rows("cards").map(num),
    budget,
    methods,
  };
}

/** id が同じ行があれば書きかえ、なければ追加する */
function upsert(name, rec) {
  const sh = sheet(name);
  const header = SHEETS[name];
  const id = rec.id || Utilities.getUuid();
  const row = header.map(h => {
    if (h === "id") return id;
    if (h === "updatedAt") return new Date().toISOString();
    if (h === "createdAt") return rec.createdAt || new Date().toISOString();
    const v = rec[h] ?? "";
    return TEXT_COLS.includes(h) && v !== "" ? "'" + v : v;
  });
  const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
  const at = ids.indexOf(id);
  if (at >= 0) {
    if (header.includes("createdAt")) row[header.indexOf("createdAt")] = sh.getRange(at + 2, header.indexOf("createdAt") + 1).getValue() || row[header.indexOf("createdAt")];
    sh.getRange(at + 2, 1, 1, header.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
  return id;
}

function removeRow(name, id) {
  const sh = sheet(name);
  if (sh.getLastRow() < 2) return;
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = ids.length - 1; i >= 1; i--) if (ids[i][0] === id) sh.deleteRow(i + 1);
}

function setBudget(amounts) {
  const sh = sheet("budget");
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  const data = Object.entries(amounts || {}).filter(([, v]) => Number(v) > 0).map(([k, v]) => [k, Number(v)]);
  if (data.length) sh.getRange(2, 1, data.length, 2).setValues(data);
  return { ok: true };
}

function setSetting(key, value) {
  const sh = sheet("settings");
  const keys = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
  const at = keys.indexOf(key);
  if (at >= 0) sh.getRange(at + 2, 2).setValue(value);
  else sh.appendRow([key, value]);
  return { ok: true };
}

/** 1枚の明細（例：PayPayカードの10月引き落とし分）を入れ替える。PCの kakeibo.py から送られてくる */
function setCards(stmt, items) {
  const sh = sheet("cards");
  const values = sh.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) if (String(values[i][0]) === stmt) sh.deleteRow(i + 1);
  const header = SHEETS.cards;
  const data = (items || []).map(it => header.map(h => {
    const v = h === "stmt" ? stmt : (it[h] ?? "");
    return TEXT_COLS.includes(h) && v !== "" ? "'" + v : v;
  }));
  if (data.length) sh.getRange(sh.getLastRow() + 1, 1, data.length, header.length).setValues(data);
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
