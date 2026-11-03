/**
 * 家計簿アプリのデータ保存係（Googleスプレッドシートに貼り付けて使う）  第3版
 *
 * この版から、アプリに新しい項目や種類のデータが増えても、
 * 列やシートを自動で追加するので、基本的にこのコードを貼り替える必要はありません。
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

// 最初から用意しておくシートと列（これ以外の列やシートは、必要になったら自動で増える）
const SHEETS = {
  tx:       ["id", "date", "month", "shop", "amt", "cat", "method", "memo", "type", "createdAt", "updatedAt"],  // 手入力の記録
  fixed:    ["id", "name", "day", "amt", "cat", "method", "start", "end", "memo", "type"],                    // 固定費・固定の収入
  cards:    ["stmt", "month", "d", "shop", "amt", "cat", "card", "via", "kind"],                              // PCから送られるカード明細
  budget:   ["cat", "amount"],
  settings: ["key", "value"],
};
// アプリから直接は書きかえさせないシート
const PROTECTED = ["cards", "budget", "settings"];

/** 最初に1回、またはコードを新しくしたときに実行する */
function setup() {
  for (const [name, header] of Object.entries(SHEETS)) ensureColumns(sheet(name), header);

  // 第1版の「cash」シートの記録を「tx」に移す
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const old = ss.getSheetByName("cash");
  if (old && old.getLastRow() > 1) {
    rowsOf(old).forEach(o => upsert("tx", { ...o, method: "現金", type: "expense" }));
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
      case "upsert":     return reply(withLock(() => { checkEditable(req.sheet); return { ok: true, id: upsert(req.sheet, req.rec) }; }));
      case "remove":     return reply(withLock(() => { checkEditable(req.sheet); removeRow(req.sheet, req.id); return { ok: true }; }));
      case "setBudget":  return reply(withLock(() => setBudget(req.amounts)));
      case "setMethods": return reply(withLock(() => setSetting("methods", req.methods || [])));
      case "setSetting": return reply(withLock(() => setSetting(String(req.key), req.value)));
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

// ---------- シートと列 ----------

/** シートを取り出す。なければ作る */
function sheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    ensureColumns(sh, SHEETS[name] || ["id"]);
  }
  return sh;
}

function headerOf(sh) {
  return sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String) : [];
}

/** 足りない列を右に追加して、列の並びを返す */
function ensureColumns(sh, keys) {
  const header = headerOf(sh);
  const missing = keys.filter(k => k && !header.includes(k));
  if (missing.length) {
    sh.getRange(1, header.length + 1, 1, missing.length).setValues([missing]).setFontWeight("bold");
    sh.setFrozenRows(1);
    header.push(...missing);
  }
  return header;
}

function checkEditable(name) {
  if (!/^[a-z][a-z0-9_]{0,30}$/.test(String(name)) || PROTECTED.includes(name)) throw new Error("bad_sheet");
}

function fmtDate(d) { return Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd"); }

/** 文字はそのまま文字として保存する（日付や数字に勝手に変わらないように先頭に ' を付ける） */
function cell(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v === "" ? "" : "'" + v;
  if (typeof v === "object") return "'" + JSON.stringify(v);
  return v;
}

function rowsOf(sh) {
  const values = sh.getDataRange().getValues();
  const header = values.shift().map(String);
  return values.filter(r => r.some(v => v !== "")).map(r => {
    const o = {};
    header.forEach((h, i) => { if (h) o[h] = r[i] instanceof Date ? fmtDate(r[i]) : r[i]; });
    return o;
  });
}
function rows(name) { return rowsOf(sheet(name)); }

// ---------- 読み書き ----------

function loadAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const budget = {};
  rows("budget").forEach(r => { if (Number(r.amount) > 0) budget[r.cat] = Number(r.amount); });
  const settings = {};
  rows("settings").forEach(r => { try { settings[r.key] = JSON.parse(r.value); } catch (e) { settings[r.key] = r.value; } });

  // アプリが書きこめるシートは全部まとめて返す（新しい種類のデータもそのまま届く）
  const sheets = {};
  ss.getSheets().forEach(sh => {
    const name = sh.getName();
    if (/^[a-z][a-z0-9_]{0,30}$/.test(name) && !PROTECTED.includes(name)) sheets[name] = rowsOf(sh);
  });
  return {
    tx: sheets.tx || [], fixed: sheets.fixed || [], cards: rows("cards"),
    budget, settings, sheets,
    methods: Array.isArray(settings.methods) ? settings.methods : null,
  };
}

/** id が同じ行があれば書きかえ、なければ追加する。知らない項目は列を増やして保存する */
function upsert(name, rec) {
  const sh = sheet(name);
  const id = String(rec.id || Utilities.getUuid());
  const now = new Date().toISOString();
  const data = { ...rec, id, updatedAt: now };
  const header = ensureColumns(sh, ["id", ...Object.keys(data), "createdAt"]);

  const ids = sh.getLastRow() > 1 ? sh.getRange(2, header.indexOf("id") + 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  const at = ids.indexOf(id);
  if (at >= 0) {
    const current = sh.getRange(at + 2, 1, 1, header.length).getValues()[0];
    const row = header.map((h, i) => h in data ? cell(data[h]) : current[i]);
    sh.getRange(at + 2, 1, 1, header.length).setValues([row]);
  } else {
    data.createdAt = data.createdAt || now;
    const row = header.map(h => cell(data[h]));
    sh.getRange(sh.getLastRow() + 1, 1, 1, header.length).setValues([row]);
  }
  return id;
}

function removeRow(name, id) {
  const sh = sheet(name);
  if (sh.getLastRow() < 2) return;
  const col = headerOf(sh).indexOf("id") + 1;
  if (!col) return;
  const ids = sh.getRange(1, col, sh.getLastRow(), 1).getValues();
  for (let i = ids.length - 1; i >= 1; i--) if (String(ids[i][0]) === String(id)) sh.deleteRow(i + 1);
}

function setBudget(amounts) {
  const sh = sheet("budget");
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  const data = Object.entries(amounts || {}).filter(([, v]) => Number(v) > 0).map(([k, v]) => [cell(k), Number(v)]);
  if (data.length) sh.getRange(2, 1, data.length, 2).setValues(data);
  return { ok: true };
}

/** 設定を1つ保存する（値は何でもよい。JSONとして保存する） */
function setSetting(key, value) {
  const sh = sheet("settings");
  const text = "'" + JSON.stringify(value);
  const keys = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  const at = keys.indexOf(key);
  if (at >= 0) sh.getRange(at + 2, 2).setValue(text);
  else sh.appendRow([cell(key), text]);
  return { ok: true };
}

/** 1枚の明細（例：PayPayカードの10月引き落とし分）を入れ替える。PCの kakeibo.py から送られてくる */
function setCards(stmt, items) {
  const sh = sheet("cards");
  const keys = new Set(["stmt"]);
  (items || []).forEach(it => Object.keys(it).forEach(k => keys.add(k)));
  const header = ensureColumns(sh, [...keys]);
  const col = header.indexOf("stmt") + 1;
  if (sh.getLastRow() > 1) {
    const values = sh.getRange(1, col, sh.getLastRow(), 1).getValues();
    for (let i = values.length - 1; i >= 1; i--) if (String(values[i][0]) === stmt) sh.deleteRow(i + 1);
  }
  const data = (items || []).map(it => header.map(h => cell(h === "stmt" ? stmt : it[h])));
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
