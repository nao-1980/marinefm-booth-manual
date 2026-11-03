/**
 * マリンFM 案内放送ブース マニュアル：イベント版の保存先（Googleスプレッドシート）
 *
 * このスクリプトはスプレッドシートに紐づけて使う（拡張機能 → Apps Script）。
 * シートIDはコードに書かず、紐づいているスプレッドシートを使う。
 * 書き込み用のパスコードは「プロジェクトの設定 → スクリプト プロパティ」に
 * PASSCODE という名前で保存する（コードには書かない）。
 */

var SHEET_NAME = 'events';
var TEXT_KEYS = ['id', 'name', 'date', 'time', 'place', 'setting', 'org', 'hq',
  'shelters', 'direction', 'map', 'aid', 'lost', 'aed', 'shelterIn', 'schedule', 'memo', 'updatedAt'];
var BOOL_KEYS = ['sea', 'river', 'crowd', 'stage', 'alcohol', 'kids'];
var HEADERS = TEXT_KEYS.concat(BOOL_KEYS);
var MAX_LEN = 2000;
var MAX_EVENTS = 300;

// 一覧の取得（だれでも読める）
function doGet(e) {
  return json_({ ok: true, events: readAll_() });
}

// 追加・更新・削除（パスコードが必要）
function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  if (!checkPasscode_(req.passcode)) return json_({ ok: false, error: 'passcode' });

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return json_({ ok: false, error: 'busy' });
  try {
    if (req.action === 'verify') return json_({ ok: true });
    if (req.action === 'save') return json_(saveEvent_(req.event));
    if (req.action === 'delete') return json_(deleteEvent_(req.id));
    return json_({ ok: false, error: 'bad_request' });
  } finally {
    lock.releaseLock();
  }
}

// 初回だけ手動で実行：見出し行を作り、文字列として保存されるよう書式を設定する
function setup() {
  var sheet = sheet_();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  // 移行前からページに載っていたイベントを登録しておく
  if (rowIndexOf_(sheet, 'e-hotchpotch-20261115') < 0) {
    saveEvent_({ id: 'e-hotchpotch-20261115', name: 'ホッチポッチ横浜', date: '2026-11-15', setting: 'outdoor',
      place: '象の鼻パーク（案内放送ブース）', sea: true, kids: true,
      direction: '海の方向には行かない。避難先は象の鼻パークの津波避難情報板で確認',
      map: '横浜市 津波避難情報板（象の鼻パークに設置）',
      memo: '時間・本部・救護所・迷子の待機場所は、わかりしだい「編集」から追加してください。' });
  }
}

function checkPasscode_(input) {
  var pass = PropertiesService.getScriptProperties().getProperty('PASSCODE');
  if (!pass) return false; // 未設定なら書き込みは常に拒否
  return typeof input === 'string' && input === pass;
}

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  // 日付や時刻が自動変換されないよう、すべて書式なしテキストにする
  sheet.getRange(1, 1, sheet.getMaxRows(), HEADERS.length).setNumberFormat('@');
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  return sheet;
}

function readAll_() {
  var sheet = sheet_();
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var rows = sheet.getRange(2, 1, last - 1, head.length).getValues();
  return rows.filter(function (r) { return r[head.indexOf('id')]; }).map(function (r) {
    var ev = {};
    head.forEach(function (k, i) {
      if (BOOL_KEYS.indexOf(k) >= 0) ev[k] = r[i] === true || String(r[i]).toUpperCase() === 'TRUE';
      else if (TEXT_KEYS.indexOf(k) >= 0) ev[k] = r[i] instanceof Date ? Utilities.formatDate(r[i], 'Asia/Tokyo', 'yyyy-MM-dd') : String(r[i]);
    });
    return ev;
  });
}

function rowIndexOf_(sheet, id) {
  var last = sheet.getLastRow();
  if (last < 2) return -1;
  var ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return -1;
}

// セルの先頭が = + - @ だと数式として扱われるので、先頭に ' をつけて文字列にする
function cell_(v) {
  v = String(v == null ? '' : v).slice(0, MAX_LEN);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function saveEvent_(ev) {
  if (!ev || typeof ev !== 'object') return { ok: false, error: 'bad_request' };
  var id = String(ev.id || '');
  if (!/^[A-Za-z0-9_-]{1,64}$/.test(id)) return { ok: false, error: 'bad_request' };
  if (!String(ev.name || '').trim() || !/^\d{4}-\d{2}-\d{2}$/.test(String(ev.date || '')) || !String(ev.place || '').trim()) {
    return { ok: false, error: 'required' };
  }
  ev.setting = ev.setting === 'indoor' ? 'indoor' : 'outdoor';
  ev.updatedAt = new Date().toISOString();

  var row = HEADERS.map(function (k) {
    return BOOL_KEYS.indexOf(k) >= 0 ? (ev[k] ? 'TRUE' : '') : cell_(ev[k]);
  });
  var sheet = sheet_();
  var at = rowIndexOf_(sheet, id);
  if (at < 0) {
    if (sheet.getLastRow() - 1 >= MAX_EVENTS) return { ok: false, error: 'quota' };
    sheet.appendRow(row);
  } else {
    sheet.getRange(at, 1, 1, row.length).setValues([row]);
  }
  return { ok: true, id: id };
}

function deleteEvent_(id) {
  var sheet = sheet_();
  var at = rowIndexOf_(sheet, String(id || ''));
  if (at > 0) sheet.deleteRow(at);
  return { ok: true };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
