/**
 * たいないどけいチャレンジ（忘年会）
 * スプレッドシートに紐づく Google Apps Script
 *
 * ── デプロイ手順 ──
 * 1. Googleスプレッドシートを新規作成する（名前例: たいないどけいチャレンジ）
 * 2. 拡張機能 → Apps Script を開き、この Code.gs の中身をすべて貼り付ける
 * 3. エディタ上部の関数で setup を選び「実行」。権限を許可する
 *    （実行ログにスプレッドシートのURLが出る）
 * 4. 右上「デプロイ」→「新しいデプロイ」
 *    - 種類: ウェブアプリ
 *    - 次のユーザーとして実行: 自分
 *    - アクセスできるユーザー: 全員   ←「全員（匿名）」相当。ログイン必須にしない
 * 5. 発行されたURL（末尾が /exec）をコピーする
 * 6. index.html 先頭付近の GAS_URL にそのURLを貼って保存し、GitHub Pages へ上げる
 *
 * コードを直したあとは「デプロイを管理」→ 鉛筆 → バージョン「新しいバージョン」で
 * 更新すること。同じ /exec URLのまま反映される。/dev URLは使わない（ログインが要る）。
 *
 * シート:
 *   players … 参加者
 *   config  … targetSec（10〜30）
 */

var PLAYER_HEADERS = [
  'id', 'name', 'charId', 'status', 'targetSec',
  'elapsedMs', 'errorMs', 'absErrorMs', 'createdAt', 'finishedAt', 'dummy'
];

function setup() {
  var ss = ss_();
  players_();
  config_();
  Logger.log(ss.getUrl());
  return ss.getUrl();
}

function doGet(e) {
  return handle_((e && e.parameter) || {});
}

function doPost(e) {
  var p = {};
  if (e && e.parameter) {
    Object.keys(e.parameter).forEach(function (k) { p[k] = e.parameter[k]; });
  }
  if (e && e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      Object.keys(body).forEach(function (k) { p[k] = body[k]; });
    } catch (err) {}
  }
  return handle_(p);
}

function handle_(p) {
  var action = String(p.action || 'list');
  var result;
  try {
    if (action === 'list' || action === 'config' || action === 'ping') {
      result = dispatch_(action, p);
    } else {
      var lock = LockService.getScriptLock();
      lock.waitLock(15000);
      try {
        result = dispatch_(action, p);
      } finally {
        lock.releaseLock();
      }
    }
  } catch (err) {
    result = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return emit_(result, p);
}

function dispatch_(action, p) {
  switch (action) {
    case 'ping': return { ok: true, targetSec: getTarget_() };
    case 'config': return { ok: true, targetSec: getTarget_() };
    case 'list': return { ok: true, targetSec: getTarget_(), players: readPlayers_() };
    case 'entry': return entry_(p);
    case 'cancel': return cancel_(p);
    case 'submit': return submit_(p);
    case 'setTarget': return setTarget_(p);
    case 'dummies': return makeDummies_();
    case 'reset': return resetAll_();
    default: return { ok: false, error: 'unknown action' };
  }
}

function emit_(obj, p) {
  var json = JSON.stringify(obj);
  var cb = p && p.callback ? String(p.callback) : '';
  if (cb && /^[A-Za-z0-9_$.]+$/.test(cb)) {
    return ContentService
      .createTextOutput(cb + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function ss_() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SSID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) {}
  }
  var created = SpreadsheetApp.create('たいないどけいチャレンジ');
  props.setProperty('SSID', created.getId());
  return created;
}

function players_() {
  var ss = ss_();
  var sh = ss.getSheetByName('players');
  if (!sh) {
    sh = ss.insertSheet('players');
    sh.appendRow(PLAYER_HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function config_() {
  var ss = ss_();
  var sh = ss.getSheetByName('config');
  if (!sh) {
    sh = ss.insertSheet('config');
    sh.appendRow(['key', 'value']);
    sh.appendRow(['targetSec', 15]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getTarget_() {
  var sh = config_();
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === 'targetSec') {
      var n = parseInt(vals[i][1], 10);
      if (n >= 10 && n <= 30) return n;
    }
  }
  sh.appendRow(['targetSec', 15]);
  return 15;
}

function setTarget_(p) {
  var n = parseInt(p.targetSec, 10);
  if (!(n >= 10 && n <= 30)) return { ok: false, error: 'target out of range' };
  var sh = config_();
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === 'targetSec') {
      sh.getRange(i + 1, 2).setValue(n);
      return { ok: true, targetSec: n };
    }
  }
  sh.appendRow(['targetSec', n]);
  return { ok: true, targetSec: n };
}

function cleanName_(raw) {
  var s = String(raw == null ? '' : raw).replace(/[\r\n\t]/g, '').trim();
  if (!s) return '';
  // スプレッドシート数式の注入を防ぐ
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  var chars = Array.from(s);
  if (chars.length > 16) s = chars.slice(0, 16).join('');
  return s;
}

function entry_(p) {
  var name = cleanName_(p.name);
  var charId = parseInt(p.charId, 10);
  if (!name) return { ok: false, error: 'name' };
  if (!(charId >= 0 && charId <= 19)) return { ok: false, error: 'char' };
  var target = getTarget_();
  var id = Utilities.getUuid();
  var now = Date.now();
  players_().appendRow([id, name, charId, 'ready', target, '', '', '', now, '', '']);
  return { ok: true, id: id, name: name, charId: charId, targetSec: target };
}

function cancel_(p) {
  var id = String(p.id || '');
  if (!id) return { ok: false, error: 'id' };
  var sh = players_();
  var found = findRow_(sh, id);
  if (!found) return { ok: true };
  if (String(found.row[3]) !== 'ready') return { ok: false, error: 'locked' };
  sh.deleteRow(found.index);
  return { ok: true };
}

function findRow_(sh, id) {
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(id)) return { index: i + 1, row: vals[i] };
  }
  return null;
}

function submit_(p) {
  var id = String(p.id || '');
  if (!id) return { ok: false, error: 'id' };
  var sh = players_();
  var found = findRow_(sh, id);
  if (!found) return { ok: false, error: 'not found' };
  var status = String(found.row[3] || '');
  if (status === 'done' || status === 'fled') {
    return { ok: true, duplicate: true, status: status };
  }
  var dq = String(p.dq) === '1' || p.dq === true || p.dq === 1;
  var now = Date.now();
  if (dq) {
    sh.getRange(found.index, 4, 1, 7).setValues([[
      'fled', found.row[4], '', '', '', found.row[8], now
    ]]);
    return { ok: true, status: 'fled' };
  }
  var elapsed = Math.round(Number(p.elapsedMs));
  if (!isFinite(elapsed) || elapsed < 0 || elapsed > 600000) {
    return { ok: false, error: 'bad time' };
  }
  var target = parseInt(p.targetSec, 10);
  if (!(target >= 10 && target <= 30)) target = parseInt(found.row[4], 10) || getTarget_();
  var errorMs = elapsed - target * 1000;
  var absError = Math.abs(errorMs);
  sh.getRange(found.index, 4, 1, 7).setValues([[
    'done', target, elapsed, errorMs, absError, found.row[8], now
  ]]);
  return { ok: true, status: 'done' };
}

function makeDummies_() {
  // 発表リハーサル用。誤差の幅が揃う固定オフセット（ミリ秒、遅いが正）
  var presets = [
    ['アルス', 0, 80],
    ['ガライ', 1, -420],
    ['マリナ', 2, 1600],
    ['セイル', 3, -90],
    ['カク', 4, 3500],
    ['パル', 5, -2200],
    ['ゴールド', 6, 640],
    ['キング', 7, 40],
    ['ピピ', 8, null],
    ['マオウ', 9, 4800]
  ];
  var target = getTarget_();
  var sh = players_();
  var now = Date.now();
  var rows = presets.map(function (d, i) {
    var offset = d[2];
    var done = offset !== null && offset !== undefined;
    var elapsed = done ? target * 1000 + offset : '';
    var errorMs = done ? offset : '';
    var absError = done ? Math.abs(offset) : '';
    return [
      Utilities.getUuid(), d[0], d[1], done ? 'done' : 'ready', target,
      elapsed, errorMs, absError, now + i, done ? now + i : '', '1'
    ];
  });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, PLAYER_HEADERS.length).setValues(rows);
  return { ok: true, added: rows.length, targetSec: target };
}

function resetAll_() {
  var sh = players_();
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, PLAYER_HEADERS.length).clearContent();
  return { ok: true };
}

function readPlayers_() {
  var sh = players_();
  var vals = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < vals.length; i++) {
    var r = vals[i];
    if (!r[0]) continue;
    out.push({
      id: String(r[0]),
      name: String(r[1]),
      charId: Number(r[2]),
      status: String(r[3] || 'ready'),
      targetSec: Number(r[4]) || getTarget_(),
      elapsedMs: r[5] === '' || r[5] === null ? null : Number(r[5]),
      errorMs: r[6] === '' || r[6] === null ? null : Number(r[6]),
      absErrorMs: r[7] === '' || r[7] === null ? null : Number(r[7]),
      createdAt: Number(r[8]) || 0,
      finishedAt: Number(r[9]) || 0,
      dummy: String(r[10]) === '1'
    });
  }
  out.sort(function (a, b) { return a.createdAt - b.createdAt; });
  return out;
}
