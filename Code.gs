/**
 * GREF 주간 콘텐츠 소재 보드 — 저장소(구글 시트) 연결 스크립트
 *
 * 구글 시트 > 확장 프로그램 > Apps Script 를 열고
 * 기본으로 들어있는 코드를 지운 뒤 이 파일 전체를 붙여넣으세요.
 * 그다음 [배포] > [새 배포] > 유형 "웹 앱"으로 배포하면
 * 나오는 /exec 로 끝나는 주소를 index.html 의 CONFIG.ENDPOINT 에 넣으면 끝입니다.
 */

var SHEET_NAME = '제출';
var HEADERS = ['id', 'week', 'dept', 'author', 'body', 'createdAt', 'client'];

/* ------------------------------------------------------------------ */

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.setColumnWidth(5, 480);
  }
  return sh;
}

function toIso_(v) {
  if (v instanceof Date) return v.toISOString();
  return String(v || '');
}

function readAll_() {
  var sh = sheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (!String(r[0])) continue;
    out.push({
      row: i + 2,
      id: String(r[0]),
      week: String(r[1]),
      dept: String(r[2]),
      author: String(r[3]),
      body: String(r[4]),
      createdAt: toIso_(r[5]),
      client: String(r[6])
    });
  }
  return out;
}

function out_(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ */

function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = p.callback;
  try {
    var week = String(p.week || '');
    var list = [];
    var rows = readAll_();
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (week && r.week !== week) continue;
      list.push({
        id: r.id,
        dept: r.dept,
        author: r.author,
        body: r.body,
        createdAt: r.createdAt,
        client: r.client
      });
    }
    list.sort(function (a, b) { return a.createdAt < b.createdAt ? -1 : (a.createdAt > b.createdAt ? 1 : 0); });
    return out_({ ok: true, week: week, entries: list }, cb);
  } catch (err) {
    return out_({ ok: false, error: String(err) }, cb);
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var raw = (e && e.postData && e.postData.contents) || '{}';
    var d = JSON.parse(raw);
    if (d.action === 'delete') return remove_(d);
    return add_(d);
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function add_(d) {
  var body = String(d.body || '').trim().slice(0, 4000);
  var week = String(d.week || '').slice(0, 12);
  var dept = String(d.dept || '').slice(0, 20);
  if (!body || !week || !dept) return out_({ ok: false, error: 'missing field' });

  var id = 'e' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  sheet_().appendRow([
    id,
    week,
    dept,
    String(d.author || '').slice(0, 40),
    body,
    new Date().toISOString(),
    String(d.client || '').slice(0, 40)
  ]);
  return out_({ ok: true, id: id });
}

function remove_(d) {
  var id = String(d.id || '');
  var client = String(d.client || '');
  if (!id || !client) return out_({ ok: false, error: 'missing field' });
  var rows = readAll_();
  for (var i = rows.length - 1; i >= 0; i--) {
    if (rows[i].id === id && rows[i].client === client) {
      sheet_().deleteRow(rows[i].row);
      return out_({ ok: true });
    }
  }
  return out_({ ok: false, error: 'not found' });
}

/* ------------------------------------------------------------------ *
 * (선택) 시트 메뉴에서 눌러 이번 주 취합 내용을 텍스트로 확인
 * ------------------------------------------------------------------ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('GREF 보드')
    .addItem('이번 주 취합 내용 보기', 'showThisWeek')
    .addToUi();
}

function weekKeyOf_(date) {
  var d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  d.setDate(d.getDate() + 3 - ((d.getDay() + 6) % 7));
  var jan4 = new Date(d.getFullYear(), 0, 4);
  var wk = 1 + Math.round(((d - jan4) / 86400000 - 3 + ((jan4.getDay() + 6) % 7)) / 7);
  return d.getFullYear() + '-W' + (wk < 10 ? '0' + wk : wk);
}

function showThisWeek() {
  var wk = weekKeyOf_(new Date());
  var names = { grow: '재배', seedling: '육묘', hygiene: '위생 관리', facility: '설비 관리', ai: 'AI' };
  var order = ['grow', 'seedling', 'hygiene', 'facility', 'ai'];
  var rows = readAll_();
  var lines = ['[GREF 주간 콘텐츠 소재] ' + wk, ''];
  for (var i = 0; i < order.length; i++) {
    var key = order[i];
    lines.push('■ ' + names[key]);
    var any = false;
    for (var j = 0; j < rows.length; j++) {
      if (rows[j].week === wk && rows[j].dept === key) {
        lines.push('  · ' + (rows[j].author || '익명') + ': ' + rows[j].body);
        any = true;
      }
    }
    if (!any) lines.push('  (미제출)');
    lines.push('');
  }
  SpreadsheetApp.getUi().alert(lines.join('\n'));
}
