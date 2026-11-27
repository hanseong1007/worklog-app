/**
 * 경덕농수산 업무일지 — 실시간 전자결재 (Google Sheets + Apps Script)
 *
 * 연결 (한 번 배포한 웹앱 주소를 계속 씁니다)
 *  1) 구글 시트를 열고 메뉴 [확장 프로그램] → [Apps Script]를 엽니다.
 *  2) 이 코드를 전부 붙여 넣고 저장합니다.
 *  3) [배포] → [배포 관리] → 연필(수정) → 버전 '새 버전' → [배포].
 *     주소(.../exec)는 그대로 유지됩니다.
 *
 * 결재 3단계
 *  · 담당(작성자) — 앱에서 상신할 때 자동으로 승인 처리됩니다(상신일시 = 담당 결재일시).
 *  · 본부장      — 결재함에서 승인 또는 반려.
 *  · 대표        — 본부장 승인 뒤에 승인 또는 반려.
 *  한 단계라도 반려되면 그 결재는 반려로 끝납니다.
 *
 * 시트 열 구성 (기존 14개 열은 그대로 두고 뒤에 6개 열을 덧붙입니다)
 *  A 결재번호 | B 상신일시 | C 일자 | D 소속 | E 성명 | F 근무시간 | G 업무 | H 반품
 *  I 법인카드합계 | J 기타 | K 상태(전체) | L 결재자(마지막) | M 결재일시 | N 반려사유
 *  O 본부장상태 | P 본부장결재자 | Q 본부장결재일시 | R 대표상태 | S 대표결재자 | T 대표결재일시
 *
 * F 근무시간은 출퇴근 시각이 함께 올라오면 '08:30~17:30 · 8시간'처럼 저장합니다.
 * (앱은 이 칸에서 시각을 꺼내 출근·퇴근 칸을 되살립니다. 시각이 없던 옛 줄은 그대로 둡니다.)
 *
 * 업무·반품·기타는 한 칸에 줄바꿈으로 저장합니다. 업무는 '업무명: 업무내용' 한 줄씩 들어갑니다.
 * (예전 줄처럼 업무명 아래에 '- 항목' 이 붙은 모양도 앱은 그대로 읽습니다.)
 *
 * 공용 계정 목록 (같은 스프레드시트의 '계정목록' 시트)
 *  · users_get  — portal_users 행의 JSON 문자열을 읽어 { ok, usersJson, updatedAt } 로 돌려줍니다.
 *  · users_save — req.usersJson (이미 PBKDF2/SHA-256 으로 해시된 계정 목록 JSON, 평문 비밀번호 없음)을
 *                 portal_users 행에 덮어쓰기(upsert)하고 { ok, updatedAt } 를 돌려줍니다.
 *  · 열 구성: A key | B json | C updatedAt
 *
 * 자주 쓰는 업무명·업무내용 칩 (같은 '계정목록' 시트, key = 'worklog_taskchips')
 *  · chips_get  — { ok, byUser: { [userId]: { chips:[{name,items[]}], updatedAt } } }
 *                 req.userId 가 있으면 그 계정 것만 돌려줍니다.
 *  · chips_save — req.userId 의 칩만 갈아 끼우고 다른 계정의 칩은 그대로 둡니다.
 *                 (portal_users 와 다른 key 라서 계정 목록 저장이 칩을 덮어쓰지 않습니다)
 *  · 저장 형식: { byUser: { userId: { chips:[...], updatedAt:'yyyy-MM-dd HH:mm:ss' } } }
 */
var SHEET_NAME = '업무일지결재';
var USERS_SHEET_NAME = '계정목록';
var USERS_HEADERS = ['key', 'json', 'updatedAt'];
var USERS_KEY = 'portal_users';
var CHIPS_KEY = 'worklog_taskchips';
/* 구글 시트 한 칸은 50,000자까지. 넘치면 저장이 조용히 깨지므로 미리 막는다. */
var USERS_MAX_LEN = 45000;
var HEADERS = [
  '결재번호', '상신일시', '일자', '소속', '성명', '근무시간', '업무', '반품', '법인카드합계', '기타',
  '상태', '결재자', '결재일시', '반려사유',
  '본부장상태', '본부장결재자', '본부장결재일시', '대표상태', '대표결재자', '대표결재일시'
];
/* 단계별 열 번호 (1부터) */
var STAGE_COL = {
  chief: { st: 15, by: 16, at: 17, name: '본부장' },
  ceo:   { st: 18, by: 19, at: 20, name: '대표' }
};
var WAIT = '결재대기';
var OK = '승인';
var NO = '반려';

function doPost(e) {
  var req = {};
  try {
    req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    req = {};
  }
  var action = String(req.action || 'list');
  try {
    if (action === 'submit') return jsonOut(submit_(req));
    if (action === 'decide') return jsonOut(decide_(req));
    if (action === 'list') return jsonOut({ ok: true, rows: list_() });
    if (action === 'users_get') return jsonOut(usersGet_());
    if (action === 'users_save') return jsonOut(usersSave_(req));
    if (action === 'chips_get') return jsonOut(chipsGet_(req));
    if (action === 'chips_save') return jsonOut(chipsSave_(req));
    return jsonOut({ ok: false, error: '알 수 없는 요청: ' + action });
  } catch (err) {
    return jsonOut({ ok: false, error: String((err && err.message) || err) });
  }
}

function doGet(e) {
  try {
    /* 주소 뒤에 ?action=users_get 을 붙이면 계정 목록만 읽는다 (쓰기 없음) */
    if (e && e.parameter && e.parameter.action === 'users_get') return jsonOut(usersGet_());
    if (e && e.parameter && e.parameter.action === 'chips_get') {
      return jsonOut(chipsGet_({ userId: txt_(e.parameter.userId) }));
    }
    return jsonOut({ ok: true, rows: list_() });
  } catch (err) {
    return jsonOut({ ok: false, error: String((err && err.message) || err) });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function now_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

/* 계정 목록 저장 시각 — 같은 분 안의 두 번 저장도 순서를 가릴 수 있게 초까지 */
function nowSec_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function txt_(v) {
  return String(v == null ? '' : v).replace(/\r/g, '');
}

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    styleHead_(sh, 1, HEADERS.length);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 150);
    sh.setColumnWidth(7, 320);
    sh.setColumnWidth(8, 220);
    sh.setColumnWidth(10, 260);
    return sh;
  }
  /* 기존 시트(14열)에 본부장·대표 열만 덧붙인다. 지난 데이터는 건드리지 않는다. */
  var width = sh.getLastColumn();
  if (width < HEADERS.length) {
    var from = width + 1;
    var add = HEADERS.slice(width);
    sh.getRange(1, from, 1, add.length).setValues([add]);
    styleHead_(sh, from, add.length);
    sh.setColumnWidth(15, 110);
    sh.setColumnWidth(18, 110);
  }
  return sh;
}

function styleHead_(sh, col, count) {
  sh.getRange(1, col, 1, count).setFontWeight('bold').setBackground('#e7f4ec');
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    throw new Error('다른 결재가 처리 중입니다 · 잠시 뒤 다시 시도해 주세요');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/* 업무 한 칸: 업무명 아래에 하위항목을 '- 항목' 줄로 */
function taskLines_(t) {
  t = t || {};
  var name = String(t.name || '').trim();
  var items = [];
  if (Array.isArray(t.items)) {
    t.items.forEach(function (x) { var s = String(x == null ? '' : x).trim(); if (s) items.push(s); });
  } else if (t.content) {
    String(t.content).split(/\s*\/\s*|\n/).forEach(function (x) { var s = String(x).trim(); if (s) items.push(s); });
  }
  var lines = [];
  if (name) lines.push(name);
  items.forEach(function (x) { lines.push('- ' + x); });
  return lines;
}

/* 근무시간 칸 — 출근·퇴근 시각이 올라오면 '08:30~17:30 · 8시간'처럼 함께 남긴다.
   앱은 이 칸에서 다시 시각을 꺼내 출근시각·퇴근시각 칸을 채운다. */
function hoursCell_(req, doc) {
  var dur = txt_(req.hours);
  var st = txt_(doc.start || req.start).trim();
  var en = txt_(doc.end || req.end).trim();
  if (!st || !en) return dur;
  return st + '~' + en + (dur ? ' · ' + dur : '');
}

function submit_(req) {
  return withLock_(function () {
    var sh = sheet_();
    var doc = req.doc || {};
    var at = txt_(req.at) || now_();
    var hours = hoursCell_(req, doc);
    var id = 'A' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMddHHmmss') +
      String(Math.floor(Math.random() * 90 + 10));
    var tasks = [];
    (doc.tasks || []).forEach(function (t) {
      var lines = taskLines_(t);
      if (lines.length) tasks.push(lines.join('\n'));
    });
    var returns = [];
    (doc.returns || []).forEach(function (t) {
      var s = [(t.item || ''), (t.action || '')].filter(function (x) { return !!x; }).join(' — ');
      if (s) returns.push(s);
    });
    var cards = 0;
    (doc.cards || []).forEach(function (c) {
      cards += Number(String(c.amount || '').replace(/[^0-9]/g, '') || 0);
    });
    var author = txt_(doc.name) || txt_(req.author);
    sh.appendRow([
      id, at, txt_(doc.date) || txt_(req.date), txt_(doc.dept), author, hours,
      tasks.join('\n'), returns.join('\n'), cards, txt_(doc.notes),
      WAIT, '', '', '',
      WAIT, '', '',
      WAIT, '', ''
    ]);
    sh.getRange(sh.getLastRow(), 7, 1, 4).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    return {
      ok: true, id: id,
      row: {
        id: id, at: at, date: txt_(doc.date) || txt_(req.date), dept: txt_(doc.dept),
        author: author, hours: hours, tasks: tasks.join('\n'),
        returns: returns.join('\n'), cards: cards, notes: txt_(doc.notes),
        status: 'pending', by: '', decidedAt: '', reason: '',
        stages: { chief: { st: WAIT, by: '', at: '' }, ceo: { st: WAIT, by: '', at: '' } }
      }
    };
  });
}

function list_() {
  var sh = sheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, HEADERS.length).getDisplayValues();
  var rows = values.map(function (r) { return rowObj_(r); });
  return rows.reverse().slice(0, 60);
}

/* 시트 한 줄 → 앱이 읽는 모양 (14열짜리 옛 줄도 깨지지 않게) */
function stEn_(v) {
  v = txt_(v);
  return v === OK ? 'approved' : (v === NO ? 'rejected' : '');
}

function rowObj_(r) {
  var overall = txt_(r[10]);
  var status = overall === OK ? 'approved' : (overall === NO ? 'rejected' : 'pending');
  var by = txt_(r[11]);
  var at = txt_(r[12]);
  var reason = txt_(r[13]);
  var chief = { st: txt_(r[14]), by: txt_(r[15]), at: txt_(r[16]) };
  var ceo = { st: txt_(r[17]), by: txt_(r[18]), at: txt_(r[19]) };
  if (!chief.st && !ceo.st) {
    /* 옛 시트: 상태 한 칸만 있던 시절의 줄 — 그 결정을 두 단계에 그대로 옮겨 보여 준다 */
    if (status === 'approved') {
      chief = { st: OK, by: by, at: at };
      ceo = { st: OK, by: by, at: at };
    } else if (status === 'rejected') {
      chief = { st: NO, by: by, at: at };
      ceo = { st: '', by: '', at: '' };
    } else {
      chief = { st: WAIT, by: '', at: '' };
      ceo = { st: WAIT, by: '', at: '' };
    }
  }
  return {
    id: txt_(r[0]), at: txt_(r[1]), date: txt_(r[2]), dept: txt_(r[3]), author: txt_(r[4]),
    hours: txt_(r[5]), tasks: txt_(r[6]), returns: txt_(r[7]), cards: txt_(r[8]), notes: txt_(r[9]),
    status: status, by: by, decidedAt: at, reason: reason,
    stages: {
      chief: { st: stEn_(chief.st), by: chief.by, at: chief.at, name: STAGE_COL.chief.name, reason: chief.st === NO ? reason : '' },
      ceo: { st: stEn_(ceo.st), by: ceo.by, at: ceo.at, name: STAGE_COL.ceo.name, reason: ceo.st === NO ? reason : '' }
    }
  };
}

function decide_(req) {
  return withLock_(function () {
    var sh = sheet_();
    var last = sh.getLastRow();
    if (last < 2) return { ok: false, error: '상신 내역이 없습니다' };
    var role = String(req.role || 'chief');
    if (role !== 'chief' && role !== 'ceo') return { ok: false, error: '결재 역할이 올바르지 않습니다' };
    var col = STAGE_COL[role];
    var ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) !== String(req.id)) continue;
      var row = i + 2;
      var cur = sh.getRange(row, 1, 1, HEADERS.length).getDisplayValues()[0];
      var overall = txt_(cur[10]);
      if (overall === NO) return { ok: false, error: '이미 반려된 결재입니다' };
      if (overall === OK) return { ok: false, error: '이미 모든 결재가 끝난 결재입니다' };
      if (txt_(cur[col.st - 1]) === OK || txt_(cur[col.st - 1]) === NO) {
        return { ok: false, error: '이미 ' + col.name + ' 결재가 끝났습니다' };
      }
      if (role === 'ceo' && txt_(cur[STAGE_COL.chief.st - 1]) !== OK) {
        return { ok: false, error: '본부장 결재 뒤에 대표 결재를 할 수 있습니다' };
      }
      var decision = req.decision === 'rejected' ? NO : OK;
      var by = txt_(req.by);
      var at = txt_(req.at) || now_();
      var reason = decision === NO ? txt_(req.reason).slice(0, 200) : '';
      sh.getRange(row, col.st).setValue(decision);
      sh.getRange(row, col.by).setValue(by);
      sh.getRange(row, col.at).setValue(at);

      var chiefSt = txt_(cur[STAGE_COL.chief.st - 1]);
      var ceoSt = txt_(cur[STAGE_COL.ceo.st - 1]);
      if (role === 'chief') chiefSt = decision;
      if (role === 'ceo') ceoSt = decision;

      var out;
      if (decision === NO) {
        out = NO;
        sh.getRange(row, 14).setValue(reason);
      } else if (chiefSt === OK && ceoSt === OK) {
        out = OK;
        sh.getRange(row, 14).setValue('');
      } else {
        out = WAIT;
      }
      sh.getRange(row, 11).setValue(out);
      sh.getRange(row, 12).setValue(by);
      sh.getRange(row, 13).setValue(at);
      return { ok: true, id: req.id, status: out, role: role, row: rowObj_(sh.getRange(row, 1, 1, HEADERS.length).getDisplayValues()[0]) };
    }
    return { ok: false, error: '결재번호를 찾지 못했습니다' };
  });
}

/* ============================================================
   공용 계정 목록 — '계정목록' 시트 (A key | B json | C updatedAt)
   앱이 만든 계정 목록(해시된 값만 들어 있는 JSON 문자열)을 시트에 그대로 보관해,
   다른 기기·다른 직원이 같은 계정 목록으로 로그인할 수 있게 한다.
   ============================================================ */
function usersSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(USERS_SHEET_NAME);
    sh.appendRow(USERS_HEADERS);
    styleHead_(sh, 1, USERS_HEADERS.length);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 140);
    sh.setColumnWidth(2, 520);
    sh.setColumnWidth(3, 150);
    return sh;
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(USERS_HEADERS);
    styleHead_(sh, 1, USERS_HEADERS.length);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* key 로 행 찾기 (없으면 0) */
function rowByKey_(sh, key) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var keys = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (var i = 0; i < keys.length; i++) {
    if (String(keys[i][0]).trim() === key) return i + 2;
  }
  return 0;
}

/* portal_users 행 번호 (없으면 0) */
function usersRow_(sh) {
  return rowByKey_(sh, USERS_KEY);
}

function usersGet_() {
  var sh = usersSheet_();
  var row = usersRow_(sh);
  if (!row) return { ok: true, usersJson: '', updatedAt: '' };
  return {
    ok: true,
    usersJson: txt_(sh.getRange(row, 2).getValue()),
    updatedAt: usersAt_(sh.getRange(row, 3).getValue())
  };
}

/* 저장 시각 칸을 시트가 날짜 값으로 바꿔 버릴 수 있어, 읽을 때는 늘 'yyyy-MM-dd HH:mm:ss' 글자로 맞춘다 */
function usersAt_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  return txt_(v);
}

function usersSave_(req) {
  var raw = req.usersJson;
  if (raw && typeof raw === 'object') raw = JSON.stringify(raw);
  raw = typeof raw === 'string' ? raw : '';
  if (!raw) return { ok: false, error: '저장할 계정 목록이 비어 있습니다' };
  if (raw.length > USERS_MAX_LEN) {
    return { ok: false, error: '계정 목록이 너무 큽니다(' + raw.length + '자 · 최대 ' + USERS_MAX_LEN + '자)' };
  }
  try {
    JSON.parse(raw);
  } catch (err) {
    return { ok: false, error: '계정 목록이 올바른 JSON 형식이 아닙니다' };
  }
  return withLock_(function () {
    var sh = usersSheet_();
    var at = nowSec_();
    var row = usersRow_(sh);
    if (!row) row = Math.max(sh.getLastRow(), 1) + 1;
    /* 저장 시각이 날짜 값으로 바뀌지 않도록 그 칸을 글자 서식으로 둔다 */
    sh.getRange(row, 3).setNumberFormat('@');
    sh.getRange(row, 1, 1, 2).setValues([[USERS_KEY, raw]]);
    sh.getRange(row, 3).setValue(at);
    sh.getRange(row, 2).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    return { ok: true, updatedAt: at, length: raw.length };
  });
}

/* ============================================================
   자주 쓰는 업무명·업무내용 칩 — '계정목록' 시트의 worklog_taskchips 행
   저장 형식: { byUser: { userId: { chips:[{name, items[]}], updatedAt } } }
   계정(portal_users) 행과 다른 key 를 쓰므로 계정 목록 저장과 서로 덮어쓰지 않는다.
   ============================================================ */
var CHIPS_MAX = 60;          /* 업무명 개수 */
var CHIPS_MAX_ITEMS = 30;    /* 업무명 하나에 붙는 업무내용 개수 */
var CHIPS_MAX_NAME = 20;     /* 업무명 길이 (앱과 같다) */
var CHIPS_MAX_ITEM = 80;     /* 업무내용 길이 */

/* 들어온 칩을 앱과 같은 모양으로 다듬는다 — 이름 없는 줄·빈 업무내용은 버린다 */
function chipsClean_(chips) {
  var out = [];
  var src = Array.isArray(chips) ? chips : [];
  for (var i = 0; i < src.length && out.length < CHIPS_MAX; i++) {
    var p = src[i] || {};
    var name = String(p.name == null ? '' : p.name).replace(/\s+/g, ' ').trim().slice(0, CHIPS_MAX_NAME);
    if (!name) continue;
    var items = [];
    var list = Array.isArray(p.items) ? p.items : [];
    for (var j = 0; j < list.length && items.length < CHIPS_MAX_ITEMS; j++) {
      var v = String(list[j] == null ? '' : list[j]).replace(/\s+/g, ' ').trim().slice(0, CHIPS_MAX_ITEM);
      if (v) items.push(v);
    }
    out.push({ name: name, items: items });
  }
  return out;
}

function chipsByUser_(sh) {
  var row = rowByKey_(sh, CHIPS_KEY);
  if (!row) return {};
  var raw = txt_(sh.getRange(row, 2).getValue());
  if (!raw) return {};
  var obj = null;
  try { obj = JSON.parse(raw); } catch (err) { obj = null; }
  return (obj && obj.byUser && typeof obj.byUser === 'object') ? obj.byUser : {};
}

function chipsGet_(req) {
  req = req || {};
  var sh = usersSheet_();
  var byUser = chipsByUser_(sh);
  var uid = txt_(req.userId).trim();
  if (uid) {
    var one = {};
    if (byUser[uid]) one[uid] = byUser[uid];
    return { ok: true, byUser: one, updatedAt: one[uid] ? txt_(one[uid].updatedAt) : '' };
  }
  return { ok: true, byUser: byUser };
}

function chipsSave_(req) {
  req = req || {};
  var uid = txt_(req.userId).trim();
  if (!uid) return { ok: false, error: '계정 아이디가 없습니다' };
  if (!Array.isArray(req.chips)) return { ok: false, error: '칩 목록이 올바르지 않습니다' };
  var chips = chipsClean_(req.chips);
  if (JSON.stringify(chips).length > 40000) {
    return { ok: false, error: '칩 목록이 너무 큽니다 · 업무명과 업무내용을 줄여 주세요' };
  }
  return withLock_(function () {
    var sh = usersSheet_();
    var at = nowSec_();
    var byUser = chipsByUser_(sh);            /* 다른 계정의 칩은 그대로 둔다 */
    byUser[uid] = { chips: chips, updatedAt: at };
    var out = JSON.stringify({ byUser: byUser });
    if (out.length > USERS_MAX_LEN) {
      return { ok: false, error: '시트에 저장할 공간이 부족합니다(' + out.length + '자)' };
    }
    var row = rowByKey_(sh, CHIPS_KEY);
    if (!row) row = Math.max(sh.getLastRow(), 1) + 1;
    sh.getRange(row, 3).setNumberFormat('@');  /* 저장 시각이 날짜 값으로 바뀌지 않게 */
    sh.getRange(row, 1, 1, 2).setValues([[CHIPS_KEY, out]]);
    sh.getRange(row, 3).setValue(at);
    sh.getRange(row, 2).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    return { ok: true, userId: uid, updatedAt: at, saved: chips.length };
  });
}
