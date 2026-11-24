/**
 * @OnlyCurrentDoc
 * 범용 백엔드 (Google Apps Script) — 어떤 시트 구조에도 그대로 사용
 *
 * 설치: 구글 시트 → 확장 프로그램 → Apps Script → 이 코드 붙여넣기 → 저장
 *       → 배포 → 새 배포 → 웹 앱 (실행: 나 / 액세스: 모든 사용자) → 웹 앱 URL 복사
 * 수정 후: 배포 → 배포 관리 → 연필 → 버전 '새 버전' → 배포 (URL 유지)
 *
 * 요청 방식 (GET·POST 모두 가능, POST는 Content-Type: text/plain 으로 JSON 전송 권장)
 *   action : add | list | get | update | delete | stats | tabs | schema
 *   tab    : 시트 탭 이름 (없으면 자동 생성, 기본 'data')
 *   data   : 저장할 값 {"이름":"홍길동","점수":90}  (새 키는 열로 자동 추가)
 *   id     : 행 번호 대신 쓰는 고유 ID (add 결과로 받음)
 *   token  : add 결과로 받은 수정·삭제 열쇠 (작성자 본인만 수정·삭제)
 *   key    : 교사용 관리자 열쇠 (비공개 열 조회, 전체 수정·삭제)
 *   where  : 조건 {"반":"3"}   search: 전체 열 검색어
 *   sort / desc(true) / limit / offset : 정렬·개수
 *   field / groupBy : stats에서 집계할 열 / 묶을 열
 *
 * 비공개 규칙: 열 이름이 _ 로 시작하면(예: _비밀번호, _메모) 관리자 열쇠 없이는 조회되지 않음.
 *             탭 이름이 _ 로 시작해도 관리자만 접근.
 */

var CONFIG = {
  ADMIN_KEY: '',          // 교사용 열쇠(비워 두면 관리자 기능 꺼짐). 스크립트 속성 ADMIN_KEY 가 있으면 그것을 우선 사용
  ALLOW_TABS: [],         // 허용할 탭 목록. 비우면 어떤 탭이든 자동 생성
  DEFAULT_TAB: 'data',
  MAX_ROWS: 5000,         // 탭당 최대 행 (도배 방지)
  MAX_COLS: 40,           // 탭당 최대 열
  MAX_CELL: 2000,         // 칸당 최대 글자 수
  TZ: 'Asia/Seoul'
};

var RESERVED_ = ['action', 'tab', 'key', 'token', 'id', 'data', 'where', 'search', 'sort', 'desc',
                 'limit', 'offset', 'field', 'groupBy', 'callback'];
var FIXED_ = ['id', 'createdAt', '_token'];   // 사용자가 직접 바꿀 수 없는 열

/* ───────── 진입점 ───────── */
function doGet(e) { return handle_(e, false); }
function doPost(e) { return handle_(e, true); }

function handle_(e, isPost) {
  var lock = null;
  try {
    var req = parse_(e);
    var action = String(req.action || (isPost ? 'add' : 'list')).toLowerCase();
    var writes = (action === 'add' || action === 'update' || action === 'delete');
    if (writes) { lock = LockService.getScriptLock(); lock.waitLock(10000); }

    var admin = isAdmin_(req.key);
    var tab = tabName_(req.tab, admin);
    var fn = ACTIONS_[action];
    if (!fn) throw new Error('알 수 없는 action: ' + action);
    return out_(fn(req, tab, admin));
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  } finally {
    if (lock) lock.releaseLock();
  }
}

var ACTIONS_ = {
  add: add_, list: list_, get: get_, update: update_, remove: remove_, 'delete': remove_,
  stats: stats_, tabs: tabs_, schema: schema_
};

/* ───────── 동작들 ───────── */
function add_(req, tab, admin) {
  var data = dataOf_(req);
  var keys = Object.keys(data);
  if (!keys.length) throw new Error('저장할 값(data)이 없어요');

  var sh = sheet_(tab, true);
  if (sh.getLastRow() - 1 >= CONFIG.MAX_ROWS) throw new Error('이 탭은 가득 찼어요 (최대 ' + CONFIG.MAX_ROWS + '행)');

  var headers = ensureCols_(sh, keys);
  var id = Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  var token = Utilities.getUuid().replace(/-/g, '');
  var obj = {};
  keys.forEach(function (k) { obj[k] = data[k]; });
  obj.id = id;
  obj.createdAt = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss');
  obj._token = token;

  sh.appendRow(headers.map(function (h) { return Object.prototype.hasOwnProperty.call(obj, h) ? cell_(obj[h]) : ''; }));
  return { ok: true, id: id, token: token, row: pub_(obj, admin) };
}

function list_(req, tab, admin) {
  var rows = readAll_(sheet_(tab, false));
  rows = filter_(rows, req);
  var sort = req.sort;
  if (sort) {
    var desc = truthy_(req.desc);
    rows.sort(function (a, b) { return cmp_(a[sort], b[sort]) * (desc ? -1 : 1); });
  } else if (truthy_(req.desc)) {
    rows.reverse();
  }
  var total = rows.length;
  var off = Math.max(0, parseInt(req.offset, 10) || 0);
  var lim = Math.min(1000, parseInt(req.limit, 10) || 100);
  rows = rows.slice(off, off + lim).map(function (r) { return pub_(r, admin); });
  return { ok: true, total: total, count: rows.length, rows: rows };
}

function get_(req, tab, admin) {
  var r = find_(readAll_(sheet_(tab, false)), req.id);
  if (!r) throw new Error('해당 id가 없어요');
  return { ok: true, row: pub_(r, admin) };
}

function update_(req, tab, admin) {
  var sh = sheet_(tab, false);
  var r = find_(readAll_(sh), req.id);
  if (!r) throw new Error('해당 id가 없어요');
  checkOwner_(r, req, admin);

  var data = dataOf_(req);
  var keys = Object.keys(data);
  if (!keys.length) throw new Error('바꿀 값(data)이 없어요');
  var headers = ensureCols_(sh, keys);
  var cur = sh.getRange(r.__row, 1, 1, headers.length).getValues()[0];
  headers.forEach(function (h, i) {
    if (Object.prototype.hasOwnProperty.call(data, h)) cur[i] = cell_(data[h]);
  });
  sh.getRange(r.__row, 1, 1, headers.length).setValues([cur]);
  var out = {};
  headers.forEach(function (h, i) { out[h] = cur[i]; });
  return { ok: true, row: pub_(out, admin) };
}

function remove_(req, tab, admin) {
  var sh = sheet_(tab, false);
  var r = find_(readAll_(sh), req.id);
  if (!r) throw new Error('해당 id가 없어요');
  checkOwner_(r, req, admin);
  sh.deleteRow(r.__row);
  return { ok: true, id: req.id };
}

// 집계: field(열) 하나의 개수·합·평균·최소·최대, groupBy 가 있으면 묶어서 계산
function stats_(req, tab, admin) {
  var rows = filter_(readAll_(sheet_(tab, false)), req);
  var field = req.field, by = req.groupBy;
  if ((field && isPrivate_(field) && !admin) || (by && isPrivate_(by) && !admin)) throw new Error('비공개 열이에요');

  function sum_(list) {
    var o = { count: list.length };
    if (!field) return o;
    var nums = [], counts = {};
    list.forEach(function (r) {
      var v = r[field];
      if (v === '' || v === undefined) return;
      counts[v] = (counts[v] || 0) + 1;
      if (isNum_(v)) nums.push(Number(v));
    });
    o.values = counts;
    if (nums.length) {
      var s = nums.reduce(function (a, b) { return a + b; }, 0);
      o.sum = s; o.avg = Math.round(s / nums.length * 100) / 100;
      o.min = Math.min.apply(null, nums); o.max = Math.max.apply(null, nums);
    }
    return o;
  }

  var res = { ok: true, total: rows.length, all: sum_(rows) };
  if (by) {
    var groups = {};
    rows.forEach(function (r) { var g = r[by] === '' ? '(빈 값)' : r[by]; (groups[g] = groups[g] || []).push(r); });
    res.groups = {};
    Object.keys(groups).forEach(function (g) { res.groups[g] = sum_(groups[g]); });
  }
  return res;
}

function tabs_(req, tab, admin) {
  var names = ss_().getSheets().map(function (s) { return s.getName(); })
    .filter(function (n) { return admin || n.charAt(0) !== '_'; });
  return { ok: true, tabs: names };
}

function schema_(req, tab, admin) {
  var sh = sheet_(tab, false);
  if (!sh) return { ok: true, tab: tab, columns: [], rows: 0 };
  var h = headers_(sh).filter(function (c) { return c && (admin ? c !== '_token' : !isPrivate_(c)); });
  return { ok: true, tab: tab, columns: h, rows: Math.max(0, sh.getLastRow() - 1) };
}

/* ───────── 도우미 ───────── */
function parse_(e) {
  e = e || {};
  var req = {}, k;
  var p = e.parameter || {};
  for (k in p) req[k] = p[k];
  if (e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      if (body && typeof body === 'object') for (k in body) req[k] = body[k];
    } catch (_) { /* 일반 폼 전송이면 e.parameter 로 충분 */ }
  }
  return req;
}

// data 객체 또는 (폼 전송처럼) 예약어가 아닌 모든 항목을 저장할 값으로 사용
function dataOf_(req) {
  var src = req.data;
  if (typeof src === 'string') { try { src = JSON.parse(src); } catch (_) { src = null; } }
  if (!src || typeof src !== 'object') {
    src = {};
    for (var k in req) if (RESERVED_.indexOf(k) < 0 && k.indexOf('q_') !== 0) src[k] = req[k];
  }
  var out = {};
  Object.keys(src).forEach(function (k) {
    var name = String(k).trim().slice(0, 40);
    if (!name || FIXED_.indexOf(name) >= 0) return;
    out[name] = src[k];
  });
  return out;
}

function cell_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number' || typeof v === 'boolean') return v;
  if (typeof v === 'object') v = JSON.stringify(v);
  v = String(v).slice(0, CONFIG.MAX_CELL);
  return /^[=+\-@]/.test(v) ? "'" + v : v;   // 수식 주입 방지
}

function isAdmin_(key) {
  var real = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY') || CONFIG.ADMIN_KEY;
  return !!real && String(key || '') === real;
}

function tabName_(t, admin) {
  t = String(t || CONFIG.DEFAULT_TAB).trim();
  if (!/^[0-9A-Za-z가-힣_\- ]{1,30}$/.test(t)) throw new Error('탭 이름은 30자 이내 글자·숫자·_·- 만 가능해요');
  if (t.charAt(0) === '_' && !admin) throw new Error('비공개 탭이에요');
  if (CONFIG.ALLOW_TABS.length && CONFIG.ALLOW_TABS.indexOf(t) < 0 && !admin) throw new Error('허용되지 않은 탭이에요');
  return t;
}

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(tab, create) {
  var s = ss_().getSheetByName(tab);
  if (!s && create) {
    s = ss_().insertSheet(tab);
    s.getRange(1, 1, 1, 3).setValues([FIXED_]);
    s.setFrozenRows(1);
  }
  return s;
}

function headers_(sh) {
  var n = sh.getLastColumn();
  return n ? sh.getRange(1, 1, 1, n).getValues()[0].map(String) : [];
}

// 없는 열은 오른쪽 끝에 추가하고 전체 머리글 반환
function ensureCols_(sh, keys) {
  var h = headers_(sh);
  var add = keys.filter(function (k) { return h.indexOf(k) < 0; });
  if (!add.length) return h;
  if (h.length + add.length > CONFIG.MAX_COLS) throw new Error('열이 너무 많아요 (최대 ' + CONFIG.MAX_COLS + '개)');
  var need = h.length + add.length;
  if (need > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), need - sh.getMaxColumns());
  sh.getRange(1, h.length + 1, 1, add.length).setValues([add]);
  return h.concat(add);
}

function readAll_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getDataRange().getValues();
  var h = v[0].map(String), rows = [];
  for (var i = 1; i < v.length; i++) {
    var o = { __row: i + 1 }, empty = true;
    for (var j = 0; j < h.length; j++) {
      if (!h[j]) continue;
      var x = v[i][j];
      if (x instanceof Date) x = Utilities.formatDate(x, CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss');
      if (x !== '') empty = false;
      o[h[j]] = x;
    }
    if (empty) continue;
    if (!o.id) o.id = 'r' + (i + 1);   // 교사가 시트에 직접 적은 행은 r+행번호 로 구분
    rows.push(o);
  }
  return rows;
}

function find_(rows, id) {
  id = String(id || '');
  if (!id) throw new Error('id가 필요해요');
  for (var i = 0; i < rows.length; i++) if (String(rows[i].id) === id) return rows[i];
  return null;
}

function checkOwner_(r, req, admin) {
  if (admin) return;
  if (!r._token || String(req.token || '') !== String(r._token)) throw new Error('수정·삭제 권한이 없어요 (token 필요)');
}

function filter_(rows, req) {
  var where = req.where;
  if (typeof where === 'string') { try { where = JSON.parse(where); } catch (_) { where = null; } }
  where = where && typeof where === 'object' ? where : {};
  for (var k in req) if (k.indexOf('q_') === 0) where[k.slice(2)] = req[k];   // ?q_반=3 형태도 허용
  var search = req.search ? String(req.search).toLowerCase() : '';
  var keys = Object.keys(where);
  return rows.filter(function (r) {
    for (var i = 0; i < keys.length; i++) {
      if (isPrivate_(keys[i])) return false;   // 비공개 열로 값 알아내기 방지
      if (String(r[keys[i]]) !== String(where[keys[i]])) return false;
    }
    if (search) {
      var hit = false;
      for (var c in r) if (c !== '__row' && !isPrivate_(c) && String(r[c]).toLowerCase().indexOf(search) >= 0) { hit = true; break; }
      if (!hit) return false;
    }
    return true;
  });
}

// 응답에서 내부 값 제거: __row·_token 은 항상, _열은 관리자만
function pub_(r, admin) {
  var o = {};
  for (var k in r) {
    if (k === '__row' || k === '_token') continue;
    if (isPrivate_(k) && !admin) continue;
    o[k] = r[k];
  }
  return o;
}

function isPrivate_(k) { return String(k).charAt(0) === '_'; }
function isNum_(v) { return v !== '' && v !== null && typeof v !== 'boolean' && !isNaN(Number(v)); }
function truthy_(v) { return v === true || v === 'true' || v === '1' || v === 1; }
function cmp_(a, b) {
  if (isNum_(a) && isNum_(b)) return Number(a) - Number(b);
  a = String(a === undefined ? '' : a); b = String(b === undefined ? '' : b);
  return a < b ? -1 : a > b ? 1 : 0;
}
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ───────── 연습용: 편집기에서 ▶ 실행하면 test 탭에 한 줄 저장 ───────── */
function selfTest() {
  var r = add_({ data: { 이름: '테스트', 점수: 90, _메모: '비공개' } }, 'test', false);
  Logger.log(JSON.stringify(r));
  Logger.log(JSON.stringify(stats_({ field: '점수' }, 'test', false)));
}
