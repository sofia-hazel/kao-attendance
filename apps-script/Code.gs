// KAO 관피타 출첵앱 전용 서버 (Google Apps Script)
// 새 Google 시트 > 확장 프로그램 > Apps Script 에 이 코드를 통째로 붙여넣고 웹 앱으로 배포한다.
//
// 앱(index.html)과 주고받는 방식은 예전 서버와 같다.
//   읽기: GET  ?action=get&key=kao-state            → {ok:true, value:"<JSON 문자열>"}
//   쓰기: POST {"app":"kao","action":"set","key":"kao-state","value":"<JSON 문자열>"} → {ok:true}

const DATA_SHEET = 'data';      // A열 key, B열 value(JSON)
const BACKUP_SHEET = 'backup';  // 저장할 때마다 한 줄씩: 시각, key, value
const TABLE_SHEET = '출석표';    // 사람이 읽는 표 (앱이 저장할 때마다 다시 그린다)
const BACKUP_KEEP = 1000;       // 백업은 최근 1000줄만 남긴다
const CACHE_SEC = 21600;        // 캐시 최대 6시간 (저장할 때마다 새 값으로 바뀐다)

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// key가 있는 행 번호(없으면 0)와 값을 찾는다.
function find_(sh, key) {
  const last = sh.getLastRow();
  if (!last) return { row: 0, value: null };
  const vals = sh.getRange(1, 1, last, 2).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (vals[i][0] === key) return { row: i + 1, value: String(vals[i][1]) };
  }
  return { row: 0, value: null };
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action !== 'get' || !p.key) return json_({ ok: false, message: 'bad request' });
  // 읽기는 캐시에서 먼저 꺼낸다 — 시트를 열지 않아서 빠르다.
  const cache = CacheService.getScriptCache();
  const hit = cache.get('v:' + p.key);
  if (hit !== null) return json_({ ok: true, value: hit });
  const v = find_(sheet_(DATA_SHEET), p.key).value;
  if (v) cache.put('v:' + p.key, v, CACHE_SEC);
  return json_({ ok: true, value: v });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, message: 'bad json' }); }
  if (body.action !== 'set' || !body.key || typeof body.value !== 'string') return json_({ ok: false, message: 'bad request' });
  // 깨진 값이 저장돼 전체 기록이 망가지는 것을 막는다.
  try { JSON.parse(body.value); } catch (err) { return json_({ ok: false, message: 'value is not json' }); }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return json_({ ok: false, message: 'busy, retry' });
  try {
    const sh = sheet_(DATA_SHEET);
    const f = find_(sh, body.key);
    const row = f.row || sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, 2).setValues([[body.key, body.value]]);
    CacheService.getScriptCache().put('v:' + body.key, body.value, CACHE_SEC);

    const bk = sheet_(BACKUP_SHEET);
    bk.appendRow([new Date(), body.key, body.value]);
    const extra = bk.getLastRow() - BACKUP_KEEP;
    if (extra > 0) bk.deleteRows(1, extra);

    // 표 그리기가 실패해도 저장 자체는 성공으로 둔다 — 표는 어디까지나 보기용이다.
    if (body.key === 'kao-state') {
      try { writeTable_(JSON.parse(body.value)); } catch (err) {}
    }
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// ===== 사람이 읽는 표 =====
// 앱의 출첵표 열쇠는 연습이면 '2026-11-25', 행사면 '2026-11-25#보바스병원 연주' 형태다.
// B열 '구분'에 연습/행사이름을 적어 같은 날 두 줄이 나란히 보이게 한다.
function writeTable_(state) {
  const sh = sheet_(TABLE_SHEET);
  const members = state.members || [];
  const att = state.attendance || {};
  const labels = state.labels || {};
  const keys = (state.dates || []).slice().sort();

  sh.clear();
  const header = ['날짜', '구분'].concat(members);
  sh.getRange(1, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground('#1D9E75').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  if (!keys.length) return;

  const rows = keys.map(function (k) {
    const hash = k.indexOf('#');
    const day = hash < 0 ? k : k.slice(0, hash);
    const evt = hash < 0 ? '' : k.slice(hash + 1);
    const kind = evt || labels[k] || '연습';
    const a = att[k] || {};
    return [day, kind].concat(members.map(function (m) {
      return a[m] === 'o' ? '참여' : a[m] === 'x' ? '불참' : '-';
    }));
  });
  sh.getRange(2, 1, rows.length, header.length).setValues(rows);

  for (let r = 0; r < rows.length; r++) {
    // 행사 줄은 파랗게 칠해 연습 줄과 한눈에 구분되게 한다.
    const isEvt = keys[r].indexOf('#') >= 0;
    sh.getRange(r + 2, 2).setFontColor(isEvt ? '#185FA5' : '#888888');
    for (let c = 0; c < members.length; c++) {
      const v = rows[r][c + 2];
      const cell = sh.getRange(r + 2, c + 3);
      if (v === '참여') cell.setBackground('#E1F5EE').setFontColor('#0F6E56');
      else if (v === '불참') cell.setBackground('#FCEBEB').setFontColor('#A32D2D');
      else cell.setBackground('#ffffff').setFontColor('#bbbbbb');
    }
  }
  sh.getRange(2, 1, rows.length, 2).setHorizontalAlignment('left');
  sh.setColumnWidth(1, 95);
  sh.setColumnWidth(2, 130);
  for (let c = 3; c <= header.length; c++) sh.setColumnWidth(c, 70);
}
