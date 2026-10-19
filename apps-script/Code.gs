// KAO 관피타 출첵앱 전용 서버 (Google Apps Script)
// 새 Google 시트 > 확장 프로그램 > Apps Script 에 이 코드를 통째로 붙여넣고 웹 앱으로 배포한다.
//
// 앱(index.html)과 주고받는 방식은 예전 서버와 같다.
//   읽기: GET  ?action=get&key=kao-state            → {ok:true, value:"<JSON 문자열>"}
//   쓰기: POST {"app":"kao","action":"set","key":"kao-state","value":"<JSON 문자열>"} → {ok:true}

const DATA_SHEET = 'data';      // A열 key, B열 value(JSON)
const BACKUP_SHEET = 'backup';  // 저장할 때마다 한 줄씩: 시각, key, value
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
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
