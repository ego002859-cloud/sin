/* 신부장 콘텐츠 스튜디오 — 구글 시트 저장소 API
 *
 * 대시보드(index.html) 여러 명이 같은 시트를 보고 쓴다.
 * 시트는 사람이 직접 열어 읽을 수 있게 한 행 = 한 레코드로 저장한다.
 *
 * ── 설치 ──
 * 1. 구글 시트 새로 만들기 → 확장 프로그램 → Apps Script
 * 2. 이 파일 붙여넣기 → setup() 1회 실행 (시트 3개 생성 + 접근키 발급)
 *    실행 로그에 나오는 ACCESS KEY를 복사해 둔다
 * 3. 배포 → 새 배포 → 웹 앱
 *      실행: 나 / 액세스 권한: 모든 사용자
 * 4. 배포 URL + ACCESS KEY를 대시보드 [공유 설정]에 입력
 *
 * ── 보안 ──
 * URL과 키를 아는 사람은 누구나 읽고 쓸 수 있다. 링크를 아는 사람만 쓰는 구조다.
 * 외부에 URL을 공개하지 말 것. 유출되면 setup()을 다시 돌려 키를 바꾼다.
 */

const P = PropertiesService.getScriptProperties();
const SS = () => SpreadsheetApp.getActiveSpreadsheet();

/* 시트별 컬럼. 순서를 바꾸면 기존 데이터와 어긋난다 — 추가는 항상 끝에. */
const COLS = {
  refs: ['id','handle','link','followers','cat','tags','folder','desc','why','eng','ver','insights'],
  posts:['id','stage','refId','refHandle','insight','cat','tpl','topic','hook','script',
         'caption','tags','date','link','v','l','c','s','dm','lesson'],
  meta: ['key','value']
};
const JSON_COLS = { refs:['insights'] };          // 배열/객체로 저장되는 칸
const META_KEYS = ['ig','identity','goals','notes'];

function setup() {
  Object.keys(COLS).forEach(name => {
    let sh = SS().getSheetByName(name) || SS().insertSheet(name);
    sh.clear();
    sh.getRange(1,1,1,COLS[name].length).setValues([COLS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  const key = Utilities.getUuid().replace(/-/g,'').slice(0,20);
  P.setProperty('ACCESS_KEY', key);
  P.setProperty('REV', '1');
  Logger.log('설치 완료\nACCESS KEY: ' + key + '\n이 키를 대시보드 공유 설정에 입력하세요.');
  return key;
}

function doGet(e)  { return route(e, null); }
function doPost(e) { return route(e, parseBody(e)); }

function route(e, body) {
  try {
    const p = (e && e.parameter) || {};
    const key = body && body.key ? body.key : p.key;
    if (key !== P.getProperty('ACCESS_KEY')) return json({ ok:false, error:'접근키가 올바르지 않습니다' });
    return json(body ? write(body) : read());
  } catch (err) {
    return json({ ok:false, error: String(err && err.message || err) });
  }
}

function parseBody(e) {
  // 브라우저 preflight를 피하려고 text/plain으로 온다
  if (!e || !e.postData || !e.postData.contents) throw new Error('본문이 비어 있습니다');
  return JSON.parse(e.postData.contents);
}

function read() {
  const out = { ok:true, rev:+(P.getProperty('REV') || 1) };
  out.refs  = rows('refs');
  out.posts = rows('posts').map(unflattenMetrics);
  const meta = {};
  rows('meta').forEach(r => { try { meta[r.key] = JSON.parse(r.value); } catch (x) {} });
  META_KEYS.forEach(k => { if (meta[k] !== undefined) out[k] = meta[k]; });
  return out;
}

/* ponytail: 전체 덮어쓰기. 레코드 수백 건까지는 이게 제일 짧고 안전하다.
   수천 건이 되면 변경분만 보내는 방식으로 바꾼다. */
function write(body) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok:false, error:'다른 사람이 저장 중입니다. 잠시 후 다시 시도하세요.' };
  try {
    const cur = +(P.getProperty('REV') || 1);
    if (+body.rev !== cur) {
      return { ok:false, conflict:true, rev:cur, error:'다른 사람이 먼저 저장했습니다' };
    }
    if (Array.isArray(body.refs))  put('refs',  body.refs);
    if (Array.isArray(body.posts)) put('posts', body.posts.map(flattenMetrics));
    const meta = META_KEYS.filter(k => body[k] !== undefined)
                          .map(k => ({ key:k, value: JSON.stringify(body[k]) }));
    if (meta.length) put('meta', meta);

    const next = cur + 1;
    P.setProperty('REV', String(next));
    return { ok:true, rev:next, at:new Date().toISOString() };
  } finally {
    lock.releaseLock();
  }
}

/* ── 시트 ↔ 객체 ── */
function rows(name) {
  const sh = SS().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const cols = COLS[name];
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, cols.length).getValues();
  const jc = JSON_COLS[name] || [];
  return vals.filter(r => String(r[0]).length).map(r => {
    const o = {};
    cols.forEach((c, i) => {
      let v = r[i];
      if (jc.indexOf(c) > -1) { try { v = JSON.parse(v || '[]'); } catch (x) { v = []; } }
      o[c] = v;
    });
    return o;
  });
}

function put(name, list) {
  const sh = SS().getSheetByName(name);
  const cols = COLS[name], jc = JSON_COLS[name] || [];
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, cols.length).clearContent();
  if (!list.length) return;
  const vals = list.map(o => cols.map(c => {
    const v = o[c];
    if (jc.indexOf(c) > -1) return JSON.stringify(v || []);
    if (v === null || v === undefined) return '';
    // 시트가 긴 숫자 id를 지수표기로 바꾸지 않게 문자열로 넣는다
    return (c === 'id' || c === 'refId') && v !== '' ? "'" + v : v;
  }));
  sh.getRange(2, 1, vals.length, cols.length).setValues(vals);
}

/* 지표는 시트에서 한 칸씩 보이는 편이 읽기 좋다 */
function flattenMetrics(p) {
  const m = p.m || {};
  const o = {};
  Object.keys(p).forEach(k => { if (k !== 'm') o[k] = p[k]; });
  ['v','l','c','s','dm'].forEach(k => o[k] = +m[k] || 0);
  return o;
}
function unflattenMetrics(r) {
  const o = { m: { v:+r.v||0, l:+r.l||0, c:+r.c||0, s:+r.s||0, dm:+r.dm||0 } };
  Object.keys(r).forEach(k => { if (['v','l','c','s','dm'].indexOf(k) < 0) o[k] = r[k]; });
  if (o.tpl === '' || o.tpl === null) o.tpl = null; else o.tpl = +o.tpl;
  if (o.refId === '') o.refId = null;
  return o;
}

const json = o => ContentService.createTextOutput(JSON.stringify(o))
                                .setMimeType(ContentService.MimeType.JSON);

/* 키를 잊었을 때 */
function showKey() { Logger.log(P.getProperty('ACCESS_KEY')); }
