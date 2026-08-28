/* 신부장 인스타 동기화 — Google Apps Script 웹앱
 *
 * 대시보드(index.html)가 이 웹앱 주소 하나만 호출하면 본인 계정 지표가 채워진다.
 * 토큰은 여기(스크립트 속성)에만 있다. HTML에 넣으면 파일 받은 사람 누구나 계정을 조작할 수 있다.
 *
 * ── 설치 ──
 * 1. script.google.com → 새 프로젝트 → 이 파일 붙여넣기
 * 2. 프로젝트 설정 → 스크립트 속성 → IG_TOKEN = 장기 액세스 토큰
 * 3. setup() 1회 실행 (토큰 자동 갱신 트리거 등록)
 * 4. 배포 → 새 배포 → 웹 앱 / 실행: 나 / 액세스: 모든 사용자 → 주소 복사
 * 5. 대시보드 헤더 [동기화] → 주소 붙여넣기
 */

const P = PropertiesService.getScriptProperties();
const GRAPH = 'https://graph.instagram.com';
const VER = 'v23.0';
const CACHE_SEC = 600;   // 인스타 지표는 이보다 자주 안 바뀐다. 호출 한도도 아낀다.
const LIMIT = 100;       // 가져올 최근 게시물 수

function setup() {
  if (!P.getProperty('IG_TOKEN')) throw new Error('먼저 스크립트 속성에 IG_TOKEN을 넣으세요');
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'refreshToken') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('refreshToken').timeBased().everyDays(30).create();
  refreshToken();
  Logger.log('설치 완료. 토큰 30일마다 자동 갱신됩니다.');
}

function doGet(e) {
  try {
    const fresh = e && e.parameter && e.parameter.fresh;
    const cache = CacheService.getScriptCache();
    if (!fresh) {
      const hit = cache.get('ig');
      if (hit) return json(JSON.parse(hit));
    }
    const data = collect();
    cache.put('ig', JSON.stringify(data), CACHE_SEC);
    return json(data);
  } catch (err) {
    return json({ ok: false, error: String(err.message || err) });
  }
}

function collect() {
  const token = P.getProperty('IG_TOKEN');
  if (!token) throw new Error('IG_TOKEN이 없습니다. 스크립트 속성을 확인하세요.');

  const me = api('/me?fields=username,followers_count,media_count', token);
  const media = (api('/me/media?fields=id,permalink,timestamp,caption,like_count,comments_count,media_product_type&limit=' + LIMIT, token).data) || [];

  const posts = media.map(m => {
    const ins = insights(m.id, token);
    return {
      permalink: m.permalink,
      date: (m.timestamp || '').slice(0, 10),
      caption: (m.caption || '').slice(0, 120),
      isReel: m.media_product_type === 'REELS',
      v: ins.views || 0,          // 조회
      l: m.like_count || 0,       // 좋아요
      c: m.comments_count || 0,   // 댓글
      s: ins.saved || 0,          // 저장
      sh: ins.shares || 0,        // 공유
      reach: ins.reach || 0
    };
  });

  return {
    ok: true,
    at: new Date().toISOString(),
    username: me.username,
    followers: me.followers_count || 0,
    mediaCount: me.media_count || 0,
    posts: posts
  };
}

/* 미디어 종류마다 지원 지표가 다르다. 하나 실패해도 나머지는 살린다. */
function insights(id, token) {
  const out = {};
  ['views,reach,saved,shares', 'reach,saved'].some(metrics => {
    const r = api('/' + id + '/insights?metric=' + metrics, token, true);
    if (!r || !r.data) return false;
    r.data.forEach(x => { out[x.name] = (x.values && x.values[0] && x.values[0].value) || 0; });
    return true;
  });
  return out;
}

function api(path, token, quiet) {
  const sep = path.indexOf('?') > -1 ? '&' : '?';
  const url = GRAPH + '/' + VER + path + sep + 'access_token=' + encodeURIComponent(token);
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const body = JSON.parse(res.getContentText() || '{}');
  if (body.error) {
    if (quiet) return null;
    throw new Error('Instagram: ' + body.error.message);
  }
  return body;
}

/* 장기 토큰은 60일이면 만료된다. 30일마다 갱신해서 끊기지 않게 한다. */
function refreshToken() {
  const token = P.getProperty('IG_TOKEN');
  if (!token) return;
  const res = UrlFetchApp.fetch(
    GRAPH + '/refresh_access_token?grant_type=ig_refresh_token&access_token=' + encodeURIComponent(token),
    { muteHttpExceptions: true });
  const body = JSON.parse(res.getContentText() || '{}');
  if (body.access_token) {
    P.setProperty('IG_TOKEN', body.access_token);
    Logger.log('토큰 갱신 완료');
  } else {
    Logger.log('토큰 갱신 실패 — 재발급이 필요합니다: ' + res.getContentText());
  }
}

const json = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
