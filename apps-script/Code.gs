/**
 * SmartVibe | 십이지 수학탐험대 — 학급 기록 서버 (Apps Script v2)
 * ------------------------------------------------------------------
 * 1) 구글 시트 > 확장 프로그램 > Apps Script 에 이 코드를 통째로 붙여 넣어요.
 * 2) 아래 ADMIN_PASSWORD 를 선생님만 아는 비밀번호로 바꿔요. (학생 링크에는 절대 들어가지 않아요)
 * 3) 배포 > 새 배포 > 유형: 웹 앱 / 실행: 나 / 액세스 권한: 모든 사용자 → 배포 → 웹 앱 URL 복사
 * 4) 앱의 「👩‍🏫 내 반 설정」에 그 URL과 학급 코드를 넣고 학생용 링크를 만들어요.
 *
 * 시트 구성 (처음 제출이 들어오면 자동으로 만들어져요)
 *   결과      : 제출 1건 = 1줄 (점수·소요 시간·오답)
 *   문항      : 문항 1개 = 1줄 (처음에 맞혔는지 · 처음 고른 오답 · 정답) → 관리자 탭 문항 분석
 *   보석현황  : 차시 보석 획득 기록
 *   최고기록  : 혼자 하는 게임의 학생별 최고 점수 (v3)
 *
 * v3 추가: 🌐 링크로 함께 하는 보드게임(방 코드) · 🏆 학급 명예의 전당
 *   - 보드게임 방 정보는 시트가 아니라 임시 저장소(캐시)에 6시간만 있다가 사라져요.
 */
var ADMIN_PASSWORD = '여기에-비밀번호';   // ← 꼭 바꿔 주세요!
var VERSION = 3;

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents);
    if (d && d.action) return json_(action_(d));
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var now = new Date();
    var cls = String(d.classCode || '');
    var lesson = String(d.lessonId || d.unitId || '');

    var log = sheet_(ss, '결과', ['제출시각', '학급', '학번', '이름', '차시코드', '단원/차시', '세트', '점수', '세트1점수', '세트2점수', '소요시간(초)', '보석', '보석조각', '오답목록']);
    log.appendRow([now, cls, "'" + (d.studentId || ''), d.studentName || '', lesson, d.unit || '', d.setName || '',
      num_(d.score), num_(d.set1Score), num_(d.set2Score), num_(d.elapsedSec), d.gemEarned ? '💎' : '', d.fragments == null ? '' : d.fragments, (d.wrongText || '').slice(0, 45000)]);

    var items = d.items || [];
    if (items.length) {
      var it = sheet_(ss, '문항', ['제출시각', '학급', '학번', '이름', '차시코드', '문항ID', '구분', '문항', '처음정답', '시도', '처음오답', '정답', '문제']);
      var rows = items.slice(0, 120).map(function (x) {
        return [now, cls, "'" + (d.studentId || ''), d.studentName || '', lesson, String(x.qid || ''), x.kind || '', String(x.label || '').slice(0, 200),
          x.first ? 'O' : 'X', num_(x.tries), String(x.my || '').slice(0, 200), String(x.ans || '').slice(0, 200), String(x.q || '').slice(0, 300)];
      });
      it.getRange(it.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    if (d.gemEarned) {
      var g = sheet_(ss, '보석현황', ['학급', '학번', '이름', '차시코드', '보석', '획득시각']);
      var has = g.getDataRange().getValues().some(function (r) { return String(r[1]).replace("'", '') === String(d.studentId) && String(r[3]) === lesson; });
      if (!has) g.appendRow([cls, "'" + (d.studentId || ''), d.studentName || '', lesson, d.gemAnimal || '💎', now]);
    }
    return json_({ ok: true, v: VERSION });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** 관리자 탭이 부르는 곳 — 비밀번호가 맞을 때만 데이터를 돌려줘요. (JSONP 지원) */
function doGet(e) {
  var p = e.parameter || {}, out;
  if (p.action === 'ping') out = { ok: true, v: VERSION, title: SpreadsheetApp.getActiveSpreadsheet().getName() };
  else if (p.action === 'scoreTop') out = top_(p);
  else if (p.action && /^room/.test(p.action)) out = action_(p);
  else if (p.action === 'stats') {
    if (!check_(p.key)) out = { ok: false, error: '비밀번호가 맞지 않아요.' };
    else {
      var ss = SpreadsheetApp.getActiveSpreadsheet(), cls = String(p['class'] || ''), unit = String(p.unit || '');
      out = { ok: true, v: VERSION, title: ss.getName(), results: rows_(ss, '결과', cls, unit, 2000), items: rows_(ss, '문항', cls, unit, 15000) };
    }
  } else out = { ok: true, v: VERSION, msg: 'SmartVibe 학급 기록 서버가 잘 작동하고 있어요.' };
  var body = JSON.stringify(out);
  if (p.callback && /^[\w.$]+$/.test(p.callback))
    return ContentService.createTextOutput(p.callback + '(' + body + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return json_(out);
}

/* ---------- 도우미 ---------- */
function check_(key) {
  if (!ADMIN_PASSWORD || ADMIN_PASSWORD === '여기에-비밀번호') return false;   // 비밀번호를 바꾸기 전에는 조회 불가
  return String(key || '') === String(ADMIN_PASSWORD);
}
function sheet_(ss, name, head) {
  var s = ss.getSheetByName(name);
  if (s && s.getLastRow() > 0 && String(s.getRange(1, 2).getValue()) !== head[1]) {   // 예전(v1) 시트는 이름을 바꿔 보관해요
    s.setName(name + '(이전 ' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'MMdd-HHmm') + ')'); s = null;
  }
  if (!s) { s = ss.insertSheet(name); s.appendRow(head); s.setFrozenRows(1); s.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#fef3c7'); }
  else if (s.getLastRow() === 0) s.appendRow(head);
  return s;
}
function rows_(ss, name, cls, unit, max) {
  var s = ss.getSheetByName(name); if (!s || s.getLastRow() < 2) return [];
  var v = s.getDataRange().getValues(), head = v[0], out = [];
  var iCls = head.indexOf('학급'), iLes = head.indexOf('차시코드');
  for (var r = v.length - 1; r >= 1 && out.length < max; r--) {
    var row = v[r];
    if (cls && iCls >= 0 && String(row[iCls]) !== cls) continue;
    if (unit && iLes >= 0 && String(row[iLes]).indexOf(unit) !== 0) continue;
    var o = {};
    for (var c = 0; c < head.length; c++) o[head[c]] = row[c] instanceof Date ? row[c].toISOString() : String(row[c]).replace(/^'/, '');
    out.push(o);
  }
  return out.reverse();
}
function num_(x) { return x == null || x === '' ? '' : Number(x); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* =====================================================================
   🌐 v3 · 링크로 함께 하는 보드게임 (방 코드) + 🏆 최고 기록
   ===================================================================== */
var ROOM_TTL = 21600; // 6시간
function room_(code) { var v = CacheService.getScriptCache().get('room:' + code); return v ? JSON.parse(v) : null; }
function roomSave_(code, r) { r.ts = Date.now(); CacheService.getScriptCache().put('room:' + code, JSON.stringify(r), ROOM_TTL); }
function clean_(x, n) { return String(x == null ? '' : x).replace(/[<>]/g, '').slice(0, n || 20); }
function action_(d) {
  var a = String(d.action), code = clean_(d.code, 6), r;
  if (a === 'roomCreate') {
    for (var i = 0; i < 30; i++) { code = String(Math.floor(1000 + Math.random() * 9000)); if (!room_(code)) break; }
    r = { code: code, cls: clean_(d.cls, 20), lesson: clean_(d.lesson, 20), game: clean_(d.game, 20), ver: 0, started: false,
      players: [{ name: clean_(d.name, 12) || '1P', char: clean_(d.char, 20), sid: clean_(d.sid, 12) }], g: null, party: null };
    roomSave_(code, r); return { ok: true, code: code, seat: 0, room: r };
  }
  r = room_(code); if (!r) return { ok: false, error: '방을 찾을 수 없어요. 방 코드를 다시 확인해 주세요.' };
  if (a === 'roomJoin') {
    var nm = clean_(d.name, 12), sid = clean_(d.sid, 12);
    for (var k = 0; k < r.players.length; k++) if (sid && r.players[k].sid === sid) return { ok: true, code: code, seat: k, room: r };
    if (r.started) return { ok: false, error: '이미 시작한 방이에요.' };
    if (r.players.length >= 4) return { ok: false, error: '방이 가득 찼어요. (최대 4명)' };
    r.players.push({ name: nm || (r.players.length + 1) + 'P', char: clean_(d.char, 20), sid: sid }); r.ver++; roomSave_(code, r);
    return { ok: true, code: code, seat: r.players.length - 1, room: r };
  }
  if (a === 'roomGet') { if (+d.ver === r.ver) return { ok: true, same: true, ver: r.ver }; return { ok: true, room: r }; }
  if (a === 'roomPush') {
    if (d.base != null && +d.base < r.ver && !d.force) return { ok: false, stale: true, room: r };
    if (d.g !== undefined) r.g = d.g; if (d.party !== undefined) r.party = d.party; if (d.started) r.started = true;
    if (d.players) r.players = d.players;
    r.ver++; roomSave_(code, r); return { ok: true, ver: r.ver };
  }
  if (a === 'scoreSave') {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = sheet_(ss, '최고기록', ['학급', '학번', '이름', '차시코드', '게임', '최고점수', '기록시각']);
    var v = sh.getDataRange().getValues(), cls = clean_(d.cls, 20), sid2 = clean_(d.sid, 12), les = clean_(d.lesson, 20), gm = clean_(d.game, 30), sc = Number(d.score) || 0;
    for (var i2 = 1; i2 < v.length; i2++) {
      if (String(v[i2][0]) === cls && String(v[i2][1]).replace("'", '') === sid2 && String(v[i2][3]) === les && String(v[i2][4]) === gm) {
        if (sc > Number(v[i2][5])) { sh.getRange(i2 + 1, 3, 1, 5).setValues([[clean_(d.name, 12), les, gm, sc, new Date()]]); return { ok: true, best: sc, up: true }; }
        return { ok: true, best: Number(v[i2][5]), up: false };
      }
    }
    sh.appendRow([cls, "'" + sid2, clean_(d.name, 12), les, gm, sc, new Date()]); return { ok: true, best: sc, up: true };
  }
  return { ok: false, error: '알 수 없는 요청이에요.' };
}
/** 학급 명예의 전당 (같은 반 · 같은 차시 · 같은 게임 상위 10명) — 학번은 돌려주지 않아요 */
function top_(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName('최고기록'); if (!sh || sh.getLastRow() < 2) return { ok: true, top: [] };
  var v = sh.getDataRange().getValues(), cls = clean_(d.cls, 20), les = clean_(d.lesson, 20), gm = clean_(d.game, 30), out = [];
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === cls && String(v[i][3]) === les && String(v[i][4]) === gm) out.push({ name: String(v[i][2]), score: Number(v[i][5]) });
  out.sort(function (x, y) { return y.score - x.score; }); return { ok: true, top: out.slice(0, 10) };
}
