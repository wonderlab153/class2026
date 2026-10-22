/**
 * 대입 합격 현황 입력 – Google Apps Script 서버
 *
 * 설치 순서
 *  1) 새 구글 시트를 만들고 [확장 프로그램] > [Apps Script]를 엽니다.
 *  2) 이 파일 내용을 Code.gs에 통째로 붙여 넣고 저장합니다.
 *  3) 위쪽 함수 목록에서 setup 을 골라 [실행] → 권한 허용 (시트에 '학생', '원서' 탭이 생깁니다)
 *  4) 왼쪽 톱니바퀴 [프로젝트 설정] > [스크립트 속성] > 속성 추가
 *       속성: ADMIN_PW   값: (관리자 비밀번호)
 *     ※ 관리자 비밀번호는 코드나 웹페이지에 적지 않고 여기에만 둡니다.
 *  5) [배포] > [새 배포] > 유형: 웹 앱
 *       실행 사용자: 나 / 액세스 권한: 모든 사용자
 *     → 나오는 웹 앱 URL을 index.html의 CONFIG.API_URL에 넣습니다.
 *  코드를 고친 뒤에는 [배포] > [배포 관리] > 연필 > 버전 '새 버전'으로 다시 배포해야 반영됩니다.
 */

const SH_STU = '학생';
const SH_APP = '원서';
const TZ = 'Asia/Seoul';

const STU_COLS = [
  ['sid', '학번'], ['grade', '학년'], ['cls', '반'], ['num', '번호'], ['name', '이름'],
  ['pw', '비밀번호'], ['submittedAt', '제출시각'], ['pwChanged', '비번변경']
];
const APP_COLS = [
  ['id', 'ID'], ['sid', '학번'], ['name', '이름'], ['region', '지역'], ['kind', '구분'], ['univ', '대학명'],
  ['season', '모집시기'], ['ttype', '전형유형'], ['tbig', '전형명(대)'], ['tdetail', '세부유형'], ['field', '계열'],
  ['major', '모집단위'], ['quota', '모집인원'], ['src', '입력구분'], ['stage1', '1단계'], ['final', '최종단계'],
  ['reason', '불합격사유'], ['rankType', '최초후보순위구분'], ['rank', '최초후보순위'], ['reg', '등록여부'],
  ['note', '비고'], ['updatedAt', '수정시각']
];
const OPT = {
  stage1: ['합격', '불합격'],
  final: ['합격', '불합격', '충원합격'],
  reason: ['미응시', '최저미충족', '불합격'],
  rankType: ['직접입력', '후보순위없음'],
  reg: ['Y', 'N']
};
const STAGE1_FAIL = '1단계 불합격';
const BASE_F = ['region', 'kind', 'univ', 'season', 'ttype', 'tbig', 'tdetail', 'field', 'major', 'quota'];
const RES_F = ['stage1', 'final', 'reason', 'rankType', 'rank', 'reg', 'note'];
const STUDENT_SESSION_SEC = 3 * 3600;
const ADMIN_SESSION_SEC = 2 * 3600;
const MAX_FAILS = 8;          // 이 횟수만큼 틀리면
const LOCK_SEC = 600;         // 10분 동안 로그인 막음

/* ---------- 설치 ---------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  prep_(ss, SH_STU, STU_COLS);
  prep_(ss, SH_APP, APP_COLS);
  const first = ss.getSheetByName('시트1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 2) ss.deleteSheet(first);
}
function prep_(ss, name, cols) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');   // 학번 앞자리 0, 비번 보존
  sh.getRange(1, 1, 1, cols.length).setValues([cols.map(c => c[1])]).setFontWeight('bold').setBackground('#e6edf5');
  sh.setFrozenRows(1);
  return sh;
}

/* ---------- 입구 ---------- */
function doGet() {
  return json_({ ok: true, message: '대입 합격 현황 서버가 동작 중입니다.' });
}
function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const routes = {
      studentLogin: studentLogin_, changePw: changePw_, submit: submit_,
      adminLogin: adminLogin_, adminData: adminData_, adminResetPw: adminResetPw_, adminImport: adminImport_
    };
    const fn = routes[req.action];
    if (!fn) throw new Error('알 수 없는 요청입니다.');
    out = fn(req);
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return json_(out);
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 학생 ---------- */
function studentLogin_(req) {
  const sid = cut_(req.sid, 12).replace(/\D/g, '');
  const key = 's' + sid;
  if (locked_(key)) throw new Error('로그인 시도가 너무 많습니다. 10분 뒤 다시 시도하세요.');
  const st = readStu_().find(s => s.sid === sid);
  const okName = st && st.name.replace(/\s/g, '') === String(req.name || '').replace(/\s/g, '');
  if (!st || !okName || st.pw !== String(req.pw || '')) {
    fail_(key);
    throw new Error('학번, 이름, 비밀번호를 다시 확인하세요.');
  }
  clearFail_(key);
  const token = newSession_({ role: 'student', sid: sid }, STUDENT_SESSION_SEC);
  return { ok: true, token: token, student: pub_(st), apps: readApps_().filter(a => a.sid === sid) };
}

function changePw_(req) {
  const s = sess_(req.token, 'student');
  const newPw = String(req.newPw || '');
  if (newPw.length < 4 || newPw.length > 20) throw new Error('새 비밀번호는 4~20자로 정하세요.');
  return withLock_(() => {
    const list = readStu_();
    const me = list.find(x => x.sid === s.sid);
    if (!me) throw new Error('학생 정보를 찾지 못했습니다.');
    if (me.pw !== String(req.oldPw || '')) throw new Error('현재 비밀번호가 맞지 않습니다.');
    me.pw = newPw;
    me.pwChanged = 'Y';
    writeStu_(list);
    return { ok: true };
  });
}

function submit_(req) {
  const s = sess_(req.token, 'student');
  return withLock_(() => {
    const stu = readStu_();
    const me = stu.find(x => x.sid === s.sid);
    if (!me) throw new Error('학생 정보를 찾지 못했습니다.');
    const all = readApps_();
    const mine = all.filter(a => a.sid === s.sid);
    const others = all.filter(a => a.sid !== s.sid);
    const byId = {};
    mine.forEach(a => { byId[a.id] = a; });

    const now = now_();
    const out = [];
    const seen = {};
    const incoming = Array.isArray(req.apps) ? req.apps.slice(0, 30) : [];
    incoming.forEach(c => {
      const ex = byId[c.id];
      let row;
      if (ex && ex.src === 'base') {
        row = Object.assign({}, ex);                       // 학교 제공 원서: 대학 정보는 서버 값 유지
      } else {
        row = { id: ex ? ex.id : newId_('N'), sid: s.sid, name: me.name, src: 'added' };
        BASE_F.forEach(k => { row[k] = cut_(c[k], 60); });
        if (!row.univ) return;
      }
      if (seen[row.id]) return;
      seen[row.id] = true;
      RES_F.forEach(k => { row[k] = cut_(c[k], 200); });
      const n = normalize_(row);
      n.updatedAt = now;
      out.push(n);
    });
    mine.forEach(a => { if (a.src === 'base' && !seen[a.id]) out.push(a); });  // 학교 제공 원서는 지워지지 않음

    writeApps_(others.concat(out));
    me.submittedAt = now;
    writeStu_(stu);
    return { ok: true, apps: out, submittedAt: now };
  });
}

/* ---------- 관리자 ---------- */
function adminLogin_(req) {
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PW');
  if (!real) throw new Error('관리자 비밀번호가 아직 설정되지 않았습니다. 스크립트 속성에 ADMIN_PW를 추가하세요.');
  if (locked_('admin')) throw new Error('로그인 시도가 너무 많습니다. 10분 뒤 다시 시도하세요.');
  if (String(req.pw || '') !== real) {
    fail_('admin');
    throw new Error('관리자 비밀번호가 맞지 않습니다.');
  }
  clearFail_('admin');
  return { ok: true, token: newSession_({ role: 'admin' }, ADMIN_SESSION_SEC) };
}

function adminData_(req) {
  sess_(req.token, 'admin');
  return { ok: true, students: readStu_(), apps: readApps_() };
}

function adminResetPw_(req) {
  sess_(req.token, 'admin');
  return withLock_(() => {
    const list = readStu_();
    const s = list.find(x => x.sid === String(req.sid || ''));
    if (!s) throw new Error('학생을 찾지 못했습니다.');
    s.pw = s.sid;
    s.pwChanged = '';
    writeStu_(list);
    return { ok: true };
  });
}

function adminImport_(req) {
  sess_(req.token, 'admin');
  return withLock_(() => {
    const replace = req.mode === 'replace';
    const stu = replace ? [] : readStu_();
    const apps = replace ? [] : readApps_();
    const sMap = {};
    stu.forEach(s => { sMap[s.sid] = s; });
    let addS = 0, addA = 0;

    (req.students || []).forEach(x => {
      const sid = cut_(x.sid, 12).replace(/\D/g, '');
      if (!sid || sMap[sid]) return;
      const o = { sid: sid, grade: cut_(x.grade, 4), cls: cut_(x.cls, 4), num: cut_(x.num, 4),
        name: cut_(x.name, 30), pw: sid, submittedAt: '', pwChanged: '' };
      stu.push(o); sMap[sid] = o; addS++;
    });

    const key = a => [a.sid, a.univ, a.tdetail, a.major].join('|');
    const exist = {};
    apps.forEach(a => { exist[key(a)] = true; });
    (req.apps || []).forEach(x => {
      const sid = cut_(x.sid, 12).replace(/\D/g, '');
      if (!sMap[sid]) return;
      const o = { id: newId_('B'), sid: sid, name: sMap[sid].name, src: 'base' };
      BASE_F.forEach(k => { o[k] = cut_(x[k], 60); });
      RES_F.forEach(k => { o[k] = cut_(x[k], 200); });
      if (!o.univ || exist[key(o)]) return;
      exist[key(o)] = true;
      const n = normalize_(o);
      n.updatedAt = '';
      apps.push(n); addA++;
    });

    stu.sort((a, b) => a.sid.localeCompare(b.sid));
    writeStu_(stu);
    writeApps_(apps);
    return { ok: true, addedStudents: addS, addedApps: addA, totalStudents: stu.length, totalApps: apps.length };
  });
}

/* ---------- 입력 규칙 (index.html의 normalize와 같음) ---------- */
function normalize_(a) {
  const x = Object.assign({}, a);
  ['stage1', 'final', 'rankType', 'reg'].forEach(k => { if (x[k] && OPT[k].indexOf(x[k]) < 0) x[k] = ''; });
  if (x.reason && OPT.reason.indexOf(x.reason) < 0 && x.reason !== STAGE1_FAIL) x.reason = '';
  x.note = String(x.note || '').slice(0, 200);
  if (x.stage1 === '불합격') {
    x.final = '불합격'; x.reason = STAGE1_FAIL; x.rankType = '후보순위없음'; x.rank = ''; x.reg = 'N';
    return x;
  }
  if (x.reason === STAGE1_FAIL) x.reason = '';
  if (x.final !== '불합격') x.reason = '';
  if (x.final !== '불합격' && x.final !== '충원합격') { x.rankType = ''; x.rank = ''; }
  x.rank = x.rankType === '직접입력' ? String(x.rank || '').replace(/\D/g, '').slice(0, 4) : '';
  if (x.final === '불합격') x.reg = 'N';
  if (!x.final) x.reg = '';
  return x;
}

/* ---------- 시트 읽기·쓰기 ---------- */
function sheet_(name, cols) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || prep_(ss, name, cols);
}
function read_(name, cols) {
  const sh = sheet_(name, cols);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, cols.length).getDisplayValues().map(r => {
    const o = {};
    cols.forEach((c, i) => { o[c[0]] = String(r[i]).trim(); });
    return o;
  }).filter(o => o.sid);
}
function write_(name, cols, list) {
  const sh = sheet_(name, cols);
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, cols.length).clearContent();
  if (!list.length) return;
  const rg = sh.getRange(2, 1, list.length, cols.length);
  rg.setNumberFormat('@');
  rg.setValues(list.map(o => cols.map(c => (o[c[0]] == null ? '' : String(o[c[0]])))));
}
function readStu_() { return read_(SH_STU, STU_COLS); }
function writeStu_(list) { write_(SH_STU, STU_COLS, list); }
function readApps_() {
  return read_(SH_APP, APP_COLS).map(a => { a.src = a.src === '학생추가' ? 'added' : 'base'; return a; });
}
function writeApps_(list) {
  write_(SH_APP, APP_COLS, list.map(a => Object.assign({}, a, { src: a.src === 'added' ? '학생추가' : '학교제공' })));
}

/* ---------- 접속·보안 ---------- */
function pub_(s) {
  return { sid: s.sid, name: s.name, submittedAt: s.submittedAt, initialPw: s.pw === s.sid };
}
function newSession_(data, sec) {
  const token = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
  CacheService.getScriptCache().put('sess_' + token, JSON.stringify(data), sec);
  return token;
}
function sess_(token, role) {
  const raw = token ? CacheService.getScriptCache().get('sess_' + token) : null;
  const s = raw ? JSON.parse(raw) : null;
  if (!s || s.role !== role) throw new Error('접속이 끝났습니다. 다시 로그인하세요.');
  return s;
}
function locked_(key) {
  return Number(CacheService.getScriptCache().get('fail_' + key) || 0) >= MAX_FAILS;
}
function fail_(key) {
  const c = CacheService.getScriptCache();
  c.put('fail_' + key, String(Number(c.get('fail_' + key) || 0) + 1), LOCK_SEC);
}
function clearFail_(key) { CacheService.getScriptCache().remove('fail_' + key); }
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('지금 접속이 많습니다. 잠시 뒤 다시 제출하세요.');
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------- 도구 ---------- */
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function newId_(p) { return p + Utilities.getUuid().replace(/-/g, '').slice(0, 12); }
function cut_(v, n) { return String(v == null ? '' : v).trim().slice(0, n); }
