/**
 * 2026 서울사회복지사 등반대회 — 준비물 3단 체크리스트 API (v2)
 *
 * 배포
 *  1) script.google.com 새 프로젝트 → 이 파일 내용 붙여넣기
 *  2) 아래 ADMIN_KEY 를 원하는 값으로 바꾸기
 *  3) 실행 → seed  (최초 1회, 시트 생성 + 72항목 입력)
 *  4) 배포 → 새 배포 → 웹 앱 / 실행: 나 / 액세스: 모든 사용자
 *  5) /exec URL 을 public/config.js 에 붙여넣기
 *
 * ※ 코드를 고친 뒤에는 반드시 "배포 관리 → 새 버전"으로 재배포해야 반영됩니다.
 */

var SHEET_ID   = '1PQy4OPfn3ldPuudmJwKzJtiOm8itfUPEpORyYqWfXlo';
var SHEET_NAME = '준비물체크리스트';
var LOG_NAME   = '체크로그';

/** 관리자 PIN — 반드시 바꿔서 쓰세요 */
var ADMIN_KEY = 'sasw2026';

var HEADERS = [
  'id', '구분', '준비물', '수량', '담당',
  '0910', '0911', '0912',
  '0910확인자', '0911확인자', '0912확인자',
  '갱신시각', '메모'
];

var STAGE_COL = { '0910': 6, '0911': 7, '0912': 8 };
var ACTOR_COL = { '0910': 9, '0911': 10, '0912': 11 };
var TIME_COL  = 12;
var MEMO_COL  = 13;

/** 구분별 id 접두어 */
var PREFIX = {
  '접수·명단': 'RCP',
  '기념품·배부': 'GFT',
  '스탬프·완주인증': 'STP',
  '경품·추첨': 'PRZ',
  '무대·행사장': 'STG',
  '부스-공정위원회': 'BFR',
  '부스-청년위원회': 'BYT',
  '부스-열매 만남존': 'BYM',
  '부스-어린이 이벤트': 'BKD',
  '부스-함께하는 단체': 'BTG',
  '등반로 1·2·3지점': 'TRL',
  '철수·정리': 'TRD',
  '기타': 'ETC'
};

/* ────────────────────────── 라우팅 ────────────────────────── */

function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try {
    switch (p.action) {
      /* 공개 */
      case 'set':   out = ok(setCheck(p.id, p.stage, p.value === '1', p.actor || '')); break;
      case 'note':  out = ok(setNote(p.id, p.text || '')); break;
      case 'stats': out = ok(stats()); break;

      /* 관리자 */
      case 'auth':       guard(p); out = ok({ admin: true }); break;
      case 'add':        guard(p); out = ok(addItem(p)); break;
      case 'update':     guard(p); out = ok(updateItem(p)); break;
      case 'remove':     guard(p); out = ok(removeItem(p.id)); break;
      case 'move':       guard(p); out = ok(moveItem(p.id, p.dir)); break;
      case 'renameWho':  guard(p); out = ok(renameField(5, p.from, p.to)); break;
      case 'renameCat':  guard(p); out = ok(renameField(2, p.from, p.to)); break;
      case 'resetStage': guard(p); out = ok(resetStage(p.stage)); break;
      case 'export':     guard(p); out = ok({ code: exportItemsJs() }); break;
      case 'seed':       guard(p); out = ok(seedRows(p.force === '1')); break;

      /* 당일 안내 액션은 아래 default 에서 guideRoute 로 위임 */
      default:
        var g = (typeof guideRoute === 'function') ? guideRoute(p) : null;
        out = ok(g !== null ? g : listRows());
    }
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return reply(out, p.callback);
}

function doPost(e) {
  var p = {};
  try { p = JSON.parse((e.postData && e.postData.contents) || '{}'); } catch (ignore) {}
  var out;
  try { out = ok(setCheck(p.id, p.stage, !!p.value, p.actor || '')); }
  catch (err) { out = { ok: false, error: String(err) }; }
  return reply(out, null);
}

function ok(data) { return { ok: true, data: data }; }

function guard(p) {
  if (String(p.key || '') !== ADMIN_KEY) throw new Error('관리자 인증에 실패했습니다');
}

function reply(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/* ────────────────────────── 시트 ────────────────────────── */

function sheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setBackground('#14261b').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 70);
    sh.setColumnWidth(2, 130);
    sh.setColumnWidth(3, 330);
    sh.setColumnWidth(4, 150);
    sh.setColumnWidth(5, 130);
  }
  return sh;
}

function logSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(LOG_NAME);
  if (!sh) {
    sh = ss.insertSheet(LOG_NAME);
    sh.getRange(1, 1, 1, 6).setValues([['시각', 'id', '단계', '값', '확인자', '비고']])
      .setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function audit(action, id, detail, who) {
  logSheet().appendRow([new Date(), id || '', action, detail || '', who || '', '관리']);
}

function listRows() {
  var sh = sheet();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var v = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  var out = [];
  for (var i = 0; i < v.length; i++) {
    var r = v[i];
    if (!r[0]) continue;
    out.push({
      id: String(r[0]), cat: String(r[1]), name: String(r[2]),
      qty: String(r[3]), who: String(r[4]),
      done: [truthy(r[5]), truthy(r[6]), truthy(r[7])],
      by: [String(r[8] || ''), String(r[9] || ''), String(r[10] || '')],
      memo: String(r[12] || '')
    });
  }
  return out;
}

function truthy(v) {
  if (v === true) return true;
  var s = String(v).trim().toUpperCase();
  return s === 'TRUE' || s === 'O' || s === 'Y' || s === '1' || s === 'V';
}

function rowIndexById(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function stats() {
  var rows = listRows();
  var out = { total: rows.length, stages: {} };
  ['0910', '0911', '0912'].forEach(function (k, i) {
    out.stages[k] = rows.filter(function (r) { return r.done[i]; }).length;
  });
  return out;
}

/* ────────────────────────── 체크 ────────────────────────── */

function setCheck(id, stage, value, actor) {
  if (!id) throw new Error('id가 없습니다');
  if (!STAGE_COL[stage]) throw new Error('단계 값이 잘못되었습니다: ' + stage);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet();
    var row = rowIndexById(sh, id);
    if (row < 0) throw new Error('항목을 찾을 수 없습니다: ' + id);
    var now = new Date();
    sh.getRange(row, STAGE_COL[stage]).setValue(!!value);
    sh.getRange(row, ACTOR_COL[stage]).setValue(value ? (actor || '') : '');
    sh.getRange(row, TIME_COL).setValue(now);
    logSheet().appendRow([now, id, stage, value ? 'O' : '-', actor || '', '']);
    return { id: id, stage: stage, value: !!value };
  } finally { lock.releaseLock(); }
}

function setNote(id, text) {
  var sh = sheet();
  var row = rowIndexById(sh, id);
  if (row < 0) throw new Error('항목을 찾을 수 없습니다: ' + id);
  sh.getRange(row, MEMO_COL).setValue(text);
  return { id: id, memo: text };
}

function resetStage(stage) {
  if (!STAGE_COL[stage]) throw new Error('단계 값이 잘못되었습니다: ' + stage);
  var sh = sheet();
  var last = sh.getLastRow();
  if (last < 2) return { cleared: 0 };
  var n = last - 1;
  var blanks = [], falses = [];
  for (var i = 0; i < n; i++) { blanks.push(['']); falses.push([false]); }
  sh.getRange(2, STAGE_COL[stage], n, 1).setValues(falses);
  sh.getRange(2, ACTOR_COL[stage], n, 1).setValues(blanks);
  audit('초기화', '', stage, '');
  return { cleared: n, stage: stage };
}

/* ────────────────────────── 관리자: 항목 편집 ────────────────────────── */

function nextId(cat) {
  var pre = PREFIX[cat] || 'NEW';
  var rows = listRows();
  var max = 0;
  rows.forEach(function (r) {
    var m = String(r.id).match(/^([A-Z]+)(\d+)$/);
    if (m && m[1] === pre) max = Math.max(max, parseInt(m[2], 10));
  });
  var n = String(max + 1);
  while (n.length < 2) n = '0' + n;
  return pre + n;
}

function addItem(p) {
  var cat  = String(p.cat  || '').trim();
  var name = String(p.name || '').trim();
  if (!cat)  throw new Error('구분을 입력해 주세요');
  if (!name) throw new Error('준비물 이름을 입력해 주세요');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet();
    var id = nextId(cat);
    var rows = listRows();

    /* 같은 구분의 마지막 행 바로 뒤에 끼워 넣기 */
    var insertAt = sh.getLastRow() + 1;
    for (var i = rows.length - 1; i >= 0; i--) {
      if (rows[i].cat === cat) { insertAt = i + 3; break; }
    }
    if (insertAt <= sh.getLastRow()) sh.insertRowBefore(insertAt);

    sh.getRange(insertAt, 1, 1, HEADERS.length).setValues([[
      id, cat, name, String(p.qty || ''), String(p.who || ''),
      false, false, false, '', '', '', new Date(), String(p.memo || '')
    ]]);
    sh.getRange(insertAt, 6, 1, 3).insertCheckboxes();
    audit('추가', id, cat + ' / ' + name, p.actor || '');
    return { id: id, row: insertAt };
  } finally { lock.releaseLock(); }
}

function updateItem(p) {
  var id = String(p.id || '');
  if (!id) throw new Error('id가 없습니다');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet();
    var row = rowIndexById(sh, id);
    if (row < 0) throw new Error('항목을 찾을 수 없습니다: ' + id);
    var changed = [];
    if (p.cat  !== undefined) { sh.getRange(row, 2).setValue(String(p.cat));  changed.push('구분'); }
    if (p.name !== undefined) { sh.getRange(row, 3).setValue(String(p.name)); changed.push('준비물'); }
    if (p.qty  !== undefined) { sh.getRange(row, 4).setValue(String(p.qty));  changed.push('수량'); }
    if (p.who  !== undefined) { sh.getRange(row, 5).setValue(String(p.who));  changed.push('담당'); }
    if (p.memo !== undefined) { sh.getRange(row, MEMO_COL).setValue(String(p.memo)); changed.push('메모'); }
    sh.getRange(row, TIME_COL).setValue(new Date());
    audit('수정', id, changed.join(','), p.actor || '');
    return { id: id, changed: changed };
  } finally { lock.releaseLock(); }
}

function removeItem(id) {
  if (!id) throw new Error('id가 없습니다');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet();
    var row = rowIndexById(sh, id);
    if (row < 0) throw new Error('항목을 찾을 수 없습니다: ' + id);
    var name = sh.getRange(row, 3).getValue();
    sh.deleteRow(row);
    audit('삭제', id, String(name), '');
    return { id: id, removed: true };
  } finally { lock.releaseLock(); }
}

function moveItem(id, dir) {
  var up = String(dir) === 'up';
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet();
    var row = rowIndexById(sh, id);
    if (row < 0) throw new Error('항목을 찾을 수 없습니다: ' + id);
    var target = up ? row - 1 : row + 1;
    if (target < 2 || target > sh.getLastRow()) return { id: id, moved: false };

    var a = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];
    var b = sh.getRange(target, 1, 1, HEADERS.length).getValues()[0];
    sh.getRange(row, 1, 1, HEADERS.length).setValues([b]);
    sh.getRange(target, 1, 1, HEADERS.length).setValues([a]);
    sh.getRange(2, 6, sh.getLastRow() - 1, 3).insertCheckboxes();
    return { id: id, moved: true };
  } finally { lock.releaseLock(); }
}

/** 담당자(5열) 또는 구분(2열) 이름을 전체 일괄 변경. to 를 비우면 제거 */
function renameField(col, from, to) {
  from = String(from || '').trim();
  to   = String(to == null ? '' : to).trim();
  if (!from) throw new Error('바꿀 이름을 입력해 주세요');

  var sh = sheet();
  var last = sh.getLastRow();
  if (last < 2) return { count: 0 };
  var rng = sh.getRange(2, col, last - 1, 1);
  var v = rng.getValues();
  var count = 0;

  for (var i = 0; i < v.length; i++) {
    var cur = String(v[i][0] || '');
    if (col === 5) {
      /* 담당은 콤마로 구분된 목록이므로 이름 단위로 교체 */
      var parts = cur.split(/\s*,\s*/).filter(function (s) { return s.length; });
      var hit = false, next = [];
      for (var j = 0; j < parts.length; j++) {
        if (parts[j] === from) { hit = true; if (to) next.push(to); }
        else next.push(parts[j]);
      }
      if (hit) {
        var seen = {}, uniq = [];
        for (var k = 0; k < next.length; k++) {
          if (!seen[next[k]]) { seen[next[k]] = 1; uniq.push(next[k]); }
        }
        v[i][0] = uniq.join(', ');
        count++;
      }
    } else if (cur === from) {
      v[i][0] = to;
      count++;
    }
  }
  if (count) rng.setValues(v);
  audit(col === 5 ? '담당 일괄변경' : '구분 일괄변경', '',
        from + ' → ' + (to || '(제거)') + ' / ' + count + '건', '');
  return { count: count, from: from, to: to };
}

/* ────────────────────────── 내보내기 / 시드 ────────────────────────── */

/** 현재 시트 내용을 public/items.js 형식 코드로 반환 */
function exportItemsJs() {
  var rows = listRows();
  var lines = rows.map(function (r) {
    return '  ["' + r.id + '","' + r.cat + '","' + r.name.replace(/"/g, '\\"') +
           '","' + r.qty.replace(/"/g, '\\"') + '","' + r.who + '"]';
  });
  return 'window.CHECKLIST_ITEMS = [\n' + lines.join(',\n') + '\n];\n';
}

function seed()      { return seedRows(false); }
function seedReset() { return seedRows(true); }

function seedRows(force) {
  var sh = sheet();
  var keep = {};
  if (!force) listRows().forEach(function (r) { keep[r.id] = r; });
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, HEADERS.length).clearContent();

  var values = ITEMS.map(function (it) {
    var k = keep[it[0]];
    return [
      it[0], it[1], it[2], it[3], it[4],
      k ? k.done[0] : false, k ? k.done[1] : false, k ? k.done[2] : false,
      k ? k.by[0] : '', k ? k.by[1] : '', k ? k.by[2] : '',
      '', k ? k.memo : ''
    ];
  });
  sh.getRange(2, 1, values.length, HEADERS.length).setValues(values);
  sh.getRange(2, 6, values.length, 3).insertCheckboxes();
  return { rows: values.length, kept: Object.keys(keep).length };
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('등반대회 체크리스트')
    .addItem('항목 다시 쓰기(체크 유지)', 'seed')
    .addItem('전체 초기화(체크까지 삭제)', 'seedReset')
    .addToUi();
}

/* ────────────────────────── 준비물 마스터(72항목) ────────────────────────── */
/* seed 전용 초기값입니다. 관리자 화면에서 항목을 고친 뒤에는
   ?action=export&key=... 로 최신 목록을 받아 여기와 public/items.js 를 갱신하세요. */
var ITEMS = [
  ['RCP01', '접수·명단', '노트북', '4대(접수 3대+예비 1대)', '이진선, 정소희'],
  ['RCP02', '접수·명단', '사전신청 최종 명단(엑셀 파일 + 출력본)', '각 2부', '정소희'],
  ['RCP03', '접수·명단', '현장접수용 종이 양식(비상용)', '여분 확보', '정소희'],
  ['RCP04', '접수·명단', '방명록 및 여분 용지', '3~4장', '이진선'],
  ['RCP05', '접수·명단', '상장 케이스(방명록 받침용)', '2개', '정승아'],
  ['RCP06', '접수·명단', '접수대 번호 표지(접수대 1·2·3)', '각 1개', '최봄'],
  ['RCP07', '접수·명단', '줄 유도선(공항형 한 줄 대기)', '1식', '최지혜'],
  ['RCP08', '접수·명단', '필기구 세트(볼펜·칼·가위·테이프·집게·포스트잇)', '2세트(접수대·무대)', '최지혜'],
  ['RCP09', '접수·명단', '아크릴 스탠드(현장접수 QR 거치)', '3~5개', '정소희'],

  ['GFT01', '기념품·배부', '손수건 기념품(2026년판 800개+구디자인 500개)', '총 1,300개', '양종철'],
  ['GFT02', '기념품·배부', '생수(청밀 후원)', '1,800개', '양종철'],
  ['GFT03', '기념품·배부', '아이스크림(설레임, 공제회 후원)', '1,000개(11:30 도착)', '최봄'],
  ['GFT04', '기념품·배부', '아이스박스·아이스팩(아이스크림 보관)', '일괄', '최봄'],
  ['GFT05', '기념품·배부', '문화상품권 5,000원권(어린이 이벤트)', '100장', '권하영, 정소희'],
  ['GFT06', '기념품·배부', '김밥(준비팀 간식)', '80줄(07:30~08:00 도착)', '고석우'],
  ['GFT07', '기념품·배부', '협회 단체복(티셔츠)', '24벌(행사 후 회수)', '양종철'],
  ['GFT08', '기념품·배부', '갈림길 안내요원 조끼', '배치 인원분', '이해창'],
  ['GFT09', '기념품·배부', '스태프 식별용 수건(색상 통일)', '스태프 인원분', '정승아'],

  ['STP01', '스탬프·완주인증', '스탬프 투어 용지', '1,500장', '정소희'],
  ['STP02', '스탬프·완주인증', '도장(지점별 3~5개, 1지점 추가)', '12~15개', '정소희'],
  ['STP03', '스탬프·완주인증', '지점 표지 출력물(A3, 클리어파일)', '1·2·3지점 각 2매', '이해창'],
  ['STP04', '스탬프·완주인증', '완주자 QR 출력물 및 거치 스탠드', '각 5부 이상', '고석우'],
  ['STP05', '스탬프·완주인증', '스탬프 용지 회수함', '2개', '이진선'],
  ['STP06', '스탬프·완주인증', '초록색 손수건·깃발·케이블타이', '약 30세트', '이해창'],

  ['PRZ01', '경품·추첨', '추첨용지(성함·소속·전화번호 뒷자리)', 'A4 4~6등분', '전체 직원'],
  ['PRZ02', '경품·추첨', '추첨함', '1개', '정소희'],
  ['PRZ03', '경품·추첨', '경품 실물(아이패드·헤드폰·애플워치)', '5점', '이정하, 양종철'],
  ['PRZ04', '경품·추첨', '교촌치킨 박스', '2개', '이정하'],
  ['PRZ05', '경품·추첨', '상품 이미지 출력물(1·2·3등 표시)', '각 1매', '신단비'],
  ['PRZ06', '경품·추첨', '철제테이블(경품 실물 전시)', '5개', '이재중'],
  ['PRZ07', '경품·추첨', '수령증(1~3등 서명용), 당첨자 명단 서식', '각 5부', '양종철'],
  ['PRZ08', '경품·추첨', '사전행사 경품', '20명분', '이정하'],

  ['STG01', '무대·행사장', '전체 현수막(11m×2.75m)', '1개', '양종철'],
  ['STG02', '무대·행사장', '현수막(10m×1.2m)', '1개', '양종철'],
  ['STG03', '무대·행사장', '폼보드(75cm×35cm)', '2개', '이재중'],
  ['STG04', '무대·행사장', '접수처 배너 / 안내 배너', '각 1개', '채유리'],
  ['STG05', '무대·행사장', '물통 배너 받침(야외용)', '배너 수량만큼', '채유리'],
  ['STG06', '무대·행사장', '음향 장비', '1세트', '양종철'],
  ['STG07', '무대·행사장', '마이크', '실사용 2대(여유분 반입)', '양종철'],
  ['STG08', '무대·행사장', '단상', '2개', '고석우'],
  ['STG09', '무대·행사장', '태극기(협회 자체 준비)', '1개', '고석우'],
  ['STG10', '무대·행사장', '내빈 명패(좌석 라벨)', '필요분', '정승아'],
  ['STG11', '무대·행사장', '듀라테이블 / 천막', '14개 / 3개', '이재중'],
  ['STG12', '무대·행사장', '의자 / 천막', '20개 / 3개', '이재중'],
  ['STG13', '무대·행사장', '사회자 카드, 개회식 대본', '각 3부', '이정하'],
  ['STG14', '무대·행사장', '내빈 안내용지', '50장(양면)', '정승아'],

  ['BFR01', '부스-공정위원회', '공정위원회 현수막', '1개', '고석우'],
  ['BFR02', '부스-공정위원회', '공정위원회 배너세트', '1개', '고석우'],
  ['BFR03', '부스-공정위원회', '공정위원회 판넬', '9개', '고석우'],
  ['BFR04', '부스-공정위원회', '공정위원회 안내지', '50장', '고석우'],
  ['BFR05', '부스-공정위원회', '이슈파이팅 현수막(입구 나무 거치)', '5개', '고석우'],
  ['BFR06', '부스-공정위원회', 'SNS 이벤트·서명 용지, 필기구', '일괄', '고석우'],

  ['BYT01', '부스-청년위원회', '청년사회복지사 이벤트 물품', '1세트', '신단비'],
  ['BYT02', '부스-청년위원회', '부스 안내 표지(단체명 출력물)', '1매', '신단비'],

  ['BYM01', '부스-열매 만남존', '포토존 소품·안내물', '1세트', '정지연'],

  ['BKD01', '부스-어린이 이벤트', '화살표 머리띠(담당자 식별)', '1개', '최봄'],
  ['BKD02', '부스-어린이 이벤트', '안내 표지 및 전용 테이블', '1식', '최봄'],
  ['BKD03', '부스-어린이 이벤트', '법인 도장(수령 확인 날인)', '1개', '최지혜'],

  ['BTG01', '부스-함께하는 단체', '단체명 출력 용지(테이블 비치)', '단체별 1매', '신단비'],
  ['BTG02', '부스-함께하는 단체', '철제테이블(간식 배부용)', '3개', '이재중'],

  ['TRL01', '등반로 1·2·3지점', '상비약(밴드·연고·모기 기피제)', '지점별 1세트', '이해창, 유예리, 나한송'],
  ['TRL02', '등반로 1·2·3지점', '케이블타이(표지 거치용)', '넉넉히', '이해창'],
  ['TRL03', '등반로 1·2·3지점', '개인 물통·모자', '각자 지참', '등반팀 전원'],
  ['TRL04', '등반로 1·2·3지점', '쓰레기 회수 봉투·집게(하산 정리)', '코스별 1식', '이해창'],

  ['TRD01', '철수·정리', '쓰레기봉투(일반쓰레기 회수)', '넉넉히', '정소희'],
  ['TRD02', '철수·정리', '우비·비닐(우천 대비)', '확인 후 확보', '고석우'],
  ['TRD03', '철수·정리', '투명 패킹 박스(부스별 구분)', '부스별 1개 이상', '전체 직원'],

  ['ETC01', '기타', '비상연락망(부문별 담당자 연락처)', '전 직원 배부', '고석우'],
  ['ETC02', '기타', '세부 타임라인 출력본', '전 직원 배부', '고석우'],
  ['ETC03', '기타', '홍보부스 위치 표기용 출력 용지', '단체별 1매', '신단비'],
  ['ETC04', '기타', '구급용품', '일괄', '정소희'],
  ['ETC05', '기타', '행사 당일 단체 카톡방 개설(회장 포함)', '1개', '고석우']
];


/* ══════════════════════════════════════════════════════════════
   여기서부터 당일 안내(guide.html) 편집 API
   시트 4개: 당일시간블록 / 당일업무 / 개인별업무 / 개회식큐시트
   최초 1회 guideSeed 실행 필요
   ══════════════════════════════════════════════════════════════ */

var G_BLOCK = '당일시간블록';   // blockId | 시작 | 종료 | 블록명 | 순서
var G_TASK  = '당일업무';       // id | blockId | 장소 | 업무 | 총괄 | 지원 | 유의 | 순서
var G_SLOT  = '당일시간구간';   // 시작 | 종료  (개인별 매트릭스 열 기준)
var G_MTX   = '개인별업무';     // 이름 | 구간1..5
var G_CUE   = '개회식큐시트';   // 시간 | 순서 | 담당
var G_HOST  = '주관단체';       // id | 성명 | 소속 | 순서
var G_GUEST = '주요내빈';       // id | 구분 | 성명 | 소속 | 순서

var G_HEADERS = {};
G_HEADERS[G_BLOCK] = ['blockId', '시작', '종료', '블록명', '순서'];
G_HEADERS[G_TASK]  = ['id', 'blockId', '장소', '업무', '총괄', '지원', '유의', '순서'];
G_HEADERS[G_SLOT]  = ['시작', '종료'];
G_HEADERS[G_MTX]   = ['이름', '구간1', '구간2', '구간3', '구간4', '구간5'];
G_HEADERS[G_CUE]   = ['시간', '순서', '담당'];
G_HEADERS[G_HOST]  = ['id', '성명', '소속', '순서'];
G_HEADERS[G_GUEST] = ['id', '구분', '성명', '소속', '순서'];

/* ───────── 라우팅 (Code.gs 의 doGet 에서 위임) ───────── */

function guideRoute(p) {
  switch (p.action) {
    case 'guide':        return guideRead();
    case 'guideTaskSet': guard(p); return guideTaskSet(p);
    case 'guideTaskAdd': guard(p); return guideTaskAdd(p);
    case 'guideTaskDel': guard(p); return guideRowDel(G_TASK, 'id', p.id);
    case 'guideTaskMove':guard(p); return guideMove(G_TASK, p.id, p.dir);
    case 'guideBlockSet':guard(p); return guideBlockSet(p);
    case 'guideMtxSet':  guard(p); return guideMtxSet(p);
    case 'guideCueSet':  guard(p); return guideCueSet(p);
    case 'guideHostSet':  guard(p); return guidePersonSet(G_HOST, p);
    case 'guideHostAdd':  guard(p); return guidePersonAdd(G_HOST, p);
    case 'guideHostDel':  guard(p); return guideRowDel(G_HOST, 'id', p.id);
    case 'guideGuestSet': guard(p); return guidePersonSet(G_GUEST, p);
    case 'guideGuestAdd': guard(p); return guidePersonAdd(G_GUEST, p);
    case 'guideGuestDel': guard(p); return guideRowDel(G_GUEST, 'id', p.id);
    case 'guideGuestMove':guard(p); return guideMove(G_GUEST, p.id, p.dir);
    case 'guideSeed':    guard(p); return guideSeed(p.force === '1');
    default: return null;
  }
}

/* ───────── 공용 ───────── */

function gSheet(name) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    var h = G_HEADERS[name];
    sh.getRange(1, 1, 1, h.length).setValues([h])
      .setFontWeight('bold').setBackground('#1e7a4b').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function gRows(name) {
  var sh = gSheet(name);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var w = G_HEADERS[name].length;
  return sh.getRange(2, 1, last - 1, w).getValues()
    .filter(function (r) { return String(r[0]).length > 0; })
    .map(function (r) {
      return r.map(function (c) {
        return (c instanceof Date) ? hhmm(c) : String(c == null ? '' : c);
      });
    });
}

function gFindRow(name, keyCol, key) {
  var sh = gSheet(name);
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var idx = G_HEADERS[name].indexOf(keyCol);
  var v = sh.getRange(2, idx + 1, last - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]) === String(key)) return i + 2;
  return -1;
}

function gSetCells(name, row, patch) {
  var sh = gSheet(name);
  var h = G_HEADERS[name];
  var changed = [];
  for (var k in patch) {
    var c = h.indexOf(k);
    if (c > -1 && patch[k] !== undefined) {
      var cell = sh.getRange(row, c + 1);
      if (k === '시작' || k === '종료' || k === '시간') cell.setNumberFormat('@');
      cell.setValue(String(patch[k]));
      changed.push(k);
    }
  }
  return changed;
}

function guideRowDel(name, keyCol, key) {
  var row = gFindRow(name, keyCol, key);
  if (row < 0) throw new Error('항목을 찾을 수 없습니다: ' + key);
  gSheet(name).deleteRow(row);
  audit('안내 삭제', key, name, '');
  return { id: key, removed: true };
}

function guideMove(name, id, dir) {
  var sh = gSheet(name);
  var row = gFindRow(name, 'id', id);
  if (row < 0) throw new Error('항목을 찾을 수 없습니다: ' + id);
  var w = G_HEADERS[name].length;
  var target = String(dir) === 'up' ? row - 1 : row + 1;
  if (target < 2 || target > sh.getLastRow()) return { id: id, moved: false };
  /* 같은 블록 안에서만 이동 */
  var bi = G_HEADERS[name].indexOf('blockId') + 1;
  if (bi > 0 && sh.getRange(row, bi).getValue() !== sh.getRange(target, bi).getValue()) {
    return { id: id, moved: false };
  }
  var a = sh.getRange(row, 1, 1, w).getValues()[0];
  var b = sh.getRange(target, 1, 1, w).getValues()[0];
  sh.getRange(row, 1, 1, w).setValues([b]);
  sh.getRange(target, 1, 1, w).setValues([a]);
  return { id: id, moved: true };
}

/* ───────── 읽기 ───────── */

function guideRead() {
  var blocks = gRows(G_BLOCK), tasks = gRows(G_TASK);
  if (!blocks.length) return { empty: true };

  blocks.sort(function (a, b) { return (parseInt(a[4], 10) || 0) - (parseInt(b[4], 10) || 0); });

  var byBlock = {};
  tasks.forEach(function (t) {
    if (!byBlock[t[1]]) byBlock[t[1]] = [];
    byBlock[t[1]].push(t);
  });

  var tl = blocks.map(function (b) {
    var list = (byBlock[b[0]] || []).sort(function (x, y) {
      return (parseInt(x[7], 10) || 0) - (parseInt(y[7], 10) || 0);
    });
    var bs = hhmm(b[1]), be = hhmm(b[2]);
    return {
      id: b[0], s: bs, e: be, label: b[3],
      t: bs + '~' + be,
      items: list.map(function (t) {
        return {
          id: t[0], p: t[2], task: t[3],
          lead: splitNames(t[4]), sub: splitNames(t[5]), note: t[6]
        };
      })
    };
  });

  var slots = gRows(G_SLOT).map(function (r) { return [hhmm(r[0]), hhmm(r[1])]; });
  var mtx = {};
  gRows(G_MTX).forEach(function (r) { mtx[r[0]] = [r[1], r[2], r[3], r[4], r[5]]; });
  var cue = gRows(G_CUE).map(function (r) { return [hhmm(r[0]), r[1], r[2]]; });

  var hosts = gRows(G_HOST)
    .sort(function (a, b) { return (parseInt(a[3], 10) || 0) - (parseInt(b[3], 10) || 0); })
    .map(function (r) { return [r[0], r[1], r[2], r[3]]; });
  var guests = gRows(G_GUEST)
    .sort(function (a, b) { return (parseInt(a[4], 10) || 0) - (parseInt(b[4], 10) || 0); })
    .map(function (r) { return [r[0], r[1], r[2], r[3], r[4]]; });

  return { tl: tl, slots: slots, matrix: mtx, cue: cue, hosts: hosts, guests: guests };
}

/** 시트가 "07:00" 을 시각 값(Date)으로 저장한 경우에도 HH:mm 문자열로 되돌린다 */
function hhmm(v) {
  if (v instanceof Date) {
    return ('0' + v.getHours()).slice(-2) + ':' + ('0' + v.getMinutes()).slice(-2);
  }
  var t = String(v == null ? '' : v).trim();
  if (!t) return '';
  /* "Sat Dec 30 1899 07:00:00 GMT+0827" 같은 문자열 방어 */
  var m = t.match(/(\d{1,2}):(\d{2})/);
  if (m) return ('0' + m[1]).slice(-2) + ':' + m[2];
  return t;
}

function splitNames(v) {
  return String(v || '').split(/\s*,\s*/).filter(function (s) { return s.length; });
}

/* ───────── 쓰기 ───────── */

function guideTaskSet(p) {
  var row = gFindRow(G_TASK, 'id', p.id);
  if (row < 0) throw new Error('업무를 찾을 수 없습니다: ' + p.id);
  var patch = {};
  ['장소', '업무', '총괄', '지원', '유의'].forEach(function (k, i) {
    var src = ['p', 'task', 'lead', 'sub', 'note'][i];
    if (p[src] !== undefined) patch[k] = p[src];
  });
  var changed = gSetCells(G_TASK, row, patch);
  audit('안내 수정', p.id, changed.join(','), p.actor || '');
  return { id: p.id, changed: changed };
}

function guideTaskAdd(p) {
  var bid = String(p.blockId || '');
  if (!bid) throw new Error('시간대를 선택해 주세요');
  if (!String(p.task || '').trim()) throw new Error('업무 내용을 입력해 주세요');

  var sh = gSheet(G_TASK);
  var rows = gRows(G_TASK);
  var maxN = 0, insertAt = sh.getLastRow() + 1, order = 1;
  for (var i = 0; i < rows.length; i++) {
    var m = String(rows[i][0]).match(/^(.+)-(\d+)$/);
    if (m && m[1] === bid) maxN = Math.max(maxN, parseInt(m[2], 10));
    if (rows[i][1] === bid) { insertAt = i + 3; order = (parseInt(rows[i][7], 10) || 0) + 1; }
  }
  var id = bid + '-' + String(maxN + 1).padStart(2, '0');
  if (insertAt <= sh.getLastRow()) sh.insertRowBefore(insertAt);
  sh.getRange(insertAt, 1, 1, G_HEADERS[G_TASK].length).setValues([[
    id, bid, String(p.p || ''), String(p.task), String(p.lead || ''),
    String(p.sub || ''), String(p.note || ''), order
  ]]);
  audit('안내 추가', id, String(p.task), p.actor || '');
  return { id: id };
}

function guideBlockSet(p) {
  var row = gFindRow(G_BLOCK, 'blockId', p.id);
  if (row < 0) throw new Error('시간대를 찾을 수 없습니다: ' + p.id);
  var patch = {};
  if (p.s !== undefined) patch['시작'] = p.s;
  if (p.e !== undefined) patch['종료'] = p.e;
  if (p.label !== undefined) patch['블록명'] = p.label;
  var changed = gSetCells(G_BLOCK, row, patch);
  audit('안내 시간대 수정', p.id, changed.join(','), p.actor || '');
  return { id: p.id, changed: changed };
}

/** 주관단체·주요내빈 공용 수정 */
function guidePersonSet(name, p) {
  var row = gFindRow(name, 'id', p.id);
  if (row < 0) throw new Error('명단에서 찾을 수 없습니다: ' + p.id);
  var patch = {};
  if (name === G_GUEST && p.cat !== undefined) patch['구분'] = p.cat;
  if (p.name !== undefined) patch['성명'] = p.name;
  if (p.role !== undefined) patch['소속'] = p.role;
  var changed = gSetCells(name, row, patch);
  audit('내빈 수정', p.id, changed.join(','), p.actor || '');
  return { id: p.id, changed: changed };
}

/** 주관단체·주요내빈 공용 추가 (같은 구분 맨 아래에 끼워 넣음) */
function guidePersonAdd(name, p) {
  if (!String(p.name || '').trim()) throw new Error('성명을 입력해 주세요');
  var sh = gSheet(name);
  var rows = gRows(name);
  var pre = (name === G_HOST) ? 'H' : 'G';
  var max = 0;
  rows.forEach(function (r) {
    var m = String(r[0]).match(/^[A-Z](\d+)$/);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  var id = pre + ('0' + (max + 1)).slice(-2);

  var insertAt = sh.getLastRow() + 1, order = rows.length + 1;
  if (name === G_GUEST) {
    for (var i = rows.length - 1; i >= 0; i--) {
      if (rows[i][1] === String(p.cat)) { insertAt = i + 3; order = (parseInt(rows[i][4], 10) || 0) + 1; break; }
    }
  }
  if (insertAt <= sh.getLastRow()) sh.insertRowBefore(insertAt);

  var vals = (name === G_HOST)
    ? [id, String(p.name), String(p.role || ''), order]
    : [id, String(p.cat || ''), String(p.name), String(p.role || ''), order];
  sh.getRange(insertAt, 1, 1, G_HEADERS[name].length).setValues([vals]);
  audit('내빈 추가', id, String(p.name), p.actor || '');
  return { id: id };
}

function guideMtxSet(p) {
  var name = String(p.name || '');
  if (!name) throw new Error('이름이 없습니다');
  var sh = gSheet(G_MTX);
  var row = gFindRow(G_MTX, '이름', name);
  if (row < 0) {
    row = sh.getLastRow() + 1;
    sh.getRange(row, 1).setValue(name);
  }
  var i = parseInt(p.slot, 10);
  if (!(i >= 1 && i <= 5)) throw new Error('구간 번호가 잘못되었습니다');
  sh.getRange(row, i + 1).setValue(String(p.value == null ? '' : p.value));
  audit('안내 개인별 수정', name, '구간' + i, p.actor || '');
  return { name: name, slot: i };
}

function guideCueSet(p) {
  var i = parseInt(p.index, 10);
  var sh = gSheet(G_CUE);
  var row = i + 2;
  if (row < 2 || row > sh.getLastRow()) throw new Error('큐시트 행을 찾을 수 없습니다');
  if (p.time  !== undefined) { sh.getRange(row, 1).setNumberFormat('@').setValue(String(p.time)); }
  if (p.title !== undefined) sh.getRange(row, 2).setValue(String(p.title));
  if (p.who   !== undefined) sh.getRange(row, 3).setValue(String(p.who));
  audit('안내 큐시트 수정', 'CUE' + i, '', p.actor || '');
  return { index: i };
}

/* ───────── 시드 ───────── */

function guideSeed(force) {
  var sets = [
    [G_BLOCK, GUIDE_BLOCKS], [G_TASK, GUIDE_TASKS],
    [G_SLOT, GUIDE_SLOTS], [G_MTX, GUIDE_MATRIX], [G_CUE, GUIDE_CUE],
    [G_HOST, GUIDE_HOSTS], [G_GUEST, GUIDE_GUESTS]
  ];
  var out = {};
  sets.forEach(function (pair) {
    var name = pair[0], data = pair[1];
    var sh = gSheet(name);
    if (!force && sh.getLastRow() > 1) { out[name] = 'skip'; return; }
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, G_HEADERS[name].length).clearContent();
    }
    /* 07:00 같은 값이 시각으로 자동 변환되지 않도록 해당 열을 텍스트 서식으로 고정 */
    if (name === G_BLOCK) sh.getRange(2, 2, Math.max(data.length, 50), 2).setNumberFormat('@');
    if (name === G_SLOT)  sh.getRange(2, 1, Math.max(data.length, 50), 2).setNumberFormat('@');
    if (name === G_CUE)   sh.getRange(2, 1, Math.max(data.length, 50), 1).setNumberFormat('@');
    sh.getRange(2, 1, data.length, G_HEADERS[name].length).setValues(data);
    out[name] = data.length;
  });
  return out;
}

/** 강제 초기화 — 안내 시트를 코드의 기본값으로 되돌립니다 */
function guideSeedReset() { return guideSeed(true); }

/* ───────── 기본값 ───────── */

var GUIDE_BLOCKS = [
  ['B01', '07:00', '08:00', '집결·세팅', '1'],
  ['B02', '08:00', '08:30', '의전·선발대', '2'],
  ['B03', '08:30', '09:00', '접수 시작', '3'],
  ['B04', '09:00', '09:30', '개회식', '4'],
  ['B05', '09:30', '10:00', '경품·출발', '5'],
  ['B06', '10:00', '12:00', '등반', '6'],
  ['B07', '12:00', '13:00', '완주·정리', '7'],
  ['B08', '13:00', '16:00', '식사·복귀', '8']
];

var GUIDE_TASKS = [
  ['B01-01', 'B01', '벚꽃마당', '집결', '고석우', '전체 직원', '트럭 고석우 / 카니발 정승아(이정하) / 스파크 이지선(신단비, 채유리, 최봄)', '1'],
  ['B01-02', 'B01', '전체', '인력배치 · 스탭관리', '정승아', '이해창', '비상연락망 배부', '2'],
  ['B01-03', 'B01', '무대', '업체(음향) 관리', '양종철', '이재중', '음향 등 확인', '3'],
  ['B01-04', 'B01', '등반로', '시작점 배너 설치(3곳)', '이해창', '채유리', '물통 배너', '4'],
  ['B01-05', 'B01', '접수대', '접수대 세팅', '이진선, 정소희', '최봄, 최지혜, 유예리, 나한송, 손채은, 김도현', '테이블 4개 · 테이블별 주담당 이진선·정소희·최봄·최지혜', '5'],
  ['B01-06', 'B01', '접수대', '노트북 세팅 · 명단 조회 확인', '이진선, 정소희', '', '조회 4대 + 예비 1대', '6'],
  ['B01-07', 'B01', '접수대', '번호 표지 부착, 줄 유도선 설치', '최봄, 최지혜', '', '접수대 1·2·3·4 구분', '7'],
  ['B01-08', 'B01', '무대', '무대 세팅', '이지선', '고석우, 이정하, 정지연', '태극기·단상 2개·명패', '8'],
  ['B01-09', 'B01', '무대', '경품 실물 전시 세팅', '양종철, 이재중', '', '철제테이블, 포장 없이 실물 전시', '9'],
  ['B01-10', 'B01', '무대', '현수막 거치', '양종철', '이재중, 상비군', '', '10'],
  ['B01-11', 'B01', '부스', '청년위원회 부스 준비', '신단비', '', '', '11'],
  ['B01-12', 'B01', '전체', '김밥 도착(80줄) · 1인 1줄 배분', '고석우', '', '07:30~08:00 도착', '12'],
  ['B02-01', 'B02', '입구', '의전 및 의전관리', '정승아', '회장단', '부회장별 전담 의전', '1'],
  ['B02-02', 'B02', '입구', '선발대 출발 (08:15)', '이해창', '유예리, 김도현, 나한송, 손채은, 채유리', '1지점 유예리·김도현 / 2지점 나한송·손채은 / 3지점 이해창 / 배너설치 채유리', '2'],
  ['B02-03', 'B02', '등반로', '갈림길 배치', '이해창', '남성주, 송경태, 봉우석, 박소리, 황재우, 이세경, 배영미', '입구·주차 남성주·송경태 / 입구 갈림길 봉우석 / 등반 시작 입구 박소리 / 1~2지점 사이 황재우 / 2~3지점 사이 이세경·배영미', '3'],
  ['B02-04', 'B02', '무대 주변', '음향라인 통로 확보 (~10:00)', '양종철', '상비군', '입구 통로 지속 확보', '4'],
  ['B03-01', 'B03', '접수대', '접수대 운영 (4조 16명 · ~10:00)', '이진선, 정소희', '최봄, 최지혜, 연대회의', '조별 노트북 1 · 접수 1 · 기념품+물 2 — 1조 이진선 / 2조 최봄 / 3조 정소희 / 4조 최지혜, 각 조에 연대회의 배치', '1'],
  ['B03-02', 'B03', '접수대', '줄 안내', '연대회의, 권익위원회', '', '4~7명', '2'],
  ['B03-03', 'B03', '무대', '개회식 준비', '이지선', '고석우, 이정하', '사회 조은정 부위원장(권익위원회)', '3'],
  ['B03-04', 'B03', '무대', '사전행사 (08:30)', '이경원', '양종철, 이정하', '5천원권 20장 준비', '4'],
  ['B03-05', 'B03', '부스', '이슈파이팅 부스 (공정위원회)', '고석우', '', 'SNS 이벤트·서명', '5'],
  ['B03-06', 'B03', '부스', '함께하는 단체 관리 · 청년위원회 부스', '신단비', '', '', '6'],
  ['B04-01', 'B04', '무대', '국민의례 (09:00)', '조은정', '고석우, 이정하', '국기에 대한 경례만, 애국가 생략', '1'],
  ['B04-02', 'B04', '무대', '개회식 — 내빈소개·축사·환영사·연대사', '이지선', '고석우, 이정하, 이재중', '영상촬영 이재중 · 사회자 카드 참조', '2'],
  ['B04-03', 'B04', '무대', '수어 통역', '윤남', '', '', '3'],
  ['B04-04', 'B04', '전체', '사진촬영', '정지연', '김영민, 김진래, 김태웅', '정지연 열매존 근처 / 김영민·김진래·김태웅 개회식·스케치·전체', '4'],
  ['B05-01', 'B05', '무대', '경품 이벤트 (09:30~09:36)', '이경원', '양종철, 이정하', '오세훈 시장님 1등(1명) · 4등(약 100명) / 경품전달 이정하 · 수령증 서명 양종철', '1'],
  ['B05-02', 'B05', '무대', '기념촬영 (09:36~09:38)', '이지선', '', '단체사진', '2'],
  ['B05-03', 'B05', '무대→등반로', '주요 안내 및 등반 시작 (09:38~09:40)', '이지선', '', '안산 등반', '3'],
  ['B05-04', 'B05', '부스구역(원형광장)', '어린이 이벤트 · 문화상품권 5천원권', '최봄', '최지혜', '개회식 종료 후 · 담당자 머리띠 착용 · 수령자 손목 도장', '4'],
  ['B06-01', 'B06', '1지점 숲속무대', '스탬프 1지점 운영', '유예리, 김도현', '', '지점 표지 게시 · 도장 5개', '1'],
  ['B06-02', 'B06', '2지점 전망대', '스탬프 2지점 운영', '나한송, 손채은', '', '도장 5개 · 팀 하산 안내 병행', '2'],
  ['B06-03', 'B06', '3지점 너와집쉼터', '스탬프 3지점 운영', '이해창, 채유리', '', '도장 5개', '3'],
  ['B06-04', 'B06', '등반로 전 구간', '갈림길 안내', '이해창', '박소리, 배영미, 이세경, 황재우, 상비군', '협회 수건·깃발 · 초입 집중 배치', '4'],
  ['B06-05', 'B06', '등반로', '등반 스케치 촬영', '김영민, 김진래, 김태웅', '', '', '5'],
  ['B06-06', 'B06', '부스구역', '청년위원회 부스 운영', '신단비', '', '', '6'],
  ['B06-07', 'B06', '무대', '무대 대기 · 내빈 응대', '이지선, 정승아', '', '', '7'],
  ['B06-08', 'B06', '부스구역', '아이스크림 현장 도착 (11:30)', '최봄', '', '공제회 후원 1,000개 · 아이스박스', '8'],
  ['B07-01', 'B07', '부스구역(원형광장)', '아이스크림 배분', '최지혜', '최봄, 잔여인력', '설레임 1,000개(11:30 현장 도착)', '1'],
  ['B07-02', 'B07', '부스구역(원형광장)', '완주 확인', '이진선', '이재중, 잔여인력', '스탬프 3개 확인 후 용지 회수', '2'],
  ['B07-03', 'B07', '부스구역(원형광장)', 'QR 등록 안내', '이정하', '신단비, 잔여인력', 'QR 출력물 게시 · 완주 확인자만 개인정보 입력', '3'],
  ['B07-04', 'B07', '종료지점', '등반코스 종료지점 샛길 안내 (11:30)', '정소희, 정지연', '', '배너 챙기기', '4'],
  ['B07-05', 'B07', '접수대~숲속무대', '등반코스 안내 표식 철거 (11:00)', '이재중, 신단비', '', '', '5'],
  ['B07-06', 'B07', '행사장', '정리 및 청소', '양종철', '고석우, 정승아, 봉우석, 송경태, 김영민, 김진래', '분리수거는 행사장 배출 · 일반쓰레기는 협회로 회수', '6'],
  ['B07-07', 'B07', '무대 앞', '내빈 안내', '이지선, 정승아', '회장단', '', '7'],
  ['B07-08', 'B07', '벚꽃마당', '참가자 해산 (12:00~12:30)', '전체', '', '', '8'],
  ['B08-01', 'B08', '이동', '식사 장소 이동 (13:00)', '고석우', '전체 직원', '도보 13분 · 677m', '1'],
  ['B08-02', 'B08', '연탄생고기집 홍은점', '진행요원 식사 (13:30~15:30)', '고석우', '전체 직원', '', '2'],
  ['B08-03', 'B08', '협회', '복귀 · 물품 정리 (15:30~16:00)', '고석우', '전체 직원', '차량 주차, 트럭 보관 · 트럭은 일요일 반납이라 당일 전량 하차', '3']
];

var GUIDE_SLOTS = [
  ['07:00', '08:00'],
  ['08:00', '10:00'],
  ['10:00', '12:00'],
  ['12:00', '13:00'],
  ['13:00', '16:00']
];

var GUIDE_MATRIX = [
  ['이지선', '무대 세팅', '의전(입구) / 개회식 준비·진행', '무대 대기', '내빈 안내', '스파크 운전 / 식사·물품 정리'],
  ['정승아', '인력배치·스탭관리', '의전 / 내빈 전담 배정 확인', '내빈 응대', '내빈 안내', '카니발 운전 / 식사·물품 정리'],
  ['고석우', '집결 총괄', '개회식 지원 / 이슈파이팅 부스(공정위원회)', '등반로 연락 총괄', '정리·청소', '트럭 운전 / 식사 인솔·물품 정리'],
  ['이정하', '무대 세팅', '사전행사·개회식 / 경품 이벤트 지원', '-', '완주 확인·QR 등록', '식사·물품 정리'],
  ['양종철', '음향업체 관리', '음향라인 통로 확보 / 경품 실물 전시 / 현수막 거치 / 수령증 서명', '-', 'QR 등록 안내 / 정리·청소', '식사·물품 정리'],
  ['이재중', '업체 관리 지원', '경품 실물 전시 / 현수막 거치 / 개회식 영상 촬영', '-', '등반코스 안내 표식 철거', '식사·물품 정리'],
  ['이해창', '시작점 배너 설치', '스탭관리 지원 / 선발대 출발 인솔', '스탬프 3지점 / 갈림길 안내 총괄', '철수 합류', '식사·물품 정리'],
  ['채유리', '시작점 배너 설치', '선발대', '스탬프 3지점 운영', '철수 합류', '식사·물품 정리'],
  ['신단비', '청년위원회 부스 준비', '함께하는 단체 관리 / 청년위원회 부스', '청년위원회 부스 운영', '등반코스 안내 표식 철거', '식사·물품 정리'],
  ['이진선', '접수대 세팅 / 노트북·명단 조회', '접수대 1 총괄·운영 / 경품 배부', '-', '완주 확인(스탬프 확인·용지 회수)', '식사·물품 정리'],
  ['정소희', '접수대 세팅 / 노트북 세팅', '접수대 2 총괄·운영 / 현장접수 처리 / 경품 배부', '-', '종료지점 샛길 안내', '식사·물품 정리'],
  ['최봄', '접수대 세팅 / 번호 표지·줄 유도선', '접수대 접수 / 어린이 이벤트', '어린이 이벤트 · 아이스크림 도착 대기(11:30)', '기념품·경품 배부 지원', '식사·물품 정리'],
  ['최지혜', '접수대 세팅 / 번호 표지·줄 유도선', '접수대 접수 / 어린이 이벤트 지원', '-', '아이스크림 배분', '식사·물품 정리'],
  ['유예리', '접수대 세팅', '선발대', '스탬프 1지점 운영', '철수 합류', '식사·물품 정리'],
  ['정지연', '무대 세팅', '행사 사진 촬영 총괄(열매존 인근)', '-', '종료지점 샛길 안내', '식사·물품 정리'],
  ['나한송', '접수대 세팅', '선발대', '스탬프 2지점 운영', '철수 합류', '식사·물품 정리'],
  ['손채은', '접수대 세팅', '선발대', '스탬프 2지점 운영', '철수 합류', '식사·물품 정리'],
  ['김도현', '접수대 세팅', '선발대', '스탬프 1지점 운영', '철수 합류', '식사·물품 정리'],
  ['조은정', '-', '개회식 사회 / 국민의례 진행 / 경품 이벤트', '-', '-', '-'],
  ['허곤', '-', '내빈 맞이', '-', '-', '-'],
  ['홍영호', '-', '선발대 기수(1지점까지 인솔)', '-', '-', '-'],
  ['구본영', '-', '접수대 협조', '-', '완주 확인·기념품 지원', '-'],
  ['안보현', '-', '접수대 협조', '-', '완주 확인·기념품 지원', '-'],
  ['태혜영', '-', '접수대 협조', '-', '완주 확인·기념품 지원', '-'],
  ['박소리', '-', '선발대', '갈림길 안내', '-', '-'],
  ['배영미', '-', '선발대', '갈림길 안내', '-', '-'],
  ['이세경', '-', '선발대', '갈림길 안내', '-', '-'],
  ['황재우', '-', '선발대', '갈림길 안내', '-', '-'],
  ['이경원', '-', '사전행사 진행 / 경품 이벤트 진행', '-', '-', '-'],
  ['김영민', '-', '개회식·스케치·전체 촬영(시장·주요내빈 스냅)', '등반 스케치 촬영', '정리 지원', '-'],
  ['김진래', '-', '개회식·전체·단체사진 촬영', '등반 스케치 촬영', '정리 지원', '-'],
  ['김태웅', '-', '회원·스케치 촬영', '등반 스케치 촬영', '-', '-'],
  ['윤남', '-', '수어 통역', '-', '-', '-'],
  ['상비군', '현수막 거치', '음향라인 통로 확보 지원', '갈림길 안내', '등반코스 안내 표식 철거 / 청소', '-'],
  ['회장단', '-', '회원 맞이 · 내빈 의전(관리사무소 앞 대기)', '-', '내빈 안내', '-'],
  ['연대회의', '접수대 세팅', '접수인원 관리 / 기념품·경품 배부', '-', '아이스크림 배부', '-']
];

var GUIDE_CUE = [
  ['09:00', '개회 및 내빈소개', '사회자 조은정 (권익위원회 부위원장) · 5분'],
  ['09:05', '개회사', '곽경인 (서울시사회복지사협회장) · 5분'],
  ['09:10', '시장님 축사', '오세훈 서울특별시장 · 5분'],
  ['09:15', '축사', '남인순 (국회 부의장) · 김경우 (서울시의회 보건복지위원장) · 6분'],
  ['09:21', '환영사', '박운기 (서대문구청장) · 3분'],
  ['09:24', '연대사', '김연은 (서울시사회복지시설연대회의 상임대표) · 3분'],
  ['09:27', '회원축사', '한양호, 박소영 (청년사회복지사) · 3분'],
  ['09:30', '경품추첨', '오세훈 시장님 1등(1명)·4등(약 100명) / 조은정 부위원장, 이경원 대표 진행 · 6분'],
  ['09:36', '기념촬영', '단체사진 · 2분'],
  ['09:38', '안내 및 등반', '주요 안내 및 안산 등반 · 2분']
];

var GUIDE_HOSTS = [
  ['H01', '곽경인', '사회복지사협회장 · 성수종합사회복지관장', '1'],
  ['H02', '김연은', '서울시 사회복지단체연대회의 회장 · 생명의전화종합사회복지관장', '2']
];

var GUIDE_GUESTS = [
  ['G01', '서울시·국회·구청', '오세훈', '서울특별시장', '1'],
  ['G02', '서울시·국회·구청', '남인순', '국회 부의장 (더불어민주당, 송파구병) ※협회 회원', '2'],
  ['G03', '서울시·국회·구청', '김영호', '국회의원 (더불어민주당, 서대문구을)', '3'],
  ['G04', '서울시·국회·구청', '박운기', '서대문구 구청장', '4'],
  ['G05', '서울시의회', '김경우', '보건복지위원회 위원장', '5'],
  ['G06', '서울시의회', '이병도', '운영위원회 위원장 · 보건복지위원회 위원', '6'],
  ['G07', '서울시의회', '이승미', '보건복지위원회 위원', '7'],
  ['G08', '서울시의회', '함대건', '도시안전건설위원회 부위원장', '8'],
  ['G09', '유관기관·전임회장', '김병민', 'G3서울기획위원회 위원장', '9'],
  ['G10', '유관기관·전임회장', '장재구', '제11·12대 회장 (중앙사회복지관 관장)', '10'],
  ['G11', '유관기관·전임회장', '심정원', '제14·15대 회장 (성산종합사회복지관 관장)', '11'],
  ['G12', '유관기관·전임회장', '조남범', '서울시사회복지협의회 회장', '12'],
  ['G13', '유관기관·전임회장', '박병삼', '서울시사회복지행정연구회 회장', '13'],
  ['G14', '중앙협회', '조석영', '한국장애인복지관협회 회장', '14'],
  ['G15', '중앙협회', '조범기', '한국시니어클럽협회 회장', '15'],
  ['G16', '공동주최 단체 공동대표', '엄종숙', '서울시장애인복지시설협회 회장', '16'],
  ['G17', '공동주최 단체 공동대표', '최성남', '서울시정신재활시설협회 회장', '17'],
  ['G18', '공동주최 단체 공동대표', '김남용', '서울구립노인복지관협회 회장', '18'],
  ['G19', '서울시 사회복지 직능단체', '곽금봉', '서울시노인복지협회 회장', '19'],
  ['G20', '서울시 사회복지 직능단체', '신재원', '서울시노인복지관협회 회장', '20'],
  ['G21', '서울시 사회복지 직능단체', '장현준', '서울시재가노인복지협회 회장', '21'],
  ['G22', '서울시 사회복지 직능단체', '이소영', '서울시아동복지협회 회장', '22'],
  ['G23', '서울시 사회복지 직능단체', '권수정', '서울시한부모가족복지시설협회 회장', '23'],
  ['G24', '서울시 사회복지 직능단체', '최선자', '서울시장애인복지관협회 회장', '24'],
  ['G25', '서울시 사회복지 직능단체', '홍금화', '서울시장애인주간보호단기거주시설협회 회장', '25'],
  ['G26', '서울시 사회복지 직능단체', '이민규', '서울시장애인직업재활시설협회 회장', '26'],
  ['G27', '서울시 사회복지 직능단체', '장경환', '서울노숙인시설협회 회장', '27'],
  ['G28', '서울시 사회복지 직능단체', '이선화', '서울지역자활센터협회 회장', '28'],
  ['G29', '서울시 사회복지 직능단체', '김은영', '서울시지역아동센터협의회 회장', '29'],
  ['G30', '서울시 사회복지 직능단체', '최유연', '서울시여성폭력피해지원시설협의회 대표', '30'],
  ['G31', '서울시 사회복지 직능단체', '한미영', '서울시가족센터협회 회장', '31'],
  ['G32', '서울시 사회복지 직능단체', '이혜경', '서울시아동청소년그룹홈협의회 회장', '32'],
  ['G33', '서울시 사회복지 직능단체', '이율기', '서울시니어클럽협회 회장', '33'],
  ['G34', '서울시 사회복지 직능단체', '문기덕', '서울시아동보호전문기관협회 회장', '34'],
  ['G35', '서울시 사회복지 직능단체', '최재옥', '서울시정신요양시설협회 회장', '35'],
  ['G36', '서울시 사회복지 직능단체', '신건철', '서울시발달장애인평생교육센터협의회 회장', '36'],
  ['G37', '서울시 사회복지 직능단체', '조한종', '서울시50플러스센터협의회 회장', '37'],
  ['G38', '서울시 사회복지 직능단체', '권현수', '서울시우리동네키움센터협의회 회장', '38'],
  ['G39', '회장단', '강현덕', '영등포구가족센터 센터장', '39'],
  ['G40', '회장단', '송주혜', '서울시립뇌성마비복지관 관장', '40'],
  ['G41', '회장단', '허곤', '더홈 원장', '41'],
  ['G42', '회장단', '안진경', '동부외국인주민센터 센터장', '42'],
  ['G43', '회장단', '조민혜', '관악봉천지역자활센터 센터장', '43'],
  ['G44', '회장단', '김아래미', '서울여자대학교 교수', '44'],
  ['G45', '회장단', '송향숙', '서울시립성북노인종합복지관 관장', '45'],
  ['G46', '회장단', '한양호', '마포장애인종합복지관 팀장', '46'],
  ['G47', '함께하는 단체', '이재홍', '사회복지종사자 권익지원센터 센터장', '47'],
  ['G48', '함께하는 단체', '박유빈', '직장갑질119 온라인노조 사회복지지부 지부장 직무대행', '48']
];
