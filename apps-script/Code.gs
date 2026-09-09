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

      default: out = ok(listRows());
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
