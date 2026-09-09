/**
 * 2026 서울사회복지사 등반대회 — 준비물 3단 체크리스트 API
 *
 * 배포 방법
 *  1) script.google.com 새 프로젝트 → 이 파일 내용 붙여넣기
 *  2) 실행 → seed  (시트 생성 + 72항목 입력, 최초 1회)
 *  3) 배포 → 새 배포 → 유형: 웹 앱
 *       - 실행: 나(본인)
 *       - 액세스: 모든 사용자
 *  4) 발급된 /exec URL 을 public/config.js 의 CHECKLIST_API 에 붙여넣기
 *
 * 프런트엔드는 JSONP(callback 파라미터)로 호출하므로 CORS 설정이 필요 없습니다.
 */

var SHEET_ID   = '1PQy4OPfn3ldPuudmJwKzJtiOm8itfUPEpORyYqWfXlo';
var SHEET_NAME = '준비물체크리스트';
var LOG_NAME   = '체크로그';

var HEADERS = [
  'id', '구분', '준비물', '수량', '담당',
  '0910', '0911', '0912',
  '0910확인자', '0911확인자', '0912확인자',
  '갱신시각', '메모'
];

var STAGE_COL = { '0910': 6, '0911': 7, '0912': 8 };   // 1-based
var ACTOR_COL = { '0910': 9, '0911': 10, '0912': 11 };
var TIME_COL  = 12;

/* ────────────────────────── 라우팅 ────────────────────────── */

function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try {
    switch (p.action) {
      case 'seed':  out = { ok: true, data: seedRows(p.force === '1') }; break;
      case 'set':   out = { ok: true, data: setCheck(p.id, p.stage, p.value === '1', p.actor || '') }; break;
      case 'note':  out = { ok: true, data: setNote(p.id, p.text || '') }; break;
      case 'stats': out = { ok: true, data: stats() }; break;
      default:      out = { ok: true, data: listRows() };
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
  try {
    out = { ok: true, data: setCheck(p.id, p.stage, !!p.value, p.actor || '') };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return reply(out, null);
}

function reply(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
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
    sh.getRange(1, 1, 1, 5).setValues([['시각', 'id', '단계', '값', '확인자']])
      .setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
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
      id:   String(r[0]),
      cat:  String(r[1]),
      name: String(r[2]),
      qty:  String(r[3]),
      who:  String(r[4]),
      done: [truthy(r[5]), truthy(r[6]), truthy(r[7])],
      by:   [String(r[8] || ''), String(r[9] || ''), String(r[10] || '')],
      memo: String(r[12] || '')
    });
  }
  return out;
}

function truthy(v) {
  if (v === true) return true;
  var s = String(v).trim().toUpperCase();
  return s === 'TRUE' || s === 'O' || s === 'Y' || s === '1' || s === 'V' || s === '✔';
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
    sh.getRange(row, STAGE_COL[stage]).setValue(value ? true : false);
    sh.getRange(row, ACTOR_COL[stage]).setValue(value ? (actor || '') : '');
    sh.getRange(row, TIME_COL).setValue(now);

    logSheet().appendRow([now, id, stage, value ? 'O' : '-', actor || '']);
    return { id: id, stage: stage, value: value, actor: actor || '' };
  } finally {
    lock.releaseLock();
  }
}

function setNote(id, text) {
  var sh = sheet();
  var row = rowIndexById(sh, id);
  if (row < 0) throw new Error('항목을 찾을 수 없습니다: ' + id);
  sh.getRange(row, 13).setValue(text);
  return { id: id, memo: text };
}

function stats() {
  var rows = listRows();
  var out = { total: rows.length, stages: {} };
  ['0910', '0911', '0912'].forEach(function (k, i) {
    out.stages[k] = rows.filter(function (r) { return r.done[i]; }).length;
  });
  return out;
}

/* ────────────────────────── 시드 ────────────────────────── */

/** 메뉴에서 직접 실행: 항목만 새로 쓰고 체크 상태는 유지 */
function seed()      { return seedRows(false); }
/** 강제 초기화: 체크 상태까지 모두 지움 */
function seedReset() { return seedRows(true); }

function seedRows(force) {
  var sh = sheet();
  var keep = {};
  if (!force) {
    listRows().forEach(function (r) { keep[r.id] = r; });
  }
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, HEADERS.length).clearContent();

  var values = ITEMS.map(function (it) {
    var k = keep[it[0]];
    return [
      it[0], it[1], it[2], it[3], it[4],
      k ? k.done[0] : false, k ? k.done[1] : false, k ? k.done[2] : false,
      k ? k.by[0] : '',      k ? k.by[1] : '',      k ? k.by[2] : '',
      '', k ? k.memo : ''
    ];
  });
  sh.getRange(2, 1, values.length, HEADERS.length).setValues(values);
  sh.getRange(2, 6, values.length, 3).insertCheckboxes();
  return { rows: values.length, kept: Object.keys(keep).length };
}

/** 스프레드시트를 열었을 때 메뉴 추가 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('등반대회 체크리스트')
    .addItem('항목 다시 쓰기(체크 유지)', 'seed')
    .addItem('전체 초기화(체크까지 삭제)', 'seedReset')
    .addToUi();
}

/* ────────────────────────── 준비물 마스터(72항목) ────────────────────────── */
/* public/items.js 와 동일한 내용입니다. 수정 시 양쪽 모두 반영하세요. */
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
