# 등반대회 준비물 3단 체크리스트

2026 서울사회복지사 등반대회(2026. 9. 12.) 준비물 72항목을
**9. 10. 패킹 → 9. 11. 최종 확인 → 9. 12. 현장 반입** 3단계로 나눠 체크하는 웹앱입니다.

- 프런트: 정적 HTML 1개 (빌드 없음) → Vercel
- 백엔드: Google Apps Script 웹앱 → Google 스프레드시트
- 통신: JSONP (CORS 설정 불필요)

```
checklist/
├─ public/
│  ├─ index.html      화면 전체 (HTML+CSS+JS)
│  ├─ config.js       ★ GAS 웹앱 URL 입력
│  └─ items.js        준비물 마스터 72항목 (오프라인 폴백)
├─ apps-script/
│  └─ Code.gs         시트 API + 시드
├─ vercel.json
└─ README.md
```

## 1단계 — Apps Script 배포

1. <https://script.google.com> 에서 새 프로젝트를 만들고 `apps-script/Code.gs` 내용을 붙여넣습니다.
2. 상단에서 함수 `seed` 를 선택하고 실행합니다.
   - 스프레드시트에 `준비물체크리스트` 시트가 생기고 72항목이 입력됩니다.
   - 최초 실행 시 스프레드시트 접근 권한 승인이 필요합니다.
3. **배포 → 새 배포 → 유형: 웹 앱**
   - 설명: `등반대회 체크리스트 API`
   - 실행: **나(본인 계정)**
   - 액세스 권한: **모든 사용자**
4. 발급된 URL(`.../exec`)을 복사합니다.

> 시트 ID는 `Code.gs` 상단 `SHEET_ID` 에 이미 넣어 두었습니다.
> `1PQy4OPfn3ldPuudmJwKzJtiOm8itfUPEpORyYqWfXlo`

## 2단계 — URL 연결

`public/config.js` 한 줄만 고칩니다.

```js
window.CHECKLIST_API = "https://script.google.com/macros/s/AKfy.../exec";
```

비워 두면 데모 모드로 열립니다(체크가 저장되지 않고 상단에 노란 띠가 표시됩니다).

## 3단계 — GitHub → Vercel

```bash
cd checklist
git init
git add .
git commit -m "등반대회 준비물 3단 체크리스트"
git branch -M main
git remote add origin https://github.com/<계정>/climb-checklist.git
git push -u origin main
```

Vercel에서 New Project → 이 저장소 선택 → 프레임워크는 **Other**,
`vercel.json`이 `outputDirectory: public`을 잡아 주므로 나머지는 기본값으로 배포합니다.

CLI로 바로 올리려면:

```bash
npx vercel --prod
```

## 화면 사용법

| 기능 | 설명 |
|---|---|
| 상단 3개 탭 | 확인 단계 전환. 단계마다 색이 달라 지금 무엇을 체크하는지 헷갈리지 않습니다 (패킹=파랑 / 확인=주황 / 당일=빨강) |
| 구분별 / 사람별 | 13개 구분으로 보거나, 18명 담당자별로 봅니다. 사람별에서는 남은 항목이 많은 사람이 위로 올라옵니다 |
| 담당자 선택 | 특정 담당자 것만 필터. 진행률도 그 사람 기준으로 다시 계산됩니다 |
| 남은 것만 | 현재 단계에서 아직 체크 안 된 항목만 표시 |
| 확인한 사람 | 이름을 입력하면 체크할 때 시트 `0910확인자` 등 칸에 함께 기록됩니다 |
| 행 오른쪽 작은 점 | 다른 두 단계의 체크 상태. 9/11 화면에서도 9/10에 챙겼는지 바로 보입니다 |

주소창 해시에 상태가 남아서(`#me=고석우&s=1&v=person`) 그대로 공유하면
상대방이 같은 화면으로 열립니다. 담당자별 링크를 카톡으로 뿌릴 때 편합니다.

## 시트 구조

`준비물체크리스트` 시트

| 열 | 내용 |
|---|---|
| A | id (RCP01 …) — 절대 바꾸지 마세요. 체크 저장 키입니다 |
| B~E | 구분 / 준비물 / 수량 / 담당 |
| F~H | 0910 / 0911 / 0912 체크박스 |
| I~K | 각 단계 확인자 |
| L | 갱신시각 |
| M | 메모 |

`체크로그` 시트에 체크·해제 이력이 시각·항목·확인자와 함께 누적됩니다.
당일 누가 언제 확인했는지 확인할 때 씁니다.

### 항목을 고칠 때

시트에서 B~E 열을 직접 고쳐도 됩니다.
항목을 추가·삭제할 때는 `apps-script/Code.gs`의 `ITEMS` 와
`public/items.js` 의 `CHECKLIST_ITEMS` 를 **양쪽 모두** 수정한 뒤
스프레드시트 메뉴 `등반대회 체크리스트 → 항목 다시 쓰기(체크 유지)` 를 실행하세요.
기존 체크 상태는 id 기준으로 유지됩니다.

## API

| 호출 | 결과 |
|---|---|
| `?action=list` | 전체 항목 + 체크 상태 |
| `?action=set&id=RCP01&stage=0910&value=1&actor=고석우` | 체크 저장 |
| `?action=note&id=RCP01&text=...` | 메모 저장 |
| `?action=stats` | 단계별 완료 수 |
| `?action=seed` / `&force=1` | 항목 다시 쓰기 / 전체 초기화 |

모든 호출에 `&callback=fn` 을 붙이면 JSONP로 응답합니다.
