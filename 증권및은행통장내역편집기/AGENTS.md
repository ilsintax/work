# AGENTS.md — 통합 통장내역 편집기

- 프로젝트 루트: 이 문서가 있는 폴더 (`developement/TEST/증권및은행통장내역편집기`)
- 적용 지침 (사본은 `instructions/`, 공통 원본은 `../../지침/`)

| 파일 | 버전 | SHA256 | 적용 범위 |
|---|---|---|---|
| instructions/AGENTS_EXCEL_PROJECT_FINAL.md | 2.1 (2026-11-03) | 3ff6eaa45e73c939b9165ad3e508d605e17c8fafef4ae25af33d191b50a00296 | 주 지침: 엑셀 입력·정제·검증·저장 전반 |
| instructions/AGENTS_HTML_EXCEL_UI_FINAL.md | 2.1 (2026-11-03) | 9af3577d45bbc55d086ae6d3215183c8934a0cc348600b32d152df66a95aed19 | 보조: 업로드·미리보기·다운로드 화면 규칙만 |

## 읽는 순서
적용 지침 → `개발구성.md` → `개선제안서.md` → `HISTORY.md` 최근 기록 → `HANDOFF.md`

## 프로젝트 전용 규칙 (사용자 반복 지시)
1. `!업무허브`의 기존 도구(stockbalance, 통장적요자동생성기)와 기존 작업물은 절대 수정하지 않는다. 필요한 코드는 `reference/`의 복사본에서만 읽고 가져온다.
2. 최종 결과 형식은 수지라 업로드용을 유지한다: `결제일 / 입금 / 출금 / 잔액 / (참고용)잔액 / 의뢰인/수취인`.
3. 증권 수수료·제세금은 지금처럼 별도 출금 행으로 나눈다(기본값 유지).
4. 기관사전(JSON)은 이 폴더 `dictionary/`에 둔다.
5. 거래내역이 아닌 파일(예: 잔고증명)도 자동 제외하지 않는다. 판별 결과를 표시하고 사용자가 결정한다.
6. 은행·증권 여부, 기관, 계좌를 사용자가 파일을 열어 확인하지 않아도 되도록 자동 판별을 우선한다. 판별 실패·잔액 불일치만 사용자 검토로 보낸다.
7. 계좌번호는 필수가 아니다. 계좌 이름은 파일명(시트가 여러 개면 파일명 + 시트명)으로 단다(기존 도구 방식). 계좌명은 화면에서 고칠 수 있다.
8. 목표는 "모든 계좌를 한 화면에서 편집하는 도구"다. 자동 결과를 계좌별로 열 지정·기초잔액·분류·적요까지 고칠 수 있어야 한다.
9. 입출금에 영향을 주지 않는 내역(주식 입고·출고, 제거 항목, 미분류)은 버리지 않고 "입출금외 내역"으로 따로 보관한다(주식 입고 여부 등 확인용).
10. 입출금외 내역의 적요에는 수량 × 거래단가 금액을 계산해 붙인다. 예: `공모주입고 / 브릴스 / 46주 / @19500 / 잔고46주 / 897,000원`
11. 증권 변환 로직은 기존 도구 최종본을 `tools/extract-securities-engine.js`로 그대로 잘라 쓴다. `engine/securities-legacy.js`는 직접 고치지 않는다.
12. GitHub 저장소(GitHub Pages)로 팀 전체가 쓴다. 작업 내용은 새로고침하면 사라지는 것이 정상(서버 저장 없음, 엑셀은 브라우저 안에서만 처리).
13. 사전 반영은 브라우저 메모리에만 적용 → "사전 파일 저장"으로 institutions.js 를 받아 관리자가 저장소에 올리는 방식. 사전에는 기관 단위 지식만 저장하고 계좌번호·고객명·금액은 넣지 않는다.
14. 고객 데이터(샘플 엑셀, 샘플별 기준값 tests/expected*.json, version/)는 저장소에 올리지 않는다(.gitignore). 올리기 전 `node tools/check-publish.js` 통과 필수.
15. 증권사마다 처리 방식이 다르므로 증권사별 처리 방식(형식·분류·수수료/비용 처리·추가 적요·특수 규칙·검증 방식·외화 처리)을 사전에 증권사 단위로 기억한다. 외화는 같은 계좌라도 국내와 다른 열·분류를 쓰므로 해외 사전에 따로 기억한다.
16. 열 이름(제목) 매칭으로 다른 증권사·은행으로 잘못 잡히는 일이 없도록 보수적으로 판별한다. 확실하지 않으면 추측하지 말고 "확인 필요"로 보내 사용자가 고르게 한다.
17. 운영 방식: 사용자가 예외가 나온 증권사·은행 계좌 파일을 올리면 Claude 가 분석해 그 기관의 처리 방식을 사전에 기록한다(아래 "예외 계좌 분석 절차"). 화면의 "사전 반영"은 보조 수단.

## 백업
- 변경 전 백업 위치: `version/YYYYMMDD_HHMMSS_작업명/` (files/, manifest.json, RESTORE.md, VERIFIED)
- 백업 명령(대체 방식 자동화, Git Bash): `bash tools/backup.sh <작업명> <작업ID> <파일...>`
- 공통 도구 `tools/Backup-Project.ps1`은 이 폴더(SynologyDrive, 재분석 지점)에서 `Linked root is unsupported`로 실행되지 않는다.
  대체 방식: 변경 대상 파일을 `cp -p`로 files/에 복사 → `sha256sum`으로 원본·사본 비교 → manifest.json·RESTORE.md 작성 → 전부 일치할 때만 VERIFIED 생성. (예: `version/20261110_174529_2단계문서갱신전/`)
- 샘플 엑셀(루트의 *.xlsx, *.xls)은 사용자 원본 데이터다. 프로그램은 읽기만 하며 테스트가 원본 수정 시각 변화를 검사한다.
- 백업 제외: `node_modules/`(npm install로 재생성), `tests/output/`(테스트 산출물)

## 실행·검증
- 실행: `index.html`을 Edge/Chrome에서 연다. (`lib/`, `dictionary/`, `engine/` 폴더가 같은 위치에 있어야 함)
- 개발 환경 준비(새 PC): Node.js LTS 설치 → 루트에서 `npm install`
- 검증 1 (엔진): 루트에서 `npm test` — 샘플 24개 시트를 `tests/expected.json`과 비교. 상세: `npm test -- -v`
- 검증 2 (변환): 루트에서 `node tests/run-convert.js` — 샘플 전체 변환 후 계좌별 상태·결과 행수·기초/기말 잔액을 `tests/expected-convert.json`과 비교. 상세: `-v`
- 검증 3 (화면): 루트에서 `node tests/browser-check.js` — Edge를 화면 없이 띄워 업로드 → 상태 비교 → KB증권 수수료 포함 편집 → IM 열 직접 지정 → 다운로드 파일 시트 확인. 스크린샷 `tests/output/browser-*.png`
- 검증 5 (1단계 보수적 판별·검증 강도): `node tests/run-verification.js` — 거짓 자동 완료 방지, 파일명 없음/충돌/열 1개 다름 판별, 확인 완료·요약 시트, 기초잔액 입력
- 검증 6 (2단계 사전 구조): `node tests/run-structure.js` — 구조 1→2 변환, 사전 검사·거부, 대신 규칙 id 무관, 증권사 공통 분류, 키워드 규칙 덮어쓰기, 국내/해외 분리, 빈 사전 동작
- 검증 7 (3단계 외화 분리): `node tests/run-foreign.js` — 키움해외 CSV 통화별 분리·환전 양쪽 반영·숨은 비용·하루 단위 검증·통화별 허용오차·잔액 보정·통화별 사전 반영·다시 판별
- 검증 8 (4단계 순서 복원): `node tests/run-order.js` — 샘플 순서를 메모리에서 섞어(뒤 절반 역순·가운데 1/3 역순·같은 날 안 뒤집기) 잔액 연결로 복원되는지, 원본 순서 파일은 손대지 않는지, 거래 누락 파일은 거짓 완료가 안 되는지
- 검증 9 (5단계 추가 적요): `node tests/run-memo.js` — 증권 추가 적요 1·2·열별 값 빼기·사전 반영(형식 memo)·같은 형식 자동 적용, 은행 열별 값 빼기
- 검증 10 (기관 선택·형식 이름·웹페이지 표): `node tests/run-picker.js` — 웹페이지 제목 표 합치기, 고른 기관으로 다시 판별, 은행 이름 목록, 형식 이름
- 사전 검사: `npm run check-dictionary [-- 파일]` — 화면에서 저장한 institutions.js 를 올리기 전에 검사(오류면 종료 코드 1). 구조 변환 저장: `node tools/check-dictionary.js --migrate`(먼저 백업)
- 검증 4 (사전): `node tests/run-dictionary.js` — IM 열 지정·KB 수수료 포함·KB국민은행 적요 열을 사전에 반영 → 재판별 결과, 다른 계좌 불변, 저장 파일에 고객 정보 없음, 실제 사전 파일 불변
- 전체: `npm run test:all` / 공개 전 검사: `npm run check-publish`
- 기존 도구 원본이 바뀌었을 때 증권 엔진 재생성: reference/ 갱신 → `node tools/extract-securities-engine.js` → 검증 2·3
- 사전 재생성(주의, 덮어씀): `node tools/extract-presets.js --force` — 먼저 version/ 백업.
- Bash에서 node가 안 잡히면: `export PATH="/c/Program Files/nodejs:$PATH"`

## 위치
- 데이터(샘플): 루트의 엑셀 파일 23개 (이동하지 않음)
- 참고용 원본 복사본: `reference/`
- 배포본: 루트의 `index.html` + `lib/` + `dictionary/` + `engine/`(engine.js, securities-legacy.js, convert.js, dictionary.js) — 함께 옮겨야 동작. GitHub Pages 는 저장소 루트 배포.
- 사용설명서: 루트 `README.md` (팀원용 사용법 + 사전 공유 절차 + GitHub Pages 배포. GitHub 첫 화면에 표시되므로 사용설명서 역할을 겸함)
- 다운로드 결과: 브라우저 다운로드 폴더

## 예외 계좌 분석 절차 (규칙 17)
1. 파일 위치 확인(보통 프로젝트 루트). 고객 데이터이므로 이동·수정하지 않는다(*.xlsx/*.xls/*.csv 는 .gitignore).
2. 진단: `node tools/analyze-account.js <파일>` — 구조, 파일명 약칭, 국내·해외 사전 형식별 열 일치율(빠진 열), 보수적 판정, 통화 구성, 통화별 변환·잔액 검증, 불일치 행, 환전 행.
   보고서를 남길 때는 `cases/YYYYMMDD_기관_요약.txt` 에 저장한다(cases/ 는 .gitignore, 로컬 전용).
3. 원인 분류: ① 새 형식(열 이름) ② 거래종류 분류 ③ 비용·정산 처리 ④ 외화·환전 ⑤ 잔액·순서 문제(데이터 누락 포함) ⑥ 코드 기능 부족.
4. 사전 기록: 그 증권사(국내/해외 구분)의 처리 방식으로 기록한다. 기록 위치(구조 2): 열 이름·형식 지문 → profiles[], 거래종류 분류 → 형식 분류 + 증권사 공통 분류(shared), 행 제외·금액 대체 같은 처리 → 증권사 specialRules(모든 형식) 또는 형식 specialRules, 키워드 → 증권사 keywordRules(공통 securities.keywordRules 를 덮어씀). 원화·외화가 한 파일에 섞인 해외 계좌 → 그 형식의 currencySplit(통화 열, 원화 쪽 금액·잔액 열과 분류, 외화 정산금액 열). 사례는 cases/ 에 남긴다(첫 사례: cases/20261110_키움해외_요약.txt).
   열 구성만 다른 새 양식(특수 처리 없음)은 그 증권사의 새 형식(profiles, 형식 이름 label·형식 지문 signature 포함)으로 기록한다(사례: cases/20261111_미래에셋_776_요약.txt).
   사례 1건에서 나온 특수 처리 규칙(통화 분리·잔액 기준 분류 등)은 일반 규칙(형식 profiles)이 아니라 증권사의 예외 규칙(exceptions[])으로 기록한다(사용자 지시 20261110-017): signature = 사례 파일의 전체 열 이름(열 이름이 모두 같을 때만 적용), cases = 사례 수, source = 사례 기록 파일, classificationBasis: "balance" = 입금/출금을 사전 목록이 아니라 잔액으로 분류(같은 거래종류도 입금/출금이 될 수 있음). 같은 형식의 사례가 더 쌓이면 일반 규칙으로 올릴지 사용자와 정한다. 다른 증권사 형식에 억지로 맞추지 않는다(규칙 16). 계좌번호·고객명·금액은 기록하지 않는다.
   ⑥ 코드 기능이 필요한 경우는 사용자에게 알리고 별도 작업으로 진행한다.
5. 검증: 그 파일이 자동 완료(또는 남는 차이의 원인 설명)되는지, `npm run test:all` 로 기존 샘플 결과가 바뀌지 않는지 확인한다. 새 파일을 기준값(tests/expected*.json, 로컬 전용)에 추가한다.
6. 기록: 사전 changelog(고객 정보 없이 "기관: 변경 내용"), HISTORY.md(작업ID·원인·변경·검증), `npm run check-dictionary`·`npm run check-publish` 통과 후 사용자에게 "dictionary/institutions.js 를 올리면 팀 전체 적용"이라고 보고.
