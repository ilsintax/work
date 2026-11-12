# HANDOFF — 통합 통장내역 편집기

- 작업 ID: 20261112-028
- 마지막 갱신: 2026-11-12 (KST)
- 상태: 로컬 바로 저장(프로젝트 사전 폴더 연결)·사전 가져오기 도구 추가. GitHub 에는 아직 이전 버전(사용자 지시로 업로드 보류).
- 작업 컴퓨터: 사무실 PC (Windows 11, SynologyDrive 동기화 폴더)
- 마지막 작업본: index.html, engine/(engine.js, convert.js, dictionary.js, securities-legacy.js), dictionary/institutions.js(구조 2, 내용 버전 22) — 다음 작업 시작 시 sha256sum 으로 확인
- 사용자 확인 완료: 키움해외 원화 입금 누락 = 입금·출금 중 하나만 조회한 파일(20261110-016). 맞추려면 "전체"로 다시 받기
- 사용자 확인 완료: (계좌번호) 거래내역 = 미래에셋증권 → 형식2로 사전 기록(20261111-021)
- 다음 단계
  1. GitHub 업로드 — 올릴 파일은 README "GitHub에 올리는 파일" 참고(공개/비공개·저장소 이름 결정 대기), iM증권 형식 배포 사전 반영 여부
- 명령: `npm run test:all`, `npm run check-dictionary [-- 파일]`, `npm run check-publish`, `npm run analyze -- <파일>`, 백업 `bash tools/backup.sh <작업명> <작업ID> <파일...>`
- Node 경로: bash 에서는 `export PATH="/c/Program Files/nodejs:$PATH"` 필요
- 알려진 문제: 통화 분리 설정은 키움 해외 형식1에만 있음. 은행형은 하루 단위 검증·잔액 보정 없음. 잔액 0 + 금액 있는 행은 기존 엔진 규칙상 기준점 아님. 공통 백업 도구 실행 불가(대체 방식). 테스트는 샘플이 있는 PC에서만.
- 동기화: SynologyDrive. 다른 PC는 Node 설치 후 `npm install`.
