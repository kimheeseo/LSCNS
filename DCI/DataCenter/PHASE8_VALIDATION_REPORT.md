# Phase 8 검증 보고서

**작업일:** 2026-10-11  
**대상:** LS Datacenter 3D v4.7.0 · Phase 7-A/B 및 Server Room 보완

## 보완 내용

- Phase 6의 랙 안내선은 현재 표시 자산만 순회하며 서버실 추가 line/box/cylinder를 실내 경계로 다시 제한합니다.
- 서버실 구조 기둥의 Y 좌표를 base 기준으로 수정해 기둥이 천장 위로 연장되는 오류를 제거했습니다.
- Phase 7은 열린 패널에서 공유 패킷 변경을 약 0.4초 간격으로 감지해 자동 갱신합니다. 서버당 GPU/NIC 입력 근거와 미확정 수량·길이를 올바르게 표시합니다.
- 호환성은 PASS/UNKNOWN/FAIL을 항상 반환합니다. compatible은 PASS만 true이고 nonBlocking은 차단 없음입니다. SKU/데이터시트 URL/거리/endpoint 프로토콜·포트/connector/FEC 근거 누락은 UNKNOWN, 상충·초과는 FAIL입니다.
- Phase 8은 모드 전환 없이 서버실 모델을 검사합니다. 광 단선과 사용자 지정 랙 수를 포함한 기존 운영 상태를 보존합니다.
- 화면·README·현재 가이드·Phase 8 테스트는 v4.7.0 표기로 통일했습니다. 이전 Word 가이드북은 파일명의 v4.5.3을 유지한 역사적 문서이며 현재 동작은 Phase8_Guide.md를 기준으로 합니다.
- CI 변경 감지에 공통 엔진·validator·room geometry·단위 테스트를 포함했습니다.

## 실행 결과

| 검증 항목 | 결과 | 근거 |
|---|---|---|
| 변경 JavaScript 문법 및 diff 검사 | PASS | node --check, git diff --check |
| 공유 BOM 엔진 테스트 | PASS | 기본 계산, 사용자 지정 GPU/NIC, 누락·빈값·불완전 길이 |
| 호환성 테스트 | PASS | PASS/UNKNOWN/FAIL, 빈 제품 목록, 포트 초과, FEC/connector 상충, 1.6T 및 IB alias |
| 서버실 geometry 및 실제 구조 렌더러 테스트 | PASS | 내부/외부 자산·경로, 실제 column의 천장 경계 |
| Chromium desktop / mobile 회귀 | PASS | 82개 통과, 실패 0, 페이지 오류 0, 콘솔 오류 0 |
| 서버실 실제 추가 draw 콜백 검사 | PASS | 출력 line/box/cylinder primitive 외부 0건 |
| UI 입력 자동 BOM 반영 | PASS | 랙 20→30, 서버 16→24, GPU 128→192, NIC 16→24 |
| 종합 점검 상태 보존 | PASS | 광 단선·사용자 지정 랙 수·설정·선택·카메라·로그·흐름 유지 |

실행 환경: Linux Playwright 1.56.1 / Chromium 141 / SwiftShader WebGL 및 합성 모바일 터치. 물리적 Windows GPU/Android/iPhone 측정이 아닙니다. 테스트는 저장소 소스를 로컬 HTTP 서버에서 실행하며 GitHub Pages 배포는 별도 확인합니다.

재현 명령:

```sh
node DCI/DataCenter/tests/phase7_engine.test.cjs
node DCI/DataCenter/tests/phase7_compatibility.test.cjs
node DCI/DataCenter/tests/server_room_geometry.test.cjs
npm install --no-save --prefix /tmp/lsdc-phase1 playwright@1.56.1
/tmp/lsdc-phase1/node_modules/.bin/playwright install --with-deps chromium
NODE_PATH=/tmp/lsdc-phase1/node_modules node DCI/DataCenter/tests/phase8_browser_regression.cjs
```

회귀 결과 JSON·보고서·스크린샷은 tests/test-results/phase8/에 생성되며 GitHub Actions artifact로 보관합니다. 상태 보존 비교는 점검 함수를 동기 실행하는 구간의 운영 데이터를 대상으로 하며 이후 정상 시뮬레이션 시간 진행을 멈추지 않습니다.

## 과거 실행과 한계

최초 개발 환경에서는 Chromium 실행 파일이 없어 브라우저 검사가 차단됐습니다. 이후 [GitHub Actions 기준 실행](https://github.com/kimheeseo/LSCNS/actions/runs/38124668013)은 79개를 통과했습니다. 위 82개는 그 뒤의 보완 검사를 추가한 결과입니다.

PASS는 입력된 근거에 대한 규칙 검사입니다. 데이터시트 URL 원문, 실제 SKU의 lane/cage 구조·MPO polarity·FEC 설정·벤더 지원 목록의 독립 검증이나 구매 승인까지 완료했다는 뜻이 아닙니다. 랙 수가 증가해도 토폴로지가 자동 증설되는 기능은 이 요약 엔진에 포함되지 않습니다. MEP는 개념 좌표·AABB 사전검사이며 시공/BIM 승인은 별도입니다.
