# STEP 14 / 20 — ToR · Rear Rack PDU A/B · NIC/Cable 3D 정밀화

- 작업일: 2026-10-15 (KST)
- 대상: `DCI/DataCenter/LS_Datacenter_3D.html?engine=three` / `main`
- 판정: **구현 및 소스 기반 모의 검증 완료 / 독립 Chromium·WebGL 실화면 검증 대기**
- 사용자 승인 범위: STEP 14만 작업. STEP 15 이후의 신규 개발은 이 업무에 포함하지 않음.

## 기존 코드 재점검

이번 시작 시점 GitHub `main`에는 이미 `rack-network-cabling.mjs`(선택 가능한 ToR 8포트, NIC→ToR, 후면 Rack PDU A/B 및 선택 필터), `scene-adapter.mjs`의 컨트롤 UI 연결이 존재하였다. 다만 STEP 14 전용 테스트·결과 보고서가 없었고, 다음 회귀 위험이 있었다.

1. 기존 `rack-components.mjs`에서 광 패치패널이 ToR보다 위에 배치되어 ToR 최상단 배치 조건과 달랐다.
2. 서버만 레일을 따라 이동하고 NIC 배선의 고정된 원래 TubeGeometry 시작점은 따라가지 않았다.
3. 다른 종류의 배선을 필터로 선택할 때 일부 링크를 `visible=false` 처리하여 전체 경로 문맥을 볼 수 없었다.

## 이번 구현 보완

| 기능 | 변경 및 구현 증거 |
|---|---|
| ToR 최상단 | 42U의 U41–42 및 48U의 U47–48을 ToR 2U 예시 모델 영역으로 지정. 패치패널·ODF는 아래 영역으로 배치. 실제 제조사 모델의 U 배치 규격이 아님. |
| 독립 ToR 포트 | 랙마다 8개 selectable port mesh, `partMap` 상 각 `tor-port` 식별자 유지 |
| NIC 연결 | 서버의 실제 Group.position.z 인출량을 읽어 해당 서버의 광 링크 원점을 함께 이동. 이동량 변화가 있을 때 연결 TubeGeometry를 다시 생성하고 종전 geometry를 dispose. |
| PDU A/B | 후면 전원 A와 B를 각각 독립 경로와 색상·선택 대상으로 유지 |
| 링크 필터 | 광, 전원 A, 전원 B 선택 시 매칭 링크 opacity 0.94 / 비매칭 0.11; 다른 링크를 지우지 않고 반투명하게 유지 |
| 기존 인터페이스 | 기존 `rackCabling`, `selectRackPart`, 서버 선택·레일·전후면 카메라·문 조작 API 유지 |

여기서 서버 NIC 연결, 코너 굴곡 및 A/B 전원은 실제 라우팅 데이터가 아닌 **교육용 위치 및 연결 예시**다. 일반적인 양쪽 전원 공급을 시각화한 것이며 실제 장비의 이중 PSU 및 전기 정격을 검증한 것은 아니다.

## 수행한 검증

**실제로 실행한 테스트 (GitHub 소스 로드 → JavaScript 실행 환경에서 THREE 호환 객체 기반 단위·통합 검사):**
- 42U/48U 각각 ToR의 실제 U 위치·패치패널 대비 상단 위치 검사
- 랙별 8개 pickable `tor-port`의 독립 `partMap` 등록 검사
- 각 서버당 NIC→ToR 한 경로, PDU-A 및 PDU-B 각 한 경로 검사
- 모든 링크가 계속 표시되는지 / 선택된 계통만 강조되고 나머지는 투명해지는지 검사
- 선택 서버 이동량 -0.63 시 광케이블 시작점 실제 갱신 / 이전 geometry dispose 검사
- 3종 링크의 기하 재생성, 서버 복귀 후 케이블 원위치 검사
- 모든 링크 강조 모드 복귀 검사

**실행 결과: 18/18 PASS (두 랙 높이 조합), 실제 저장된 `step14.unit.test.mjs` 테스트 3/3 PASS.** 테스트 파일의 소스 구문 검사도 정상. 위 검증은 진짜 GPU 장면 캡처가 아니며 서버의 그래픽·픽셀·프레임 속도 및 DOM 이벤트를 검증하지 않는다.

**작성했지만 실행하지 않은 실제 브라우저 검사:**
- `step14.playwright.spec.mjs`: ToR 클릭 포트 선택, 후면 A/B 필터 opacity, 광 NIC 인출 후 TubeGeometry 시작점 추적, 모바일 390px 컨트롤 스크린샷 및 JS 콘솔 오류 검사.
- 현재 작업 환경에서 외부 실행 페이지에 대한 독립 Chromium/WebGL 검증 및 새 스크린샷 확보는 **미실행**. 위 파일 존재만으로 Playwright PASS를 주장하지 않는다.

### 직접 재검증 명령

```bash
node --test DCI/DataCenter/ls3d-three/step14.unit.test.mjs
npm install --no-save @playwright/test
npx playwright install chromium
LS3D_BASE_URL=https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html npx playwright test DCI/DataCenter/ls3d-three/step14.playwright.spec.mjs
```

기존 2D와 BOM의 blob SHA는 현재 조회에서도 동일함:
- `LS_Datacenter_Campus.html`: `9c84f24f2101043d7c113ce2d0da50e7245bc318`
- `index.html`: `692767f0c234df3b8310642e3acd032137c68baa`

## 수정·생성 파일

- `ls3d-three/rack-components.mjs` — ToR를 가장 높은 U 위치로 조정
- `ls3d-three/rack-network-cabling.mjs` — NIC service-loop 동기화 / 강조 투명도
- `ls3d-three/step14.unit.test.mjs` — 실제 소스 실행 가능한 Node 단위 테스트
- `ls3d-three/step14.playwright.spec.mjs` — 실제 브라우저 회귀 검증 초안
- `architecture/STEP14_RACK_NETWORK_CABLING_REPORT.md` — 본 결과 문서
- `LS3D_PROGRESS_STATUS.md` 및 `LS3D_V4_20_TASK_ROADMAP.md` — 14단계 상태 동기화

변경 범위에 포함하지 않은 것: 구형 2D Tour, DataCenter BOM, 기존 제품 DB, STEP 15–20 기능 구현.

## 잔여 확인 항목

1. GitHub Pages에서 최신 `rack-components.mjs`·`rack-network-cabling.mjs`가 정확하게 로드되는지 확인.
2. 실제 Chrome/Edge WebGL에서 42U/48U 랙 전후면 시점, ToR 포트 클릭, 광·A/B 강조와 케이블 이동·복귀 확인.
3. 모바일 터치 및 케이블/후면 도어 간의 깊이 가림, 장시간 실행 메모리/프레임 점검.
4. 장비 제조사별 포트 수, 케이블 종류, 링크 속도, 전류/용량, 이중화 검증은 추후 별도 데이터 검증이 필요.

실행 페이지: https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=three
