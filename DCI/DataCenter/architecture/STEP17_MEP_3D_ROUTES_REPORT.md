# STEP 17 / 20 — 전력·냉각·광 경로를 실제 3D 배관·케이블 트레이로 시각화

- 작업일: 2026-10-15 KST · `kimheeseo/LSCNS` `main`
- 결과: **코드 구현 + 소스 기반 모의 검증 완료 / 실제 브라우저 WebGL QA 미완료**
- 대상: `DCI/DataCenter/LS_Datacenter_3D.html?engine=three`
- 후속 18~20 단계는 이번 작업에서 진행하지 않음.

## 1. 변경 전 구조와 문제

`scene-adapter.mjs`의 `routeScene(flags)`는 기존 `T.flowPaths` 2D 배열을 `THREE.Line(new THREE.BufferGeometry(),new THREE.LineBasicMaterial())` 방식의 가는 직선으로 표현했다. 계통별 배선이 현장 경로/설비별로 구분되지 않고 물리적 트레이·배관 형상이나 클릭 가능한 링크 단위 메타데이터가 없었다.

## 2. 구현 내용

신규 `ls3d-three/campus-mep-routes.mjs`를 단일 Three.js scene render 루프에 연결하였다.

| 계통 | 실제 데이터 소스 ID 기반 대표 연결 (개념) | 3D 객체 |
|---|---|---|
| **전력** | Transformer(`xfmr-a`) → Switchgear → UPS → PDU → ToR/서버랙. PDU에서 고객 A/B/C Cage 랙별 분기 | 색상 구분 `THREE.TubeGeometry`, 수평 채널용 3D ladder tray(실제 box base / side / rung), 계통 말단 표시 |
| **냉각** | Cooling Tower → Chiller → Pump → CDU Plant → Gray CDU → GPU Rack, Pump → CRAH A 보조 분기 | 반경이 더 큰 3D Tube 배관 + 독립 Pipe Coupler Mesh |
| **광통신** | Carrier → MMR → ODF → Core → Spine A → Leaf A → ToR, ODF → Data Hall MMR/ODF, Hall MMR → 고객 A/B/C 랙 | 3D 광 경로 Tube·Overhead tray·터미널, raycast 가능한 회선 선택 |

- **설비 ID 정합:** 기존 HTML의 `assetMap` 및 `visibleAssets()`와 일치하는 설비가 존재할 때만 경로를 생성. 원본 설비와 BOM을 증감하거나 좌표를 변환·저장하지 않는다.
- **3D 경로:** 경로마다 시작 설비 몸체에서 공용 설비 고도까지 올라간 뒤 직교형 수평 waypoint를 거쳐 수신 장비로 하강. 경로의 Tube 표면 및 트레이 측벽/격자에 실제 3D 깊이가 존재. CAD 경로를 수치 재현한 것은 아님.
- **구역 필터:** Campus는 전체 표시, 나머지 구역은 시작 또는 끝 설비가 그 구역인 링크만 노출. `server-room`과 Rack inspect 시 외부 MEP 경로 숨김.
- **계통 필터:** 기존 `flags.flow.power/cooling/fiber` 사용 설정을 그대로 적용. 새로운 별도 계통 강조 및 개별 구간 선택 가능.
- **클릭·상세:** `mep:tube:<id>` 및 단말 마커에 `mepRouteId` 저장. 3D 객체 선택 또는 오른쪽 상단 MEP 패널에서 회선 클릭 시 발신 설비/수신 설비/계통 표시.
- **가독성:** 선택된 경로의 불투명도 0.94, 나머지 0.09. 비선택 회선을 숨기지 않고 지오메트리 보존.
- **성능:** `T.currentZone`·시설 모드·설비 좌표·필터 상태로 signature 생성. 변함없으면 매 프레임 곡선을 재생성하지 않음. 변경 시 이전 Curve의 Geometry/Material dispose; 기존 RAF 외 추가 동작 루프는 없음.
- **불확실성:** `verified:false`, `lengthMeters:null`, `ratedSpeed:null`, `connector:미지정·미검증` 데이터 계약으로 실제 길이나 커넥터/정격 성능을 꾸며내지 않음. 경로 구성·고도 역시 교육용 예시.

**주의:** Utility Yard → Gray Zone → Hall / Cooling Plant → Hall / Network → Hall 구간은 위상 연결을 설명하기 위한 개념적 배관/케이블이다. 외벽 관통부, 방화 구획, 실제 케이블 하중, 전력 차단/보호, 냉각 수리해석, 통신 링크 감쇠/처리량 및 건물 간 정식 CAD/BIM routing validation은 수행하지 않았다.

## 3. 실제 수행한 검사 결과

GitHub에서 읽은 실제 STEP 17 모듈과 원본 `LS_Datacenter_3D.html`을 JavaScript 실행 환경에서 실행:

- **실제 자산 60개를 직접 로드한 모의 검사 22/22 PASS.**
- 3개 계통에 실제 존재하는 설비 ID만을 사용하여 시각 경로 생성.
- UPS→PDU→ToR, Cooling Tower→CDU→GPU, Carrier→MMR→ODF→Core→Spine→Leaf→ToR 및 MMR→고객 A/B/C 경로 검사.
- 모든 링크에 `THREE.TubeGeometry` 형태의 모의 인스턴스, 출발/도착 drop, 케이블 트레이/배관 coupler 존재 검증.
- 단일 광 링크 선택 시 해당 Tube의 opacity/metadata가 변경되고 다른 링크는 반투명으로 남는지 확인.
- 구역 필터·전력 필터 On/Off·정적 frame에서 geometry 재생성 방지·zone 변경 시 dispose·server-room/inspect 격리·최종 dispose 검증.
- **`step17.unit.test.mjs`에 저장한 별도 회귀 검사 5/5 PASS:** 실제 테스트 파일의 JavaScript 테스트 본문을 불러 모듈과 함께 Mock Three 호환 객체에서 실행하였다. `node --test` CLI의 별도 로그는 없음.

위 검사들은 **실제 GPU 렌더링이나 사용자가 화면을 클릭한 브라우저 테스트를 대체하지 않는다.**

### 브라우저 검증 시도 및 한계

Chromium이 있는 환경에서 GitHub Pages에 접속을 시도했으나 `Page.goto: net::ERR_BLOCKED_BY_ADMINISTRATOR`로 브라우저 네트워크 접근이 차단됨. 원격 사이트 장애의 증거가 아니다. 실제 WebGL 실화면·새 데스크톱/모바일 스크린샷·WebGL 콘솔·FPS 테스트 결과는 **미확보**.

작성만 한 `step17.playwright.spec.mjs` 4종 수용성 테스트:
1. 3D Tube/물리 트레이/배관 렌더와 3계통 선택.
2. MMR → Co-location A 광경로 선택, 발신/수신 정보 및 강조.
3. Hall 구역 선택, 냉각 계통 필터, server-room 격리.
4. 390px 모바일에서 MEP 패널/광경로/스크린샷 확인.

### 재실행

```bash
node --test DCI/DataCenter/ls3d-three/step17.unit.test.mjs
npm install --no-save @playwright/test
npx playwright install chromium
LS3D_BASE_URL=https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html npx playwright test DCI/DataCenter/ls3d-three/step17.playwright.spec.mjs
```

## 4. 변경/생성 파일

- `ls3d-three/campus-mep-routes.mjs` — 신규 MEP 3D Tube / tray / pipe / 물리 연결 선택 엔진.
- `ls3d-three/scene-adapter.mjs` — 기존 2D Line 라우트 함수 대체, raycast/diagnostics/cleanup 통합.
- `ls3d-three/bootstrap.mjs`, `LS_Datacenter_3D.html` — STEP 17 ES 모듈 캐시 경로 업데이트.
- `ls3d-three/step17.unit.test.mjs` — 소스 단위 테스트 5개.
- `ls3d-three/step17.playwright.spec.mjs` — 브라우저 테스트 코드 4개 (미실행).
- `architecture/STEP17_MEP_3D_ROUTES_REPORT.md` — 이 개발 보고서.
- `LS3D_V4_20_TASK_ROADMAP.md`, `LS3D_PROGRESS_STATUS.md` — 상태 반영.

기존 2D Tour(`LS_Datacenter_Campus.html`) 및 BOM(`index.html`) 파일은 이번 단계에서 변경하지 않는다. 기존 `T.flowPaths`와 `?engine=legacy` 방식은 그대로 보존한다.

## 5. 다음 단계로 넘길 항목

18단계: 카메라/모바일 및 전체화면/클리핑·패널 중첩·FPS/메모리 최적화. 19단계: 전체 Playwright 수용성 회귀와 스크린샷. 20단계: 최종 배포 및 문서화.
