# LS Datacenter Campus 3D — 통합 개발 진행 현황

> 작성일: 2026-10-15 (KST)  
> 저장소: https://github.com/kimheeseo/LSCNS · 브랜치: `main`  
> 작업 원칙: 번호별 구현 → 실제 검증 → 완료 보고 → 사용자 확인 후 다음 번호

## 진행 상태

| 단계 | 업무 | 상태 | 검증 |
|---|---|---|---|
| 01/20 | 기존 3D 코드 분석·보호 범위 | 완료 | 정적 코드 분석 및 기준 SHA 기록 |
| 02/20 | Three.js 전환 구조 설계 | 완료 | 아키텍처·호환 API·롤백·검증 계획 기록 |
| 03/20 | Three.js 엔진·3D Scene·조명·카메라 | **구현 완료·기본 브라우저 수동 6/6 PASS / 최종 완료 판정 보류** | 사용자 M01~M06 PASS, 독립 Chromium 환경 접근 차단, P0 회귀 및 Playwright 미검증 |
| 04/20 | 캠퍼스/건물·출입문·유리·바닥·천장·케이블 트레이 | **3D 구현 완료, 통합 Mock 12/12 PASS / 실브라우저 검증 대기** | 신규 architecture.mjs, Roof/Section 동작/힌지 및 server-room 격리 API 검사, Playwright 테스트 파일 준비 |
| 05/20 | 7개 구역별 설비 3D 구분 및 LOD/인스턴싱 | **코드 구현 완료, API Mock 13/13 PASS / 실브라우저·FPS 검증 대기** | 설비 Factory·구역별 형상, 3단계 LOD, 반복 부품 GPU InstancedMesh, 기존 Rack/2D/BOM 보존 |
| 06/20 | 엔지니어·경비원 3D 모델 및 관절·보행 애니메이션 | **코드 구현 완료, 통합 Mock 17/17 PASS / 실브라우저 GPU 검증 대기** | 역할별 6명, 헬멧·경비모·ID·무전기·관절·워크/점검·정지, 회사 서버실 격리, 기존 NPC/랙 유지 |
| 07/20 | NPC A* 경로 탐색·벽/설비/사람 회피·출입문 통과 | **코드 구현 및 실제 자산 기반 9/9 경로 PASS, 통합 Mock 16/16 PASS / 실제 WebGL2 브라우저 검증 대기** | 1.25 격자 A*, 랙/건물 footprint, Hall 포털, 근거리 보행자 회피, 작업자 시점 동일 좌표 |
| 08/20 | NPC 설비 점검 상태머신 | **구현 코드 및 Three.js 연결 확인 / 브라우저 QA 대기** | `inspection-fsm.mjs`에서 점검 단계 및 전환 구현 확인 |
| 09/20 | 방문객 보안 검색 절차 | **구현 코드 및 Three.js 연결 확인 / 브라우저 QA 대기** | `visitor-security.mjs`에서 8단계 방문 동선 및 보안 장치 시각화 확인 |
| 10/20 | 42/48U 랙 부품 독립 객체 | **구현 코드 및 Three.js 연결 확인 / 브라우저 QA 대기** | `rack-components.mjs` 및 `step10.playwright.spec.mjs` 존재 확인, 실제 Chromium PASS 미확인 |
| 11/20 | 전·후면 랙 도어 및 카메라 | **실제 Three 코드 구현·모의 테스트 통과 / Chromium QA 대기** | 독립 문 힌지, 전면/후면 시점 조작, 이전 레거시 상태 연동 |
| 12/20 | 서버 레일 인출 및 NIC/포트 | **코드 구현·모의 테스트 통과 / Chromium QA 대기** | 한 번에 선택한 서버만 인출, 42/48U 서버 14개/개념 NIC 56포트(샘플) |
| 13/20 | 광 카세트 상세 조작 | **코드 구현·모의 테스트 통과 / Chromium QA 대기** | 카세트 이동/커버 개폐, 랙당 8개 커넥터·3D 광섬유, 포트별 raycast |
| 14/20 | ToR 최상단·후면 PDU A/B·NIC 배선·필터 정밀화 | **구현 완료·실제 모듈 Mock 18/18 PASS, Node 테스트 내용 3/3 PASS / Chromium QA 대기** | U 최상단 ToR, 포트 8개, 인출되는 NIC 연결선, 비선택 링크 반투명화·기존 조작 API 보존 |
| 15/20 | MMR 별도 4구획·Carrier/ODF/DWDM/Core 모델·4개 출입문·광회선 선택 | **실제 소스 구현 및 JS Mock 16/16 + Scene Bridge 8/8 + Node 호환 테스트 4/4 PASS / Chromium QA 대기** | 구획 출입문 실제 `rotation.y`, 선택 Fiber `TubeGeometry`, 상세 메타정보·선택 경로 강조·구역 전환 상태 보존 |
| 16/20 | Co-location | **기존 구현 존재 / 이번 차수 변경·재검증 없음** | 사용자 다음 승인에 따라 별도 진행 |
| 17~20/20 | 전역 배선·카메라/성능·종합 테스트·최종 배포 | **이번 차수 미진행** | 별도 요청 이후 진행 |

**완전 검증 완료로 기록된 단계는 STEP 01~02. STEP 14~15는 소스 기반 모의 검증을 완료했으나 브라우저 검증은 대기 중이다. STEP 03 사용자 기본 수동 브라우저 6/6 PASS. STEP 04~07 Mock 검사 PASS. STEP 08~10 코드·Three API Mock 21/21 PASS(이전 결과). STEP 11~13 부품 Mock 24/24 및 Scene Adapter 통합 Mock 17/17 PASS. STEP 04~13 신규 WebGL2/모바일/Playwright/FPS 최종 검증은 대기.**

### 2026-10-15 수동 검증 결과
- 사용자 확인 6종 PASS: 3D 건물·설비 표시, Three.js r186·PBR 표시, 7개 구역 전환, 랙 선택·상세정보, PC/터치 카메라 회전·확대, 도어 및 서버 인출.
- 본 결과는 사용자 수동 확인이며 Playwright GPU 실행·콘솔 로그·FPS·모든 설비 데이터 일치 검증은 아니다.
- 독립 Chromium 접속 시도는 외부 URL 접근 정책으로 `net::ERR_BLOCKED_BY_ADMINISTRATOR`; **시험 환경 차단(사이트 장애 증거 아님)**.
- [추가 회귀 검증 우선순위·수정 후보·콘솔 진단](./architecture/STEP03_BROWSER_VERIFICATION_20261015.md).


## 개발 자료와 소스

- [20개 업무 로드맵](./LS3D_V4_20_TASK_ROADMAP.md)
- [02 전환 구조 설계](./architecture/STEP02_THREEJS_MIGRATION_ADR.md)
- [03 엔진 개발 및 검증 보고서](./architecture/STEP03_THREEJS_ENGINE_REPORT.md)
- [03 사용자 브라우저 6/6 결과 및 잔여 점검](./architecture/STEP03_BROWSER_VERIFICATION_20261015.md)
- [04 건축 3D 모델링 및 검증 보고서](./architecture/STEP04_BUILDING_ARCHITECTURE_REPORT.md)
- [04 신규 건축 모듈](./ls3d-three/architecture.mjs)
- [04 Playwright 브라우저 회귀 테스트 코드](./ls3d-three/step04.playwright.spec.mjs)
- [05 설비 모델링·LOD·인스턴싱 개발 보고서](./architecture/STEP05_EQUIPMENT_LOD_REPORT.md)
- [05 신규 장비 형상 Factory](./ls3d-three/equipment-factory.mjs)
- [05 Playwright 브라우저 회귀 테스트](./ls3d-three/step05.playwright.spec.mjs)
- [06 캐릭터 모델/보행 애니메이션 개발 보고서](./architecture/STEP06_NPC_CHARACTERS_REPORT.md)
- [06 역할별 3D 캐릭터 모듈](./ls3d-three/npc-characters.mjs)
- [06 Playwright 브라우저 회귀 테스트](./ls3d-three/step06.playwright.spec.mjs)
- [07 A* 탐색·충돌 회피 결과 보고서](./architecture/STEP07_NPC_NAVIGATION_REPORT.md)
- [07 NPC A* Navigator](./ls3d-three/npc-navigation.mjs)
- [07 Playwright 브라우저 회귀 테스트](./ls3d-three/step07.playwright.spec.mjs)
- [08 점검 상태머신 보고서](./architecture/STEP08_INSPECTION_FSM_REPORT.md), [FSM 코드](./ls3d-three/inspection-fsm.mjs), [Playwright](./ls3d-three/step08.playwright.spec.mjs)
- [09 방문객 출입 보안 보고서](./architecture/STEP09_VISITOR_SECURITY_REPORT.md), [모듈](./ls3d-three/visitor-security.mjs), [Playwright](./ls3d-three/step09.playwright.spec.mjs)
- [10 랙 42/48U/부품 선택 보고서](./architecture/STEP10_RACK_PARTS_REPORT.md), [모듈](./ls3d-three/rack-components.mjs), [Playwright](./ls3d-three/step10.playwright.spec.mjs)
- [11 랙 전·후면 문 및 시점 개발 보고서](./architecture/STEP11_RACK_DOORS_CAMERA_REPORT.md)
- [12 개별 서버 인출 및 NIC 개발 보고서](./architecture/STEP12_SERVER_RAIL_NIC_REPORT.md)
- [13 광 카세트·커넥터·광섬유 보고서](./architecture/STEP13_FIBER_CASSETTE_REPORT.md)
- [11 독립 문 모션](./ls3d-three/rack-door-controls.mjs)
- [12 서버 레일·NIC 모듈](./ls3d-three/rack-server-rails.mjs)
- [13 광 카세트 모듈](./ls3d-three/rack-optical-cassette.mjs)
- [11~13 브라우저 회귀 스크립트](./ls3d-three/step11to13.playwright.spec.mjs)
- [14 랙 배선 상세 보고서](./architecture/STEP14_RACK_NETWORK_CABLING_REPORT.md)
- [15 MMR 독립 공간 결과 보고서](./architecture/STEP15_MMR_REPORT.md)
- [15 MMR 3D 모델 모듈](./ls3d-three/mmr-room.mjs)
- [15 Node 호환 회귀 테스트](./ls3d-three/step15.unit.test.mjs) · [Playwright 브라우저 시험 초안](./ls3d-three/step15.playwright.spec.mjs)
- [3D 웹페이지 HTML](./LS_Datacenter_3D.html)
- [Three.js Scene Adapter](./ls3d-three/scene-adapter.mjs)
- [Three.js Bootstrap/Legacy Fallback](./ls3d-three/bootstrap.mjs)
- [Playwright 브라우저 테스트 코드](./ls3d-three/step03.playwright.spec.mjs)

## 사용자 실행 링크

- 기본 WebGL: https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html
- 새 Three.js 시험 모드: https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=three
- 구형 WebGL 강제: https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=legacy

참고: 사용자 기본 화면/조작 검증 6/6 PASS는 **STEP 03 버전**에서의 수동 확인이다. STEP 04에서 새로 추가한 외벽/유리/도어/트레이의 **실제 브라우저 동작은 아직 검증하지 않았다**. 독립 자동 검증 및 배포 서버 로그는 확보하지 않았다.

## STEP 04 추가 구현 사항 (2026-10-15)
- 밀폐된 Gray/Data Hall/Network/MMR/NOC 건축 shell + Utility/Cooling 개방형 설비 야드를 실질 3D Mesh로 구분.
- 건물 입구 2짝 유리문(THREE.Group hinge), 레이어 바닥·raised-floor 줄눈·오버헤드 cable trays·부분 ceiling·roof/section 옵션을 연결.
- 실제 GitHub Scene Adapter + Architecture 모듈로 수행한 API mock 12/12 PASS. **실 GPU 렌더 FPS 및 Playwright는 미실행**.
- 상세: [STEP04_BUILDING_ARCHITECTURE_REPORT.md](./architecture/STEP04_BUILDING_ARCHITECTURE_REPORT.md).

## STEP 05 추가 구현 사항 (2026-10-15)
- 7개 구역 장비의 고유 특징(변압기/UPS/발전기/냉각탑/칠러/CDU/ODF/CRAH 등)을 개별 3D Mesh로 구현.
- Three.LOD 3단계 및 GPU InstancedMesh로 포트·환기 그릴·라디에이터 핀·루버 반복 구성. 기존 Rack door/slide/cassette 등 상태 유지.
- 실제 GitHub source를 재조회하여 실행한 JavaScript Mock 통합 13/13 PASS: 샘플 25개 설비 중 24개 시각 모델 분류, 반복 부품 18개 인스턴싱 배치/226개 부품 처리.
- 위 수치는 Mock 형상 생성 통계이며 GPU draw calls / FPS 개선 수치가 아니다. 실제 PC/모바일 브라우저 회귀 검증 필요.
- 상세 [STEP05_EQUIPMENT_LOD_REPORT.md](./architecture/STEP05_EQUIPMENT_LOD_REPORT.md).

## STEP 06 추가 구현 사항 (2026-10-15)
- 기존 블록 NPC를 직원·보안요원 역할별 Three.js 관절 모형으로 교체. 안전모·경비모·사원증·무전기·작업복·안전화·반사조끼·팔다리 관절을 구현.
- T.npcPose() 좌표/방향/상태를 기존대로 유지하여 walk/pause/inspect/idle 애니메이션을 연결. STEP 07 경로 탐색·충돌 회피는 아직 없음.
- 연결 GitHub 실제 코드 기반 Three API Mock 통합 17/17 PASS. AI Campus 6명·Hall 1명·회사 서버실 경비원 제외·태블릿 자세·재생성/해제 확인.
- 사용자 GPU 실화면/모바일/Playwright 자동 실행과 FPS 검증은 미실행.
- 상세 [STEP06_NPC_CHARACTERS_REPORT.md](./architecture/STEP06_NPC_CHARACTERS_REPORT.md).

## STEP 07 추가 구현 사항 (2026-10-15)
- 1.25 개념 좌표 격자, 보행자 반경 0.39, 8방향 최소 힙 A*·충돌 안전 경로 단축·근거리 동료 회피 구현.
- AI/On-Prem Data Hall 출입문을 이용한 경유점; 회사 서버실 전용 26×24 공간 경계, 건물·설비 footprint 충돌 확인.
- HTML에 있는 운영 자산 74개 배치를 직접 파싱/실행한 9개 모드·작업자 경로 검증 9/9 PASS. 2명 맞교행 회피 모의 검사 7/7 PASS. STEP 03~07 API Mock 통합 16/16 PASS.
- 기존 직원 HUD/추적 카메라와 Three 캐릭터가 동일 navigator pose를 사용. 기존 화재 대피 재정의 스크립트는 우선 유지(화재 안전성은 별도 미검증).
- 실제 PC/모바일 WebGL2 브라우저·Playwright FPS/메모리 검증 대기. [STEP07 보고서](./architecture/STEP07_NPC_NAVIGATION_REPORT.md).

## STEP 08~10 구현 및 검증 기록 (2026-10-15)
- **STEP 08:** 이동/정렬/점검/태블릿/이탈 개념 점검 상태머신, 역할·시나리오별 대상 선택. 기존 A* 좌표 보존.
- **STEP 09:** 방문객 접수·X-ray 수하물 컨베이어·금속 탐지·카드리더·게이트 힌지 8단계. 경비원 2명 기존 NPC 연동, 회사 서버실에서 방문객 숨김.
- **STEP 10:** 개념 42/48U 랙, 서버/ToR/ODF/광카세트/PDU A·B/패치 패널의 독립 Group, click-pick 메타데이터 및 HUD. 기존 도어·인출 인터페이스 유지.
- 실자산 74개를 사용한 통합 Three API Mock **21/21 PASS** 이후 상태머신 재진입·보안 동선/출입문 연출을 보완하고 개별 시험 수행. **실 Chromium/WebGL2·모바일·GPU FPS/메모리·Playwright 실행 로그는 확보하지 않음**.
- [STEP08 상세](./architecture/STEP08_INSPECTION_FSM_REPORT.md) · [STEP09 상세](./architecture/STEP09_VISITOR_SECURITY_REPORT.md) · [STEP10 상세](./architecture/STEP10_RACK_PARTS_REPORT.md).

## STEP 11~13 추가 구현·검증 (2026-10-15)

- STEP 11: 전면 문은 원래 `T.mechanical(a).door`를 보존, 독립 후면 문 hinge motion 도입. 선택 랙의 앞/뒤 시점은 `T.cameraDesired`로 직접 설정. 구형 `m.rear` fallback은 유지.
- STEP 12: 서버 개별 group의 position.z만 이동하도록 STEP 10 공통 server-tray를 개선. 서버별 가이드 레일·NIC·4개 개념 NIC 포트 추가, 하단 조작 UI로 서버 선택/인출.
- STEP 13: 인출 카세트에 8개 어댑터·페룰·커넥터 및 실제 THREE.TubeGeometry 3D 광섬유 경로와 열리는 커버 구현. Raycast 시 정확한 `fiber-port` ID를 상위 카세트보다 우선 처리.
- **모의 테스트:** 42U/48U 부품 동작 24/24, 전체 Scene Adapter 결합 17/17 PASS. 실제 페이지 WebGL2·모바일 FPS·카메라 클리핑·터치 및 Playwright는 **미실행**.
- 11~13번 보고서 및 `step11to13.playwright.spec.mjs`를 참조. 기존 2D Tour와 BOM 계산 파일은 변경 금지/보존.

## 브라우저 확인 방법

### 직접 확인
1. Chrome 또는 Edge에서 `?engine=three` 링크를 연다.
2. 화면 상단 엔진 표시가 `Three.js r186 · PBR`로 나타나는지 확인한다. `기존 WebGL 복귀`이면 실패 원인 조사 필요.
3. 7개 구역을 하나씩 선택하고 장비 메쉬가 표시되는지 확인한다.
4. 설비 클릭·우측 패널·확대·이동·지붕/단면·rack door/slide/cassette 버튼을 검사한다.
5. 개발자 도구(F12) 콘솔에서 `window.LS3D_THREE.status`가 `ready`, `window.LS3D_THREE.renders`가 증가하고 `window.LS3D_THREE.diagnostics`에서 `assets>0`인지 확인한다.
6. Console의 오류 및 Network의 JS 404/CORS 오류를 기록한다.
7. `?engine=legacy`에서도 기존 기능이 작동하는지 비교한다.

### Playwright 스크립트 (Chromium 설치된 개발 환경)
```bash
npm install --no-save @playwright/test
npx playwright install chromium
npx playwright test DCI/DataCenter/ls3d-three/step03.playwright.spec.mjs
```

스크립트는 기본/Three 모드, 7개 구역, server-room 범위, 모바일 viewport 및 캡처를 검사하도록 작성되어 있다. **스크립트가 저장된 것과 실행·통과한 것은 별개**다.

## 보호 대상 확인

- 2D Gold Pixel Tour: `DCI/DataCenter/LS_Datacenter_Campus.html` blob `9c84f24f2101043d7c113ce2d0da50e7245bc318`
- DataCenter BOM: `DCI/DataCenter/index.html` blob `692767f0c234df3b8310642e3acd032137c68baa`
- STEP 03 당시 3D `LS_Datacenter_3D.html` blob `25174e7ae90ab9d300f3f1e1fa66b65615f3a8c3`

## 다음 작업 조건

STEP 14 랙 배선 정밀화에 이어 STEP 15 MMR의 Carrier/ODF/DWDM/Core 전면 장비·개별 힌지 출입문·클릭 가능한 광 링크 및 연결 메타데이터를 main에 반영했다. STEP 15 실제 코드 Mock **16/16 PASS**, Scene Bridge 연동 **8/8 PASS**, 커밋된 Node 호환 시험 소스의 JS 실행 **4/4 PASS**. **독립 Chromium/WebGL2·모바일·스크린샷·성능 실검증은 미실행**이므로 완전 검증 판정은 보류한다. [STEP15 보고서](./architecture/STEP15_MMR_REPORT.md). STEP 16은 사용자 별도 요청 시 진행한다.
