# LS Datacenter Campus 3D v4 — 20개 독립 작업 및 검증표

- 작성: 2026-10-15 (KST)
- 대상 저장소/브랜치: `kimheeseo/LSCNS` / `main`
- 대상 3D 파일: `DCI/DataCenter/LS_Datacenter_3D.html`
- 원본 3D 파일 감사 기준 blob SHA: `2beb285d8512e1e062032846c5e3d1d801d8f959`
- 금지: `DCI/DataCenter/LS_Datacenter_Campus.html`, `DCI/DataCenter/index.html` 및 기존 BOM 계산·제품 데이터 소스 수정/삭제 금지
- 기준 시점: 분석 당시 3D 페이지는 v4.7.0. 코드를 보지 않고 새 기능을 구현 완료로 판정하지 않는다.
- 운영 규칙: **항목 하나 완료 → 실제 구현 및 검증 결과 보고 → 사용자 확인 후 다음 번호 진행**. 완료 기준 미충족 시 `진행 중/미완료`로 보고한다.

## 현재 완료 상태 (2026-10-15 KST)
- **STEP 01 / 20: 완료** — v4.7.0 정적 감사 및 원본 2D/BOM 보호 SHA 확정.
- **STEP 02 / 20: 완료** — Three.js 전환 아키텍처·API 계약·단일 RAF·분리 캔버스·복귀 플래그·acceptance matrix 작성. 상세: [STEP02_THREEJS_MIGRATION_ADR.md](./architecture/STEP02_THREEJS_MIGRATION_ADR.md).
- **STEP 02 검증:** GitHub 정적 검사 D01~D09 PASS; 브라우저 동작·Three 구현 검증 D10 NOT RUN(차후 STEP 03/19). 계획서/설계 문서만 변경.
- **STEP 03 / 20: 엔진 구현 완료 + 사용자 수동 브라우저 핵심 6/6 PASS / P0 회귀 및 자동화 미검증** — HTML + Three.js Scene Adapter + Bootstrap + Playwright 스크립트 구축. 사용자 6개 항목 통과 보고, 독립 Chromium은 환경 접근 차단. **완전 완료 판정은 잔여 검증 후**. [개발 보고서](./architecture/STEP03_THREEJS_ENGINE_REPORT.md) · [사용자 검증 및 잔여 과제](./architecture/STEP03_BROWSER_VERIFICATION_20261015.md).
- 구현된 버전 고정: Three.js **0.186.1**, WebGLRenderer. `?engine=legacy`(기존)·`?engine=three`(신규) 모드 선택은 **코드에 구현**, 사용자 기본 실행 확인 6/6 PASS. 같은 버전 addons 경로는 import map에 정의되어 있으며 모든 addon 기능 구현을 뜻하지 않는다.
- **STEP 04 / 20: 3D 건축 구현 완료·정적 및 통합 Mock 12/12 PASS / 실브라우저 검증 대기** — [STEP04 건축 모델링 결과 보고서](./architecture/STEP04_BUILDING_ARCHITECTURE_REPORT.md), [Three.js 아키텍처 모듈](./ls3d-three/architecture.mjs), [STEP04 Playwright 테스트](./ls3d-three/step04.playwright.spec.mjs). 외벽·유리·실제 문 pivot·raised floor·ceiling·suspended cable tray·지붕/단면 연동. 신규 GPU 실측은 아직 수행하지 않음.
- **STEP 05 / 20: 7개 구역 설비 외형 고도화 + THREE.LOD + InstancedMesh 구현 / Mock 13/13 PASS / 실제 브라우저 및 FPS 검증 대기** — 비랙 설비 20여 가지 분류, 변압기/UPS/냉각탑/Chiller/ODF/CRAH 등 독립 Mesh, 3단계 LOD, 반복 부품 인스턴싱, 랙 원거리 단순화, renderer 리사이즈 최적화. [STEP05 개발/검증 보고서](./architecture/STEP05_EQUIPMENT_LOD_REPORT.md) · [신규 설비 Factory](./ls3d-three/equipment-factory.mjs) · [Playwright 회귀 시험](./ls3d-three/step05.playwright.spec.mjs).
- **STEP 06 / 20: 엔지니어·기술자·운영자·경비원 3D 관절 캐릭터 구현 완료 / Mock 17/17 PASS / 실브라우저 검증 대기** — 작업복·안전모·경비모·반사조끼·사원증·무전기·팔/다리 관절 및 보행/대기/점검 포즈를 구현. 기존 NPC 위치·이동 경로·작업자 카메라 보존. [STEP06 보고서](./architecture/STEP06_NPC_CHARACTERS_REPORT.md) · [캐릭터 모듈](./ls3d-three/npc-characters.mjs) · [Playwright 시험](./ls3d-three/step06.playwright.spec.mjs). 실제 WebGL2/모바일 FPS는 미측정.
- **STEP 07 / 20: A* 경로 탐색·설비/벽 회피·보행자 근거리 회피 및 Hall 출입문 통과 구현 / 실제 배치 9/9 경로 검사·통합 Mock 16/16 PASS / 실브라우저 QA 대기** — [STEP07 개발 보고서](./architecture/STEP07_NPC_NAVIGATION_REPORT.md), [A* Navigator](./ls3d-three/npc-navigation.mjs), [Playwright 테스트](./ls3d-three/step07.playwright.spec.mjs). 기존 시설·랙·경로 데이터와 6번 캐릭터 및 작업자 추적 카메라 보존. 화재 대피 시나리오 안전성은 미검증.
- **STEP 08 / 20: 점검 FSM 코드 구현·Mock 통합 검증 / 실브라우저 대기.** 이동→정렬→점검→태블릿→이탈, 장애별 우선 설비 선택. [보고서](./architecture/STEP08_INSPECTION_FSM_REPORT.md) · [FSM 코드](./ls3d-three/inspection-fsm.mjs) · [브라우저 시험](./ls3d-three/step08.playwright.spec.mjs).
- **STEP 09 / 20: 방문객 출입 보안 8단계 시각 시퀀스 코드 구현·Mock 검증 / 실브라우저 대기.** 접수/X-ray 수하물/금속 탐지/카드리더/게이트와 경비원 연출. [보고서](./architecture/STEP09_VISITOR_SECURITY_REPORT.md) · [코드](./ls3d-three/visitor-security.mjs) · [브라우저 시험](./ls3d-three/step09.playwright.spec.mjs).
- **STEP 10 / 20: 42/48U 랙과 서버·ToR·ODF·PDU 등 독립 객체 및 선택 HUD 코드 구현·Mock 검증 / 실브라우저 대기.** [보고서](./architecture/STEP10_RACK_PARTS_REPORT.md) · [코드](./ls3d-three/rack-components.mjs) · [브라우저 시험](./ls3d-three/step10.playwright.spec.mjs).
- **STEP 11~20: 미착수.** 사용자 별도 승인 전 진행하지 않는다.

## STEP 01 현황 분석 및 원본 보호 — 완료 기준
- 3D HTML과 2D/BOM 엔트리 파일 확인, 버전과 SHA 기록
- 보존해야 하는 주요 객체·기능 인터페이스와 렌더링 방식 기록
- 20단계 순서/완료 조건 문서화
- 이 단계에서 기존 웹앱 코드를 변경하지 않음

### 현재 3D 코드 분석
- 자체 WebGL 1 컨텍스트 `canvas.getContext('webgl')`, 직접 작성된 vertex/fragment shader, world-space 삼각형/선 데이터 및 GPU 버퍼를 사용한다. Three.js 기반이 아니다.
- 정적 지오메트리 캐시(`LS3D_PHASE4_CACHE`), frustum culling, 랙 상세 단계별 LOD 및 프레임 메트릭 기능이 일부 존재한다. 엔진 교체 시 성능 회귀 우려가 있다.
- 7개 구역 데이터 배열 `Z`: campus, utility, cooling, gray, hall, network, ops.
- 설비 데이터 배열 및 맵: `assets`, `assetMap`; 설비별 유형, 구역, 위치, 크기, 교육용 전력 수치 등.
- 상태: `facilityMode`(aidc/server-room/onprem), `currentZone`, `selected`, `scenario`, `showRoof`, `showSection`, `flow`, `simTime`.
- 랙 기계 동작: `rackMotion`, `setMechanical`, `stepMechanical` (door, slide, cassette, rear). 기존에도 논리 상태/기본 동작이 있으므로 새 3D 객체 계층으로 이식하며 회귀 검증한다.
- NPC: `npcs`, `npcRoutes`, `routeSample`, `npcPose`, `drawPersonModel`. 고정된 경유점·일시정지 기반 동작이 주이며 장애물 우회 및 실제 공간 네비게이션은 추가 검증/개선 대상.
- 경비·X-ray/금속탐지기: `drawGuardAsset`, `drawXray`, `drawWalkDetector`; 현행 연출과 실제 상호작용 시퀀스는 구분한다.
- 렌더/테스트 인터페이스: `drawScene`, `render3D`, `frame`, `window.__LS3D_TEST__`; 기존 방문객 투어 및 장애/대피 시뮬레이션 코드와 호환 필요.
- CSS 의존: `LS_Datacenter_3D_Extensions.css`; BOM 엔트리는 별도 `index.html`이다.
- 제한: 본 감사는 정적 소스 검사다. **실제 브라우저에서의 동작 정상 판정은 19단계 이후**. HTML에 있던 자산형태는 실제 제조사 CAD/BIM과 동일하다고 보장할 수 없다.

## 20개 단위 작업
| 번호 | 원래 단계 | 작업 | 검증 및 완료 기준 |
|---|---|---|---|
| 01 | 1단계 렌더링 | 현황 감사·기능 보존 기준 | 3D/2D/BOM 경계와 기존 기능, baseline blob SHA, 20개 작업표 저장 |
| 02 | 1단계 렌더링 | Three.js 전환 설계 및 호환 인터페이스 | 자체 WebGL vs Three.js 기술 결정, 모듈·이행·롤백 계획 및 acceptance test |
| 03 | 1단계 렌더링 | 3D 엔진·카메라·PBR 조명·기본 씬 구현 | 실제 Three.js mesh, GPU 정상 렌더, 7개 구역 데이터 연계, 유실 기능 없음 |
| 04 | 1단계 렌더링 | 캠퍼스·건물·출입문·내부 공간 실체화 | 벽/문/유리/바닥/천장/케이블 트레이를 깊이 있는 3D 객체로 모델링 |
| 05 | 1단계 렌더링 | 7개 구역 설비 외형 고급화·기본 최적화 | 변압기/UPS/냉각 설비 등 고유 형상, 조명 개선, LOD/인스턴싱/프러스텀 처리 |
| 06 | 2단계 NPC/보안 | 엔지니어·경비원 3D 캐릭터 | 관절 모델, 작업복/안전모/ID, 보행·대기 애니메이션 정상 |
| 07 | 2단계 NPC/보안 | NPC 경로 탐색·충돌 회피 | A*/NavMesh, 벽/랙/사람 충돌 방지, 연속 이동 및 문 통과 |
| 08 | 2단계 NPC/보안 | 설비 점검 행동 상태머신 | 이동→정렬→점검→태블릿→이동, 직무별 시설 및 장애 응답 |
| 09 | 2단계 NPC/보안 | 입구 방문 보안 절차 | 데스크·보안요원 2명·컨베이어/X-ray·금속탐지기·카드리더·문 연속 연출 |
| 10 | 3단계 랙 | 42/48U 랙·부품 독립 객체 | 프레임, 랙별 서버·ToR·ODF·PDU·패치패널의 식별·선택 |
| 11 | 3단계 랙 | 전면·후면 랙 도어/카메라 | 힌지 회전, 닫기, 후면 독립 시점, 구조/클리핑 정확 |
| 12 | 3단계 랙 | 서버 레일 인출 및 NIC/포트 | 선택 서버만 실제 3D 이동·복귀, 포트 식별·설명 |
| 13 | 3단계 랙 | 광 카세트 정밀 조작 | 선택 카세트 인출, 어댑터/커넥터/섬유 경로 3D 표시 |
| 14 | 3단계 랙 | ToR·후면 A/B PDU·케이블 관리 | ToR 최상단 배치, NIC 링크, 전후 배선 및 포트 연결 강조 |
| 15 | 4단계 연결 | MMR 독립 공간 | Carrier/ODF/DWDM/Cross-connect/경계 및 외부-내부 링크 시각화 |
| 16 | 4단계 연결 | 고객 Co-location 공간 | A/B/C Mesh Fence·출입문·고객별 랙/전력/광 분리 |
| 17 | 4단계 연결 | 광·전력·냉각 경로 3D 시각화 | 물리 커브/트레이, 경로별 선택/강조·범례·정확한 예시 데이터 구분 |
| 18 | 5단계 통합 | 카메라·UI·모바일/성능 | 전체 화면·1/3인칭·지붕/단면·설비 찾기·터치, FPS/메모리 점검 |
| 19 | 5단계 통합 | Playwright 회귀·기능 테스트 | 7구역, NPC, 보안, 랙, 흐름, PC/모바일 스크린샷, JS/WebGL 콘솔 오류 확인 |
| 20 | 5단계 통합 | 최종 배포·산출물 보고 | GitHub main 커밋, Pages URL 확인, 변경 파일/모델/애니메이션/테스트/제약 보고 |

## 공통 품질 게이트
1. 작업 단위 커밋·변경 파일·테스트 기록을 남긴다. 완료 여부는 코드 및 검증 증거와 일치해야 한다.
2. 원본 2D Gold Pixel Tour, 기존 BOM, 기존 계산 자료/제품 DB는 변경하지 않는다.
3. 모델 동작은 3D 객체의 위치/회전으로 구현하고 정적인 2D 도식 교체만으로 완료 처리하지 않는다.
4. 라이브러리·GLB 파일은 명시적 버전/상대경로로 GitHub Pages에서도 작동하도록 한다.
5. 실제 CAD/제조사 검증이 없는 모델과 용량·길이·장애 수치는 개념/예시로 표시한다.
6. 각 번호 종료 시 'N번 업무 완성했습니다'는 해당 번호 완료 조건을 실제 확인한 경우에만 사용한다.

### Step 01 정적 검사 증거
- 3D: https://github.com/kimheeseo/LSCNS/blob/main/DCI/DataCenter/LS_Datacenter_3D.html (`2beb285d8512e1e062032846c5e3d1d801d8f959`)
- 2D: https://github.com/kimheeseo/LSCNS/blob/main/DCI/DataCenter/LS_Datacenter_Campus.html (`9c84f24f2101043d7c113ce2d0da50e7245bc318`)
- BOM: https://github.com/kimheeseo/LSCNS/blob/main/DCI/DataCenter/index.html (`692767f0c234df3b8310642e3acd032137c68baa`)
- 본 문서 생성 시 3D/2D/BOM 자체 소스 미변경. 정적 검사와 브라우저 동작 검증을 혼동하지 않는다.
