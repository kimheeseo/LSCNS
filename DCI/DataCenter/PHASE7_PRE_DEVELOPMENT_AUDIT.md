# Phase 7 사전 코드 분석

**기준:** `main` 브랜치에서 읽은 `DCI/DataCenter/` 파일, 2026-10-11 확인

## 1. 현재 3D 데이터 생성 구조

- `LS_Datacenter_3D.html`이 7개 구역(`Z`)과 설비 객체(`assets`)를 만들고 `addAsset()`에서 구역 원점에 상대 좌표를 더해 월드 좌표를 구성합니다.
- `facilityProfiles['server-room'].ids`는 회사 서버실 전용 랙·스위치·ODF·UPS·PDU·CRAC 객체 8개만 표시합니다. `visibleAssets()`가 이를 필터링합니다.
- `LS_Datacenter_3D_Extensions.js`는 별도의 설계 설정 `LS3D_CONFIG`, 네트워크 `nodes`/`links`, 흐름 경로, JSON 스냅샷을 구성합니다. `buildPayload()`가 `T.exportConfig`로 `lsdc-twin-bom/1.0` 패킷을 반환합니다.
- `LS_Datacenter_3D_Phase6.js`는 `T.Z`의 Data Hall 중심 좌표에서 MEP 경로를 계산합니다. 기존 경로는 `x−18`부터 `x+17`까지 약 35 m로, 서버실의 26 m 실내 폭보다 깁니다.

## 2. 기존 BOM 엔진과 입·출력

- 웹 BOM 계산은 `capacity-model.js`, `optical-superpod.js`, `bom-catalog.js`, `product-catalog.js` 및 화면 계층에서 수행합니다.
- Twin→BOM 브리지는 v1.0 패킷의 `facility`, `optical.topology.links`를 입력 화면에 매핑합니다. 현재 GPU 수는 `rackCount × gpuServerRatio × 8`로 가정하며, 해당 값은 패킷에 명시된 NIC 수가 아닙니다.
- 3D export 패킷에는 랙 수·GPU 비율·광 토폴로지 링크·길이가 있지만 GPU/서버/NIC의 명시적 수량 필드가 없습니다. 엔진은 이 값을 임의 확정하지 않고 스키마 입력 또는 명시적 가정으로 구분해야 합니다.
- 현재 Phase 7 `estimate()`는 별도 UI 입력(랙 수, uplink/랙, 속도, 거리)으로 자체 수량을 산출하고 `T.exportConfig()`와는 사후 비교만 합니다. 독립 계산이라 결과가 다를 수 있습니다.

## 3. 3D와 BOM의 잠재 불일치 원인

- Phase 7 UI 수량은 사용자가 별도로 입력하고, Twin packet의 링크 수·거리와는 동기화되지 않습니다.
- 3D 시각 객체 수와 `LS3D_CONFIG.rackCount`가 다를 수 있습니다. `layout()`은 렌더링 최적화를 위해 GPU 랙을 최대 40개만 표시합니다.
- 공유 패킷의 GPU 서버당 8 GPU는 브리지의 고정 가정입니다. NIC/서버당 GPU 및 물리 포트 연결 수는 현재 스키마에 없습니다.
- `nodes`는 논리 토폴로지 노드이며 실제 장비 SKU·포트 cage 수를 확정하는 데이터가 아닙니다.

## 4. 렌더 객체와 설비 물량 구분

- `T.assets`는 장면 자산 객체, `layoutData.assets`는 JSON에 담기는 형상 메타데이터입니다. 둘 다 설비 수량과 동일하다고 볼 수 없습니다.
- `facility.rackCount`는 설계 입력 수량이고 `layout()`은 최대 40개 랙 형상만 표시합니다. 필터링·LOD는 설계 물량을 변경하지 않아야 합니다.
- 링크는 광케이블 어셈블리, 섬유 코어, 양단 광모듈, 논리 포트와 구분해야 합니다. 현재 패킷은 일부 링크에 `type`, `media`, `cores`, `lengthM`를 제공하나 SKU·양단 포트 ID는 없습니다.

## 5. 재사용 가능한 기존 함수/API

- `T.exportConfig()` / `buildPayload()` — v1.0 공유 입력 스냅샷
- `T.visibleAssets()`, `T.facilityMode`, `T.Z`, `T.flowPaths` — 표시 객체·구역 원점·시각 route 데이터
- `LS3D_CONFIG` — 네트워크 설정·중복도·랙 수·냉각·전력 설정
- BOM 계산기 내부의 `capacity-model.js`, `optical-superpod.js`, `bom-catalog.js` — 기존 BOM 페이지 계산/카탈로그 경로. Phase 7은 이 구현을 덮어쓰지 않아야 함.
- Phase 6의 선분–AABB 거리 `clearance()` — 개념 MEP 간격 계산

## 6. 신규 개발 범위

- v1.0 packet을 단일 입력으로 받는 BOM summary 엔진과 표준 `rows` 출력
- 프로토콜·미디어·속도·CPO/Pluggable·거리 정보를 검토하는 경고/차단 validator. SKU가 없으면 공식 호환으로 단정하지 않음.
- Phase 7 UI에서 자체 입력 산식 대신 공통 엔진/validator 결과 렌더링
- 서버실 전용 RoomBounds와 bounds 기반 자산 및 route 유효성 검사
- 서버실 전용 짧은 Busway·냉각 라우트와 Phase 8 점검 항목

## 7. 기존 기능 영향 위험

- 3D HTML은 core inline script 이후 Extensions→Phase2A/2B→Phase45→Phase6→Phase7→Phase8 순으로 외부 모듈을 로드합니다. 신규 공통 엔진과 validator는 Phase7 앞에서 로드해야 합니다.
- 서버실 변형은 같은 7-zone 좌표 공간을 일부 공유하므로 캠퍼스 좌표 기반 route가 남으면 실내 경계 밖으로 렌더링될 수 있습니다.
- shared API는 `layoutData.assets`와 `facility.rackCount`를 혼동하거나 LOD 상태에 의존하지 않도록 해야 합니다.
- 외부 route 차단은 `server-room` 모드에만 적용하고 AIDC/On-Premise 경로·시나리오를 보존해야 합니다.
- 제조사 SKU/reach 자료가 없는 링크는 차단 근거가 없으므로 미확인 경고로 분류해야 합니다.
