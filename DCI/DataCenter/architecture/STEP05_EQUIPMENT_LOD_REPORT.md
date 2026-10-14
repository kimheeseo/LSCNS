# STEP 05 / 20 — 7개 구역 3D 설비 외형 고도화·LOD·인스턴싱 검증 보고서

- 작성일: 2026-10-15 (KST)
- 저장소: `kimheeseo/LSCNS` · 브랜치 `main`
- 대상: [Three.js 3D Digital Twin](../LS_Datacenter_3D.html?engine=three)
- 개발 경계: **STEP 05만 수행**; STEP 06 NPC 관절·애니메이션 작업은 착수하지 않는다.
- 판정: **실제 3D 구현·구문 검사·모의 통합 실행 13/13 PASS. 실브라우저 WebGL2/성능/FPS는 미검증.**

## 1. 개발 목적

STEP 03의 범용 직육면체 중심 설비가 서로 비슷해 보이는 한계를 줄이고, 부품의 대표 특징을 실제 3D Mesh로 나타낸다. 반복 요소는 `THREE.InstancedMesh`로 GPU 드로우 호출을 묶고, 떨어진 설비는 `THREE.LOD`로 단순화한다. 기존 랙 문·서버·카세트 조작 기능과 BOM 수치·설비 id는 변경하지 않는다.

## 2. 실제 개발 범위

| 구역 | 장비 타입 | 3D에서 구분되는 형상 |
|---|---|---|
| 01 Campus | X-ray 검사, 금속 탐지기, 보안 게이트, 접수 데스크 | 컨베이어·검사 터널/차폐막, 양측 탐지 기둥, 카드 리더, 안내/접수대 |
| 02 Utility Yard | 변압기, 수배전반, 발전기, ATS, 연료 탱크 | 변압기 방열판/3상 부싱, Switchgear 다중 베이·차단기창·상태등, 발전기 엔진/배기/라디에이터, ATS 전면 화면, 원통 탱크 |
| 03 Cooling Plant | 냉각탑, Chiller, 냉수 탱크, Pump/HX, CDU | 냉각탑 루버·팬·하부 수조, 칠러 압축기·코일·배관, 원통 저장 탱크, 판형 열교환기·펌프, 냉매 매니폴드 |
| 04 Gray Zone | UPS, 배터리 Bank, PDU, CDU, Power Busway | UPS 제어 패널·환기구, Battery drawer·상태등, PDU/냉각 분배장치, 전력 도체·행거 |
| 05 Data Hall | CRAH/CRAC, ODF, 일반/고밀도 Rack, 고객 Cage | 공기 토출/흡입 루버, 광 포트·패치 트레이, 기존 Rack 도어·서버 인출 유지, Cage 경계 프레임 개념 형태 |
| 06 Network/MMR | Carrier, Meet-Me Room, ODF, Core/Spine/Leaf | 다층 통신 섀시·Port LED/광패널·수직 배선 덕트. **기존 mechanically operable Core/Spine/Leaf Rack 분류는 조작 동작 보존을 우선** |
| 07 NOC/Safety | 관제 데스크, 보안, 소화설비, 유지보수 | 별도 모니터/화면, 컨트롤 데스크, 4개 소화용기·헤더 배관, 정비대 |

실제 제조사/자사 제품과 동일한 CAD 형상·측정 치수·내부 전력 손실 모델이라는 의미는 아니다.

## 3. 구현 파일

### 3-A. `ls3d-three/equipment-factory.mjs` (신규)

- `createEquipmentFactory(THREE,{cube,cylinder,material,boxGeometry})` 엔진. 기존 데이터의 `asset.type`, `asset.id` 문자열로 **24개 카테고리**를 분류한다(샘플 데이터 기준; 실전 모든 장비 종류 수와 혼동 금지).
- 변압기 코어, 라디에이터, 3상 부싱 / 발전기 배기관 / 냉각탑 팬·루버 / UPS 전면 UI / CDU 매니폴드 / 통신 포트 / 작업 공간을 실제 `THREE.Mesh` 등으로 별도 생성한다.
- 자세한 세부 구조는 `THREE.LOD`의 3단계(0,65,155 거리 단위)로 생성: 근거리 상세 모델, 중거리 외함·도어, 원거리 실루엣. 단위는 사용자 개념 장면 좌표다.
- 동일 부품 반복(광 포트, LED, 열교환 핀, 그릴, 배터리 표시등 등)은 `THREE.InstancedMesh`에 `setMatrixAt()`으로 위치·크기를 넣어 다중 개별 draw 대신 배치한다.
- 기존 공유 `BoxGeometry`와 재질 캐시를 사용. 인스턴싱/LOD 통계를 `window.LS3D_THREE.diagnostics.equipment`에 노출한다. 일부 통계(`instancedBatches` 등)는 **누적 생성 횟수**다.

### 3-B. `ls3d-three/scene-adapter.mjs` (수정)

- 비랙 장비는 공장 모듈로 생성하고, 기존 `T.isMechanicalRack()` 장비는 랙 조작 `makeRack` 모델을 보존.
- 랙도 근거리 상세 `rack-high`와 원거리 `rack-low`를 준비, 카메라 기준 96 장면 단위 초과 및 미선택이면 단순 형태로 전환. 선택된 랙은 항상 상세 모델 유지.
- 비랙 설비 `THREE.LOD.update(camera)` 적용, 구역/시설 모드가 바뀔 때 기존 id 기반 registry 갱신과 결합.
- 구역 이탈한 설비의 임시 링 `RingGeometry`·`MeshBasicMaterial` 등을 `disposeAsset`에서 해제. 공유 Base Box/Cylinder Geometry/Material은 재사용/어댑터 dispose 시 정리.
- `WebGLRenderer.setPixelRatio` / `setSize`는 화면 크기 또는 DPR 변경 시에만 실행해 매 프레임 canvas 크기를 재설정하는 일을 방지.
- 기존 `T.pickAsset`, 선택/비상상태, GPU 랙 상세/문·인출 상태, NPC 루프, 시설 모드와 BOM에는 변경 없음.

### 3-C. cache-bust 및 검증

- `ls3d-three/bootstrap.mjs`: scene adapter URL `?v=20261015-step05`.
- `LS_Datacenter_3D.html`: bootstrap URL `?v=20261015-step05`. 기존 WebGL 초기화/단일 RAF·2D/BOM은 유지.
- `ls3d-three/step05.playwright.spec.mjs`: 실제 Chromium 회귀 테스트 (파일만 준비, **실행 결과 없음**).

## 4. 검증 및 증거

### 4-A. 원격 소스 실제 실행(Mock) — 13 / 13 PASS

연결 GitHub에서 `equipment-factory.mjs`, `scene-adapter.mjs`, `architecture.mjs` 실제 코드를 재조회한 뒤, JavaScript Test Double(THREE.Group/LOD/InstancedMesh/WebGLRenderer의 핵심 API)을 사용해 mock 장면을 실행했다.

| 검사 | 판정 |
|---|---|
| 7구역 캠퍼스 장면 유지 | PASS |
| 대표 가상 설비 25개 등록/표시 | PASS |
| 20종 이상 서로 다른 장비 분류 | PASS |
| 실제 인스턴싱 코드 경로 실행 | PASS |
| 모든 대표 설비/랙의 LOD 후보 존재 | PASS |
| 씬 내 LOD 노드 20개 이상 | PASS |
| 씬 내 InstancedMesh 존재 | PASS |
| Hall 구역으로 변경 시 설비 필터링 | PASS |
| 기존 랙 도어 모델 객체 유지 | PASS |
| Company server-room 1개 구역 격리 | PASS |
| server-room 상세 검사 선택 유지 | PASS |
| 7구역 복귀 시 설비 재생성 | PASS |
| adapter.dispose 자원 정리 호출 | PASS |

샘플 입력 25개에서 관찰한 결과:
- `activeKinds`: 24개 타입 (랙 포함, 개념 샘플)
- `instancedBatches`: 18개의 반복 부품 배치 (모의 생성 결과)
- `instancedParts`: 배치 처리 대상 226개 반복 부품 (모의 생성 결과)
- **주의**: 226개의 GPU draw calls가 18개로 줄었다는 실측 비교가 아니다. 벤치마크/FPS 미실행이며 개별 재질, LOD, 광원, 건축 구조 등 다른 draw call도 존재한다.

### 4-B. 정적/보존 검증

- Equipment 모듈과 Scene Adapter 구문 검사 및 버전된 ESM Import 연계: 통과.
- 원본 2D Tour, 기존 BOM HTML, 외부 BOM 계산·제품 데이터 소스 변경 없음 (GitHub compare 및 blob SHA로 추가 점검).
- 기존 레거시 엔진은 기본값이며 `?engine=three`만 변경.

### 4-C. 추가 브라우저 검증 (NOT RUN)

실제 WebGL2 장치의 `fps`, `p95 프레임 시간`, draw calls 수, GPU/메모리 피크, 모바일 렌더 성능·발열 등은 **측정하지 않았다**. GitHub Pages의 최신 버전이 방문자 CDN/캐시에 반영된 상태도 이 보고서 작성 시점에 독립적으로 확인하지 못했다.

브라우저 수동 검사 권장 사항:

1. [Three.js 테스트 페이지](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=three) 강력 새로고침.
2. Utility에서 변압기의 3상 부싱/냉각핀, 수배전반 패널, 발전기 배기부가 서로 다른지 확인.
3. Cooling에서 두 냉각탑 팬/루버, Chiller 압축기, CDU의 매니폴드가 식별되는지 확인.
4. Gray에서 UPS/배터리/PDU가 서로 다르게 생겼는지 확인.
5. Data Hall에서 CRAH 루버·ODF 광포트를 확인하고, GPU Rack door/slide/cassette를 기존대로 열고 닫을 수 있는지 확인.
6. 7개 구역 모두 전환, 서버실만 표시, 레거시 기본/강제 URL 비교.
7. F12 Console에서 예외 및 `window.LS3D_THREE.diagnostics`, `window.LS3D_THREE.adapter.equipment` 값이 보이는지 확인.

Chromium + Playwright 지원 환경에서는:

```bash
npm install --no-save @playwright/test
npx playwright install chromium
npx playwright test DCI/DataCenter/ls3d-three/step03.playwright.spec.mjs DCI/DataCenter/ls3d-three/step04.playwright.spec.mjs DCI/DataCenter/ls3d-three/step05.playwright.spec.mjs
```

## 5. 개발 한계, 리스크 및 추후 단계

- 일부 네트워크 Core/Spine/Leaf는 기존 인터랙티브 Rack 기능을 우선하는 관계로 STEP 05 세부 네트워크 모듈 형상이 나타나지 않을 수 있다. STEP 10~14에서 실제 랙 내부/ToR/광패널 구조를 보완한다.
- 건축, 설비, 선형 트레이를 합친 전체 scene의 GPU 렌더링 호출·모바일 FPS는 추후 실제 측정·최적화가 필요하다.
- 여전히 교육용 개념 치수이며 검증된 제조사 제품 치수/전기·냉각 물성값이 아니다.
- STEP 06 NPC 캐릭터의 관절·작업복·안전모/ID 강화는 요청 전 **미착수**.

**STEP 05 판정: 코드 구현 완료, 정적 및 API Mock 13/13 PASS, 실제 브라우저 회귀·성능 검증 대기.**
