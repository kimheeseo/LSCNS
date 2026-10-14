# STEP 04 — 캠퍼스·건물·출입문·내부 3D 건축 구조 개발 보고서

> 작성: 2026-10-15 KST · 저장소: `kimheeseo/LSCNS` · 브랜치: `main`  
> 상태: **실제 Three.js 건축 모델 코드 구현 완료, 정적 검사 및 API Mock 통합 실행 12/12 PASS. 실기기 브라우저 GPU/시각 회귀 검증 미실행.**  
> 범위: 전체 개발 항목 중 **04 / 20** (STEP 05 설비 외형 고급화·LOD·GPU 최적화 미착수).

## 1. 구현 결과

기존 `?engine=three` 모드를 유지하면서 실제 Three.js MeshStandardMaterial 기반 box geometry/Group을 이용해 건물의 **공간·깊이·문 움직임**을 구현했다. SVG 이미지 또는 텍스처만 겹친 방식이 아니다.

| 구성 | 실제 구현 항목 | 주의할 점 |
|---|---|---|
| 캠퍼스 외부 | 다층 기반 지반, 포장 도로, 진입 경로, 차선, 경계 울타리, 출입 게이트 헤더 | 교육용 장면 좌표 및 설계 검토용 형태 |
| Gray Zone | Raised floor, 외벽·기둥·전면 출입 포털, 유리 창호, 강재 보, 부분 천장, 랙 상부 전력/광케이블 트레이 | 전기/기계 지원 공간. 내부 설비 배치는 기존 자산 값을 보존 |
| Data Hall | Raised floor/seam, Hot/Cold Aisle 표시, 벽/유리/강재 구조, 개폐문, 고가 수평 트레이·ceiling 구조 | 42U/48U 정밀 Rack은 STEP 10부터 |
| Network / MMR | 별도 건물 envelope, 전면 유리·출입문, 구조 기둥, 통신 트레이 | MMR 공간·Carrier 경계 정밀화는 STEP 15 |
| NOC / Safety | 독립 건축 shell, 출입문, 유리 패널, 천장/보, 상부 트레이 | 보안 장비 동작 심화는 STEP 09 |
| Utility Yard | 건물처럼 사방 막힌 구조가 아니라 **개방형 fence/gantry**, 장비 이격공간, 상부 케이블 지원 구조 | 설비 상세 형상은 STEP 05 |
| Cooling Plant | **개방형 yard/gantry**, 부분 canopy 지붕, 기둥/서비스 트레이 | 냉각탑·칠러 정밀 형상은 STEP 05 |
| 회사 서버실 | 전용 Hall 단일 구조 `26×24` 개념 크기, 독립 raised floor, 출입문/상부 트레이, campus ground 제외 | 건물 밖 구조 표시 방지, 실제 m 단위와 같다고 단정하지 않음 |

새 3D 모듈:
- [`ls3d-three/architecture.mjs`](../ls3d-three/architecture.mjs): 위 건축 구조 생성 + 자동 경첩 애니메이션.
- [`ls3d-three/scene-adapter.mjs`](../ls3d-three/scene-adapter.mjs): 기존 Zone 데이터 `T.visibleZones()`, `T.currentZone`, `flags.showRoof/showSection`에서 건축 구조를 갱신.
- [`ls3d-three/bootstrap.mjs`](../ls3d-three/bootstrap.mjs): Scene Adapter 모듈 캐시 버전 `20261015-step04`으로 갱신.
- [`LS_Datacenter_3D.html`](../LS_Datacenter_3D.html): Bootstrap 캐시 쿼리 `step04`으로 갱신. 기존 전력·냉각·BOM·NPC 시뮬레이션 코드는 무수정.

## 2. 기술 세부 구현

### 2.1 3D 객체 계층

```text
Three.Scene
└── World
    ├── zoneRoot
    │   ├── arch:campus:ground     [server-room이면 존재하지 않음]
    │   ├── arch:zone:utility      [개방형 야드]
    │   ├── arch:zone:cooling      [개방형 플랜트]
    │   ├── arch:zone:gray         [밀폐 건축물]
    │   ├── arch:zone:hall         [Data Hall / 회사 서버실]
    │   ├── arch:zone:network      [독립 Network/MMR]
    │   └── arch:zone:ops          [NOC/Safety]
    ├── assetRoot                 [STEP 03 장비 및 Rack 상태]
    ├── routeRoot                 [STEP 03 경로]
    └── peopleRoot                [STEP 03 작업자]
```

- 모든 건축물 부품은 실제 `THREE.Mesh` 형태로 존재하며 `mesh.userData.archKind`를 갖는다. 종류: `floor, wall, glass, door, beam, ceiling, roof, tray, frame, path, fence, detail`.
- 기본 Geometry/Material을 공유하여 생성하지만 서로 다른 건축 요소는 독립 메쉬로 확인할 수 있다.
- 케이블 트레이는 바닥/측면 Rail/사다리 횡봉/천장 Hanger가 각각 3D 형태로 생성된다. `power`, `fiber` 등 색을 구분하되 실제 시공 배관·전선수·통신량을 검증한 것은 아니다.
- 유리 객체는 transparent 물성, `depthWrite=false` (기존 재질 캐시 사용) 등으로 설정되어 깊이 있는 파사드를 제공한다. 실제 광학 투과율 수치가 아니다.
- 도어는 왼쪽·오른쪽 두 개의 별도 `THREE.Group` pivot, 열림 정도 `phase`에 따라 `rotation.y`가 변화한다. 캠퍼스 전경에서는 닫힌 상태이며 구역 진입 시 자동 개방한다. 문 클릭/출입권한 제어는 STEP 09 등의 별도 기능이다.
- 기존 Roof 버튼 사용 시 실제 roof slab/상부 패널을 생성한다. Section 버튼 사용 시 구조 보·기둥 및 상부 트레이를 남기고 외벽·지붕을 제거한다.
- `#detailModal` 장비 검사 중에는 `zoneRoot.visible=false`로 설정해 건축 구조가 기존 Rack 상세 검사 시점을 방해하지 않는다. 원래 Rack 상태·표시/선택 API는 유지.
- 동적 Geometry의 장면 생명주기는 기존 `clearGroup()`에서 관리한다. 공유 primitive Geometry는 어댑터 dispose 단계에서 처리.
- 서버실 `server-room` 모드는 기존 `T.visibleZones()` 기준 단일 Hall만 빌드하므로 시각적 외부 지형이 신규 architecture에서 생성되지 않는다.

### 2.2 실측값에 대한 주의

현 데이터의 `[x,z,w,d,h]` 값은 **교육용 도식 스케일**이다. 구조물 크기, 벽 높이, 케이블 간격, 내진, 도어 폭, 내화 등은 공사/BIM 인증 치수가 아니다. 신규 건축 객체는 사용자에게 시설 구역의 구조를 이해시키기 위한 시각 모델이다.

## 3. 실제 수행한 검증

### 3.1 JS 구문·모듈 연결 정적 검사

- `architecture.mjs` 함수 구문 파싱: PASS.
- `scene-adapter.mjs`에 `createArchitectureKit` import, 단일 기존 render loop/자산 선택 연결: PASS.
- `bootstrap.mjs` 및 3D HTML 내 `20261015-step04` cache key 반영: PASS.
- 새 건축 모델에서 사용되는 Zone 명칭이 기존 `T.visibleZones()` 7개 항목과 일치: PASS.

### 3.2 도형 factory Mock 테스트 — 12/12 PASS

Connector에서 GitHub `main`의 **실제 새 모듈 파일을 재조회**, fake `THREE.Group`/Mesh factory를 사용해 장면 빌드 실행.
관찰값은 Mock에서 생성된 객체 수이며 GPU 수치가 아니다.

| 검사 | Mock 실행 결과 |
|---|---|
| 캠퍼스 가시성 | `zones = 7` |
| Campus 실체 형상 | floor/wall/glass 각 1개 이상 |
| 실외 Utility/Cooling 구분 | 두 독립 `arch:zone:*` 확인 |
| 독립 문 모델 | 4개 출입구, 8개 문 leaf 메쉬 |
| 물리 케이블 트레이 | `tray > 0` |
| Raised floor/타일 seam | `detail > 20` |
| 지붕 토글 | Roof 켰을 때 roof > 0 |
| 단면 모드 | 지붕 사라짐, 외벽 개수 감소 |
| Server Room 격리 | 단일 Hall, 캠퍼스 지형 없음 |
| 실내 트레이 | 회사 서버실 트레이 존재 |
| Door hinge | 좌/우 문짝 반대 방향 회전 |
| Render adapter | 빌드·변환·dispose 실행, Mock API 예외 없음 |

기존 STEP 03 포함 통합 Mock 재시험에서도 **12/12 PASS**:
`campus-seven-zones`, `campus-real-walls-glass-floor`, `campus-four-entry-double-doors`, `three-real-cable-trays`, `roof-on`, `section-cut`, `animated-hinges`, `rack-behavior-kept`, `inspector-isolation`, `server-room-one-zone`, `server-room-trays`, `dispose`.

**단, Mock 실행은 실제 WebGL/GPU 화면, 모바일 FPS, 텍스처·재질 광학 거동을 검증하지 않는다.**

### 3.3 실브라우저 테스트 — 아직 미실행

- [STEP04 Playwright 시각/로직 회귀 코드](../ls3d-three/step04.playwright.spec.mjs) 저장 및 구문검사 통과.
- 실제 Chromium/WebGL2에서 실행한 테스트 로그/스크린샷은 **없다**. GitHub Pages 배포/ESM/모바일 성능 PASS를 주장하지 않는다.
- 기존 STEP 03 사용자 수동 6/6 PASS는 이 아키텍처 추가 **이전 버전**에 대한 결과이므로 STEP 04 기능을 검증했다고 볼 수 없다.
- 원격 raw GitHub 파일을 로컬 테스트 컨테이너에서 내려받으려는 시도는 DNS 네트워크 접근 제한으로 실패했다(사이트의 동작 결함을 의미하지 않음).

### 3.4 공개 GitHub Pages 접근 시도

- 2026-10-15: 공개 `.../LS_Datacenter_3D.html?engine=three` 및 `.../ls3d-three/architecture.mjs?v=20261015-step04` 주소를 외부 문서 접근 도구로 조회 시도.
- 관찰: 해당 검사 도구에서 **URL not accessible** 반환. **HTTP 상태 코드/실제 브라우저 동작을 얻지 못함**.
- 판정: **Pages 배포 반영 확인 불가** (404라고 단정할 수 없음). GitHub `main` 파일 생성/수정은 연결된 GitHub API에서 재조회하여 검증했다.
- 사용자 Chrome에서 새 모듈을 강력 새로고침 후 `window.LS3D_THREE.adapter.architecture`를 확인해야 실제 배포가 반영된 것으로 판단할 수 있다.

## 4. 다음 브라우저 수동 점검 절차

1. Chrome/Edge에서 https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=three 를 열고 `Three.js r186 · PBR` 표시를 확인한다.
2. Campus 전경에서 구조가 **박스형 설비만이 아닌 외벽/유리/출입 게이트/실제 지붕 위쪽 구조**로 구별되는지 살핀다.
3. Data Hall, Gray Zone, Network/MMR, NOC/Safety 각각 진입하고 **벽/바닥/천장/상부 케이블 트레이**를 확인한다.
4. Data Hall에서 `지붕 켜기`를 누르면 지붕이 생기고, `단면 보기`를 누르면 외벽과 지붕이 사라지는지 확인한다.
5. Data Hall 진입 시 두 출입문이 반대 방향으로 자동 개방하는지 확인한다.
6. Utility Yard/Cooling Plant가 밀폐 건물이 아닌 개방형 야드/캐노피인지 확인한다.
7. 운영 형태를 `회사 서버실`로 변경하여 외부 캠퍼스 지형이 보이지 않으며 랙/문/트레이가 보이는지 확인한다.
8. 기존 랙 내부 보기, 설비 선택, 7구역 버튼, 기본 `?engine=legacy`가 정상인지 확인한다.
9. F12 → Console에서 오류, `window.LS3D_THREE.adapter.architecture`의 `counts` 및 `doors`를 확인한다.

### Playwright (사용자 PC 또는 연결된 CI 환경)

```bash
npm install --no-save @playwright/test
npx playwright install chromium
npx playwright test DCI/DataCenter/ls3d-three/step04.playwright.spec.mjs
```

기존 STEP 03 테스트도 함께 실행하고 실패한 개별 항목/네트워크 오류를 기록한다.

## 5. 알려진 한계 및 보완 제안

- 신규 건축 객체는 현 단계에서 개별 메시로 생성된다. draw calls·모바일 GPU FPS·메모리는 실제로 계측하지 않았으며, 반복 구조 instancing/LOD/머티리얼 병합 등 고급 최적화는 **STEP 05** 범위다.
- 실제 문 개폐 간섭/문틀 충돌, 공조 유체·열해석 및 시공 가능 이격거리는 본 모델에서 검증하지 않았다.
- 모든 설비의 현실 형태·부품, A/B PDU/ODF/NIC, 고객 cages 및 실제 케이블 경로는 후속 단계다.
- 원래 `server-room` 외부 설비 차단과 기존 운영 모드/비상 시나리오/카메라와의 브라우저 회귀 검증은 추가 확인이 필요하다.
- **기존 2D Tour/BOM 파일, 계산 수식, 제품 DB는 변경하지 않는다.**

## 6. 업무 경계

- STEP 04: **기능 구현 및 정적·Mock 검증 통과. 실제 브라우저 검증 대기**.
- STEP 05: **사용자 요청 전까지 미착수**.
