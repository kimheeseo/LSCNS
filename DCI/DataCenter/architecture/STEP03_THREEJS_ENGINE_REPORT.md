# STEP 03 — Three.js 렌더링 엔진 개발 결과 및 검증 기록

- 기준일: 2026-10-15 KST
- 관련 개발 계획: [20단계 개발표](../LS3D_V4_20_TASK_ROADMAP.md)
- 기능 목표: Three.js 기본 Scene·WebGLRenderer·장비 메시·카메라·기존 데이터 연동
- 작업 범위: **STEP 03만**, STEP 04 건물 정밀 GLB 모델링 미착수
- 판정: **코드 구현 및 정적/모의 구동 시험 통과 / 실제 GPU 브라우저 회귀 테스트 미수행**
- 유의: 모의 Three.js API 검증과 브라우저 WebGL2 실측은 다름. 브라우저 검증 전에는 작업 전체를 '완전 검증 완료'로 선언하지 않는다.

## 1. 구현 내역

| 구성 | 파일/기능 | 구현 내용 |
|---|---|---|
| 3D Scene | `ls3d-three/scene-adapter.mjs` | 실제 `new THREE.Scene()`, `new THREE.WebGLRenderer()`, `new THREE.Mesh()`, `renderer.render(scene,camera)` 사용 |
| 조명·재질 | 같은 파일 | HemisphereLight, DirectionalLight, PBR MeshStandardMaterial, ACES 톤 매핑, sRGB 출력, 그림자, Fog |
| 7개 구역 | 같은 파일 | 기존 `T.visibleZones()`, `T.currentZone`에 연동한 실제 floor/wall/roof/ceiling support 3D 객체 생성 |
| 설비 식별 | 같은 파일 | 기존 `T.visibleAssets()`/assetId를 Three.Group·Mesh로 생성; 랙, UPS, 변압기, 냉각탑, Chiller, X-ray 기초 형상 분리 |
| 랙 기초 동작 | 같은 파일 | 기존 `T.mechanical()`에 따라 Door `rotation.y`, tray/cassette `position.z` 실시간 반영 (STEP 10~14 정밀 조작은 미구현) |
| NPC 기초 모델 | 같은 파일 | 머리/몸체/팔/다리 독립 3D 메시, `T.npcPose()` 기반 이동·간단한 보행 (동작 상태머신과 충돌회피는 06~08) |
| 광/전력/냉각 | 같은 파일 | 기존 flowPaths에 기초한 3D Line 지오메트리 표시; `server-room`은 외부 흐름 경로 표시 억제 |
| 선택 처리 | 같은 파일 | `THREE.Raycaster` → assetId → 기존 `T.pickAsset`; 설비 우측 정보창 원본 로직 유지 |
| 웹 엔진 전환 | `ls3d-three/bootstrap.mjs` | `?engine=three` opt-in, WebGL2 검사, 별도 `#sceneThree` canvas, 포인터/휠/핀치 컨트롤 |
| 복구 | 같은 파일 | import 실패·WebGL2 미지원·런타임 오류 발생 시 `#scene` 레거시 canvas로 복귀 |
| 단일 RAF | `LS_Datacenter_3D.html` | 기존 `frame(now)` / `render3D(dt)`만 주도; Three.js adapter가 매 프레임 호출됨. 추가 RAF 없음 |
| Playwright 브라우저 회귀 | `ls3d-three/step03.playwright.spec.mjs` | 레거시 모드, 7개 구역, three 렌더 프레임, server-room 격리, 모바일 캡처의 실행용 스크립트 |

## 2. 배포 및 사용 안내

- 기존 기본 페이지(안정성 우선): https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html
- **Three.js 신규 엔진 지정 주소:** https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=three
- **레거시 강제 주소:** https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=legacy
- `Three.js r186 · PBR` 레이블이 나타나면 초기화에 성공한 것이며, `기존 WebGL 복귀` 표시가 나타나면 오류·미지원 환경에서 복구되었다는 뜻이다.
- 브라우저 개발자 콘솔에서 `window.LS3D_THREE.status`, `window.LS3D_THREE.diagnostics` 확인 가능.
- **GitHub Pages 실제 배포 동기화 및 브라우저 화면 작동은 이 보고 작성 시점에 실측 검증하지 않았음.** 링크는 구현된 페이지의 예상 주소이며 정상 실행을 확인했다는 뜻이 아니다.
- Three.js `0.186.1` (npm CDN) 정확히 고정, 동일 버전 모듈 import map. 모듈 호스팅과 CDN이 접근 가능해야 한다. 최신 WebGLRenderer는 WebGL2 필요; WebGL1에서는 레거시 경로 유지.

## 3. 구현 상태와 명확한 한계

이번 STEP 03의 장비는 실제 GL Geometry로 구성했으나 **건물 및 모든 장비의 높은 외형 품질과 제조사 CAD/BIM 등가성은 구현 대상이 아니다**.
- STEP 04: 건물/출입문/케이블 트레이의 정밀 지오메트리
- STEP 05: 7구역 설비 고유 형상 강화, 인스턴싱·LOD·성능 튜닝
- STEP 06~09: 관절 모델·A* 또는 NavMesh·NPC 점검·보안 절차
- STEP 10~14: 정밀 랙·ToR·서버 인출·광 카세트·A/B PDU
- STEP 15~17: MMR/고객 Cage 및 구체적인 물리 케이블·배관 경로
- STEP 18~20: 정밀 UI/성능/회귀/최종 배포 QA

추가 제한:
1. `server-room` 모드에서 기존 외부 MEP 경로 노출을 방지하기 위해 신규 엔진의 flow path 표시를 일단 억제한다. 적법한 실내 링크 표시 복구는 STEP 17.
2. 엔진이 자체 WebGL1과 별도 WebGL2 canvas를 사용하므로 모바일 장치의 컨텍스트/메모리 성능을 실제로 측정해야 한다.
3. NPC 충돌 회피·고급 조작·실규격 랙·콘텐츠 폴리곤 예산은 후속 단계이며 아직 개발 완료로 보고하지 않는다.
4. 별도 브라우저 자동화 스크립트는 저장했으나 실행 환경의 GitHub Pages/CDN 접근 제한으로 **실제 Playwright 성공 결과나 스크린샷은 아직 없다**.

## 4. 이번 단계의 검증 결과

### 4-A. PASS — 원격 GitHub 정적 검사

- 3D HTML 내 inline JavaScript 4개와 import map JSON 파싱: **PASS**.
- Three Scene Adapter `.mjs` 구문 파싱: **PASS**.
- bootstrap `.mjs` 구문 파싱: **PASS**.
- 실제 Three.js Scene, WebGLRenderer, Mesh, 조명, rack transform, Raycaster, 7구역 API 연결 코드 존재: **PASS (소스 검사)**.
- 기존 RAF 한 계통 유지 및 별도 animation loop 추가 금지: **PASS (소스 검사)**.
- 3D HTML importmap/CDN 버전 일치: **PASS**.
- `?engine=three` 선택 시 로딩과 legacy fallback 처리: **PASS (소스 검사)**.
- 기존 BOM과 2D Gold Pixel Tour 파일을 변경 목록에 포함하지 않음: **PASS (GitHub compare 결과)**.

### 4-B. PASS — THREE API Mock 런타임 테스트 (GPU가 아님)

실제 GitHub source를 불러와 `WebGLRenderer`, `Scene`, `Mesh`, `Geometry`, `DirectionalLight` 등을 모사한 JavaScript Test Double로 `createSceneAdapter`를 구동.

| 케이스 | 관찰 |
|---|---|
| Campus 모드 1 frame | frames 1, assets 2, zones 7 |
| Data Hall 모드 | frames 2, assets 1, zones 2 (캠퍼스 바닥 포함) |
| server-room 모드 | frames 3, assets 1, zones 1 |
| dispose | `disposed === true` |
| 총괄 | Mock 구동 오류 0건 |

이 테스트는 **컴포넌트 연결의 런타임 논리 검증**이지 **GPU WebGLRenderer가 실제로 픽셀을 그렸다는 증거가 아니다**.

### 4-C. NOT RUN — 실기기 브라우저/배포 검증

| 검사 | 상태 |
|---|---|
| GitHub Pages의 `?engine=three` 사용 시 WebGL2 초기화 | 미실행 |
| PC/Android/iPhone 실제 화면 캡처 | 미실행 |
| 브라우저 콘솔/네트워크 오류 0건 | 미확인 |
| desktop/mobile fps 및 메모리 실측 | 미실행 |
| Playwright 7 zones/동작/캡처 실행 | **스크립트 준비, 실행 미완료** |

### Playwright 실행

GitHub Pages 배포가 실제로 적용된 뒤 Playwright가 준비된 Node 환경에서:

```bash
npm install --no-save @playwright/test
npx playwright install chromium
npx playwright test DCI/DataCenter/ls3d-three/step03.playwright.spec.mjs
```

테스트를 CI에 접목하기 전에 사용할 Playwright 버전과 GPU 설정을 고정하고, 대상 HTML 배포 전파 상태를 확인해야 한다.

## 5. 보호 파일 및 GitHub 변경 검증

**변경 허용 범위(이번 STEP 03):**
- `DCI/DataCenter/LS_Datacenter_3D.html` 수정
- `DCI/DataCenter/ls3d-three/scene-adapter.mjs` 생성
- `DCI/DataCenter/ls3d-three/bootstrap.mjs` 생성
- `DCI/DataCenter/ls3d-three/step03.playwright.spec.mjs` 생성
- 이 보고서 및 20개 업무 진행표

**보호 대상 원본 SHA — 변경하지 않음:**
- 2D: `LS_Datacenter_Campus.html` blob `9c84f24f2101043d7c113ce2d0da50e7245bc318`
- BOM: `index.html` blob `692767f0c234df3b8310642e3acd032137c68baa`

다음 단계 진행 게이트: 사용자 확인 후, 필요한 STEP 03 실기기 검증을 수행하고 미통과 항목을 해결하기 전에는 STEP 04에 착수하지 않는다.


## 6. 사용자 브라우저 수동 검증 후 업데이트 (2026-10-15)

앞부분의 **4-C 'NOT RUN'은 최초 작성 당시의 기준 기록**으로 보존한다. 이후 사용자가 실브라우저로 기본 6개 항목을 점검하여 **모두 PASS**를 보고했다. 이는 사용자의 실제 관찰에 근거한 결과이며, CI 로그·스크린샷이 수반된 독립 시험은 아니다.

1. 3D 데이터센터 건물·설비 렌더 표시: PASS (사용자 확인)
2. `Three.js r186 · PBR` 표시: PASS (사용자 확인)
3. 7개 구역 전환: PASS (사용자 확인)
4. 랙 클릭 및 우측 상세정보: PASS (사용자 확인)
5. 마우스/터치 카메라 회전·확대: PASS (사용자 확인)
6. 랙 문 열기·서버 인출: PASS (사용자 확인)

이후 테스트 실행 환경에서 Python Playwright/Chromium을 이용해 `?engine=three` 공개 GitHub Pages로 직접 접속했으나 `net::ERR_BLOCKED_BY_ADMINISTRATOR`로 정책상 차단되어 **자동 확인은 수행되지 않았다**. 사이트 장애로 판정하지 않는다. 콘솔·fallback·server-room 격리·BOM 정확성·연속 실행 안정성·성능을 추가 확인해야 한다.

- [브라우저 수동검증 6/6 기록, P0/P1/P2 잔여 항목, 수정 후보](./STEP03_BROWSER_VERIFICATION_20261015.md)
- STEP 03: **구현 및 핵심 수동 검증 통과, 전체 통합 검증 대기**.
- STEP 04: 미착수.
