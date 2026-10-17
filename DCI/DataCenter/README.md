# LS Datacenter 3D Digital Twin — 사용 설명서

**표시 버전: v4.8.1 (2026-10-15)** · 기존 계산/운영 코어: v4.7.0 · Three.js r186.1 / WebGL2 / PBR

**상태:** v4.8.1은 첫 방문 가이드 투어/추천 시나리오 자동 재생, 가정값·계산식 출처 공개, 광 네트워크 상세 진입을 개선했습니다. 모든 값은 교육용 가정이며 현장 설계 검증을 대체하지 않습니다.


## 2026-10-16 · Step 19 반복 실패 복구 검증

- **통합 Chromium/WebGL QA: [#57 — 9/9 PASS](https://github.com/kimheeseo/LSCNS/actions/runs/38820787435)** (`7e3b5b9`)
- **별도 Three.js QA: [#11 — 5/5 PASS](https://github.com/kimheeseo/LSCNS/actions/runs/38819653752)** (`172d3fc`; 서로 다른 커밋에서 검증)
- **GitHub Pages: [#937 — 배포 성공](https://github.com/kimheeseo/LSCNS/actions/runs/38820786560)** (`7e3b5b9`)
- 첫 방문 안내창의 클릭 차단, 짧은 화면 랙 패널 간섭, 모바일 터치 스크롤/좌표, 서버실 3D 표시 상태 판정, QA 웹 서버 경로 및 아티팩트 저장 경로를 보완했습니다.
- **[반복 실패 원인·수정 코드·커밋·검증 이력 상세 보고서](./architecture/STEP19_QA_FAILURE_RECOVERY_20261016.md)**

> 범위: GitHub-hosted Linux Chromium/SwiftShader QA입니다. 실제 스마트폰·GPU 성능 실측, 과거 다른 Phase 검증 전체의 통과를 뜻하지 않습니다.

## v4.8.1 업데이트

- 첫 방문 시 자동 표시되는 시작 화면에서 8구역 가이드 투어, 전원→냉각→광케이블 단선 시나리오, 광 네트워크 3D 화면을 바로 실행할 수 있습니다.
- 상단 **가정·수식** 패널에서 랙 전력, 냉각, UPS, PUE 관련 입력과 중간 계산값을 확인·수정합니다. PUE는 ISO/IEC 30134-2:2026 기준 링크를 제공합니다.
- **광 네트워크 3D 보기**로 Network/MMR 구역, Spine–Leaf 경로 및 광 예산 화면에 바로 이동합니다. 광 단선 시 우회 지연/대역폭은 교육용 가정으로 표시됩니다.
- 기존 2D Gold Pixel Tour, BOM 화면 및 Three.js 모듈 소스는 변경하지 않았습니다.

## 1. 실행 링크

| 구분 | 바로가기 |
|---|---|
| **최신 Three.js 3D 탐방** | [Three.js 엔진 실행](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=three) |
| **기존 3D 렌더 복귀** | [Legacy 엔진 실행](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=legacy) |
| **원본 2D Gold Pixel Tour** | [2D 탐방](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_Campus.html) |
| **Gold Anatomy 개선 탐방** | [Gold Anatomy 실행](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_Gold_Anatomy.html) |
| **Gold Anatomy 8단계 보고서** | [개편 보고서](./architecture/GOLD_ANATOMY_8_STEP_REPORT_20261017.md) |
| **데이터센터 제품·BOM 툴** | [DataCenter 메인](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/) |
| **상세 버전업 이력 TXT** | [LS3D_VERSION_UPGRADE_HISTORY.txt](./LS3D_VERSION_UPGRADE_HISTORY.txt) |
| **20단계 릴리스 보고서** | [STEP20_RELEASE_REPORT.md](./architecture/STEP20_RELEASE_REPORT.md) |
| **단계별 검증 현황** | [LS3D_PROGRESS_STATUS.md](./LS3D_PROGRESS_STATUS.md) |

권장: WebGL2 지원 최신 Chrome/Edge. Three.js가 실행되지 않는 경우 기존 Legacy 링크로 복귀하세요.

## 2. 처음 시작하는 방법

1. 최신 3D 링크 접속 후 화면에 Three.js r186 · PBR 엔진 표시가 나오는지 확인합니다.
2. **운영 형태**에서 AI Data Center, 회사 서버실, On-Premises 중 선택합니다.
3. 왼쪽 7구역 메뉴(모바일은 **구역 선택**)로 원하는 위치를 고릅니다.
4. 화면의 장비나 랙을 클릭하면 오른쪽 설비 정보가 바뀝니다.
5. **설비 확대**, **내부 보기**를 이용해 장비를 조사합니다.
6. 작업자, 출입문, 광·전력·냉각 연결을 차례로 살펴봅니다.

| 구역 | 탐방할 설비 |
|---|---|
| Campus | 캠퍼스 전체, 출입 보안 및 외부 시설 |
| Utility Yard | 변압기, 수배전, 발전기 연계 |
| Cooling Plant | 냉각탑, Chiller, Pump, CDU |
| Gray Zone | UPS, 배터리, 분전 등 기계/전기 지원 |
| Data Hall | GPU 서버, ToR, ODF, CRAH, 고객 A/B/C 공간 |
| Network/MMR | Carrier 인입, ODF, DWDM/OTN, Core/Spine |
| Operations/Security | 경비원, 방문객, 직원 운영/출입 지원 |

## 3. 3D 카메라 조작

| 기능 | PC | 모바일 |
|---|---|---|
| 회전 | 드래그 | 한 손가락 드래그 |
| 확대·축소 | 휠 | 두 손가락 핀치 |
| 화면 이동 | Shift + 드래그 / 우클릭 드래그 | 두 손가락 함께 이동 |
| 설비 선택 | 클릭 | 탭 |
| 전체 캠퍼스 | **전체** 버튼 | **전체** 버튼 |
| 카메라 각도 | **3D** / **평면** | 같은 버튼 |
| 건물 내부 | **지붕 켜기** / **단면 보기** | 같은 버튼 |
| 탐방 | **걷기**, WASD 또는 방향키 | 기기별 조작 일부 제한 가능 |
| 화면 확대 | **전체 화면** | 브라우저 지원 시 사용 |

랙을 가까이 볼 때 일부 건축 벽/지붕을 일시적으로 숨기는 기능이 있습니다. 실제 CAD 충돌해석 기능은 아닙니다.

## 4. 랙 내부 조작

1. **Data Hall**에서 GPU Rack 또는 ToR Rack을 선택합니다.
2. 하단/우측 **랙 조작 패널**의 **전면 카메라 / 후면 카메라**를 누릅니다.
3. **전면 문 열기/닫기**, **후면 문 열기/닫기**를 클릭해 실제 3D 문 회전을 확인합니다.
4. **이전 서버 / 다음 서버**로 서버를 선택하고 **선택 서버 인출/복귀**를 실행합니다.
5. **광 카세트 인출/복귀**를 눌러 내부 광 커넥터와 케이블을 조사합니다.
6. **ToR / 후면 PDU A·B / 서버 NIC 배선** 메뉴를 열어 광/전력 배선을 선택합니다.

랙 42/48U, ToR 위치, 서버 NIC, 광 카세트, PDU A/B 모델은 설명용 형상입니다. 실제 상품/발주 BOM과 동일하다는 보장은 없습니다.

### 서버 조작 패널 접기 / 펼치기 (모바일 개선)

서버 조작 패널의 오른쪽 위에 **「접기 ▴」 또는 「펼치기 ▾」 버튼**을 추가했습니다.

- **모바일(화면 폭 640px 이하)**: 랙을 선택하면 처음에는 화면 우측 하단에 작은 `랙 조작 · 랙 ID | 펼치기 ▾` 제목 표시줄만 나타납니다.
- **「펼치기 ▾」**: 이전 서버, 다음 서버, 선택 서버 인출/복귀, 광 카세트 인출/복귀, 문 열기/닫기 및 ToR/PDU 배선 제어가 다시 나타납니다.
- **「접기 ▴」**: 버튼 목록만 숨기고 3D 랙 화면을 더 넓게 봅니다.
- **PC(화면 폭 640px 초과)**: 기본은 펼침이며 원할 때 접을 수 있습니다.
- 서버 선택/문/인출 상태는 접거나 펼쳐도 **초기화되지 않습니다**. 다른 랙을 선택해도 접기 상태를 유지합니다.

## 5. MMR / 고객별 Co-location

**Network/MMR 구역:** Carrier, DWDM Handoff, ODF Cross-connect, Core의 4개 구획을 골라 장비, 광경로, 출입문 애니메이션을 확인할 수 있습니다. 광 연결 구간을 고르면 관련 경로가 선명해지고 나머지는 흐리게 표시됩니다.

**Data Hall 구역:** 고객 A/B/C의 3개 독립 Cage(격자 펜스) 중 하나를 선택합니다. 각 고객 출입문을 열고 Fiber 연결과 Rack PDU A/B 전력 경로를 독립적으로 확인할 수 있습니다.

## 6. 전력·냉각·광 통신 경로

오른쪽 **흐름 표시**에서 전력, 냉각, 광통신 버튼을 누릅니다. Three.js 화면의 **전력 · 냉각 · 광 3D 연결 보기** 패널을 열어 계통/연결 구간을 선택할 수 있습니다. 입체 Cable Tray와 Pipe, TubeGeometry 및 장비 연결 정보를 제공합니다.

연결은 교육용 개념 경로입니다. 실제 케이블 길이, 정격 전류, 광손실, 대역폭, 냉각수 압력/유량과 일치한다고 해석하면 안 됩니다.

## 7. 장애 상황·직원·보안·내보내기

- **운영 장애 시나리오:** 정상, 전원, 냉각, 네트워크 등 이벤트를 선택하고 KPI/장비 변화를 비교합니다.
- **작업자:** 상단의 **작업자** 또는 **업무 목록 · 시점 선택**을 열어 NPC 동선/근무 상태와 카메라 시점을 확인합니다.
- **방문객 보안:** Campus에서 접수/X-ray/금속탐지/카드리더/출입 게이트 연출을 확인합니다.
- **이중화 구성:** N/N+1/2N을 선택해 가정된 운영 차이를 비교합니다.
- **CSV 내보내기:** 상단 CSV 버튼으로 현재 제공되는 데이터를 저장합니다.
- **회사 서버실:** 소규모 실내 설비만 보고 외부 캠퍼스 경로를 숨기는 별도 시설 모드입니다.

## 8. 검증 상태 및 오류 대응

현재 확인된 결과: STEP 03 사용자 수동 브라우저 기본 점검 6/6 PASS, STEP 14~18의 기존 소스 기반 검증 결과와 별도로 **2026-10-16 Step 19 통합 QA #57 9/9 PASS / 별도 QA #11 5/5 PASS / Pages #937 배포 성공**을 실제 Actions 실행에서 확인했습니다. **서로 다른 커밋의 성공 결과이며**, 실제 모바일 기기 시험·지속 FPS 개선 수치·이전 모든 단계별 QA의 합격까지 검증한 것은 아닙니다. [상세 복구 보고서](./architecture/STEP19_QA_FAILURE_RECOVERY_20261016.md)를 참조하세요.

1. [GitHub Actions 통합 품질 검사](https://github.com/kimheeseo/LSCNS/actions/workflows/ls3d-step19-quality.yml)에서 최신 실행 상태를 확인합니다.
2. PASS가 나오면 해당 Actions 실행의 Artifacts에서 화면 캡처·추적·JSON 결과를 확인합니다.
3. 빈 화면이나 오류가 발생하면 Chrome/Edge의 F12 Console과 Network에서 WebGL2, three.module.js, three.core.js, 404, CORS, MIME 오류를 확인합니다.
4. 새 버전 표시가 보이지 않으면 브라우저 새로고침 또는 캐시 삭제 후 다시 접속합니다.
5. 문제가 계속되면 [Legacy 렌더](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=legacy)를 사용합니다.

### 개발자가 로컬에서 확인할 때

저장소를 다운로드하거나 Clone하여 루트 폴더에서 다음 테스트를 실행합니다.

- Node 회귀: node --test DCI/DataCenter/ls3d-three/step14.unit.test.mjs DCI/DataCenter/ls3d-three/step15.unit.test.mjs DCI/DataCenter/ls3d-three/step16.unit.test.mjs DCI/DataCenter/ls3d-three/step17.unit.test.mjs DCI/DataCenter/ls3d-three/step18.unit.test.mjs DCI/DataCenter/ls3d-three/step19.source.test.mjs
- 브라우저 패키지: npm install --no-save @playwright/test@1.56.1 three@0.186.1
- 브라우저 설치: npx playwright install --with-deps chromium
- 정적 서버: python3 -m http.server 8765 --bind 127.0.0.1
- 다른 터미널에서: npx playwright test --config DCI/DataCenter/ls3d-three/step19.playwright.config.mjs

## 9. 버전 업그레이드 및 보호 대상

**[LS3D_VERSION_UPGRADE_HISTORY.txt](./LS3D_VERSION_UPGRADE_HISTORY.txt)** 문서에서 1~20단계별 개발 내용, 소스 폴더 구성, import 버전 쿼리/캐시 무효화, 업그레이드 절차와 롤백 방법을 확인하세요. 최종 검증의 구체적 완료 조건은 **[STEP20_RELEASE_REPORT.md](./architecture/STEP20_RELEASE_REPORT.md)**에 기록합니다.

기존 2D Gold Pixel Tour와 DataCenter BOM의 원본은 이 Three.js 릴리스 범위에서 변경하지 않습니다.

> 주의: 이 도구는 교육/시각화/설계 검토용입니다. 제조사 성능 검증, 도면/CAD/BIM 적합성, 방재·보안·전력보호·통신 광손실 검사의 대체물이 아닙니다.
