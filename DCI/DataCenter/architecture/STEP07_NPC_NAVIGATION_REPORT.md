# STEP 07 / 20 — NPC 경로 탐색·충돌 회피 개발 및 검증 보고서

> 작성일: 2026-10-15 (KST) · 저장소: `kimheeseo/LSCNS`, 브랜치 `main`  
> 상태: **실제 JavaScript A* 경로 탐색 및 구역/설비/문/작업자 회피 연동 코드 구현, 모의 통합 시험 16/16 PASS. 실브라우저 GPU/Playwright 검증은 미실행.**  
> 범위: STEP 07만 구현; STEP 08 정식 점검·장애 대응 상태머신 미착수.

## 1. 목표와 변경 요약

기존 `routeForWorker(n)` 순찰 경로는 선분 사이 장비/벽체를 뚫고 지나갈 수 있었다. STEP 07은 구역의 실제 `Z.pos/size`, `assets` 장비 footprint, 기존 NPC 순찰 좌표를 이용해 2D A* 격자에서 최소 이격 경로를 생성한다.

- `ls3d-three/npc-navigation.mjs`: **신규** A* 그리드, 차단 셀, 대각선 모서리 통과 방지, 구간 단축(직선 충돌 검사), 캐싱, 문 통과, 근거리 타 NPC 회피, 이동 거리 상한, 진단.
- `ls3d-three/scene-adapter.mjs`: 신규 Navigation 객체 생성 및 공개, STEP 06 NPC 모델과 동일 렌더 경로 유지. 통계 `window.LS3D_THREE.adapter.navigation.diagnostic()` 제공.
- `LS_Datacenter_3D.html`: Three 엔진이 `ready`인 경우에만 `npcPose(n)`에 Navigator를 적용. **기존 1인칭·3인칭 추적 카메라, 직원 HUD, Three 캐릭터가 동일한 이동 좌표를 사용**하도록 연계. `__LS3D_TEST__.routeForWorker`만 추가 공개.
- `ls3d-three/bootstrap.mjs`: ESM 캐시 키 `20261015-step07` 반영.
- `ls3d-three/step07.playwright.spec.mjs`: 실제 브라우저 회귀 테스트 저장. **실행 PASS 로그 없음**.

**기존 legacy `?engine=legacy`는 기존 경로를 보존하며 변경된 Navigator는 새 Three.js 모드에만 적용된다.**

## 2. 핵심 알고리즘

### 2-A. 장애물 구성

- 셀 크기: **1.25 개념 좌표 단위**. 보행자 반경 **0.39 단위**.
- 차단 구간: 설비/랙 중심 `(x,z)` + 폭·깊이 `(w,d)` 및 보행 이격 공간.
- Campus: 기존 건물 footprint를 통과하지 않도록 외부 장애물 처리.
- 실내 Gray/Hall/Network/OPS: 벽체 둘레에 차단 띠를 생성하고 정면 출입문 중심 약 3.3 단위의 통로만 개방.
- Utility/Cooling: 개방형 야드로 구분하여 설비 footprint 회피. 실제 모든 울타리 기둥/배관은 아직 경계 입력에 포함하지 않았음.
- 회사 서버실: 기존 `hall` 위치를 중심으로 **26 × 24 개념 경계** 안에만 경로 계산; 캠퍼스 건축/설비 경로 제외.
- 보안요원 정지 좌표와 타 직원의 최근 계산 위치: 로컬 회피 대상.

### 2-B. A* 탐색

1. 소스 `routeForWorker(n)` 경유점·시설 모드·설비 배치를 읽고 장애물 grid 구성.
2. 시작/종료 경유점이 차단 셀에 있으면 가장 가까운 접근 가능한 셀로 스냅.
3. 8방향 이웃 격자 탐색. 수평/수직 이동 비용 1, 대각선 √2; 코너를 가로질러 장애물을 관통하는 대각선은 금지.
4. 이진 최소 힙 우선순위 큐 및 유클리드 휴리스틱 적용. 검색한 경로를 안전 직선 검사로 간략화.
5. 우회 구간의 실제 길이를 기준으로 동작 속도 `n.speed/8`, 기존 경유점 정지 시간 2.3을 재구성.
6. `walk` / `inspect` / `pause` 동작 분류를 STEP 06 관절 애니메이션에 전달.
7. 다른 NPC와 가까워지면 이격 위치 후보를 검사하고 가능하면 우회, 불가능하면 일시 정지. 이미 계산된 프레임과 직전 프레임 간 이동 거리를 제한.

**법적/안전 보증이 없는 교육용 경로 탐색**이다. 고밀도 군중·3D 높이 충돌·실시간 여러 층·역학 기반 ORCA/RVO 및 NavMesh 지형 이음새 등은 구현하지 않았다. 좁은 통로나 완전 봉쇄 시 회피가 실패할 수 있으며 이때 정지 동작을 우선한다.

### 2-C. 출입문 연결

AI/On-Prem Data Hall 직원은 기존 순찰 루트 앞뒤에 Hall 전면 출입 포털(중앙 `x=hall.pos[0]`, 정면 `z=hall.pos[1]+hall.size[1]/2`)을 경유하도록 **새 Three.js 경로에서만 확장**한다. 원본 `n.route`는 수정하지 않는다.

STEP 04 건물의 전면 문 모델은 Hall 구역 진입 시 자동으로 열리므로 3D NPC 이동 경로가 문 통로를 통과하도록 연계된다. **문짝 접촉 역학·모든 문/출입통제 권한·실제 출입 절차는 본 단계가 아니다.**

### 2-D. 기존 화재 대피 시나리오

현 웹앱의 화재 대피 모드는 별도 `npcPose` 래퍼로 기존 대피 좌표를 우선 적용한다. STEP 07의 A* 평시 루트가 이를 덮어쓰지 않도록 유지했다. **화재·피난 시 장애물 회피/연기 위험/인명 안전 검증은 수행하지 않음**; 향후 별도 시나리오 검증이 필요하다.

## 3. 소스/정적·Mock 실행 검증

### 3-A. 실제 운영 자산 기반 A* 경로 테스트

GitHub `LS_Datacenter_3D.html`의 데이터 선언부를 실제 JavaScript로 실행해 **구역 7개, 자산 74개**를 불러왔다. 이 위치·폭·깊이를 사용해 AI 캠퍼스, 회사 서버실, On-Premises의 직원 경로 9개 조합을 계산했다.

| 모드 | 작업자 경로 | A* 실패 구간 | 충돌 선분 검사 |
|---|---|---:|---|
| AI | campus, hall, ops, utility | 0 | 4/4 PASS |
| 회사 서버실 | hall | 0 | 1/1 PASS |
| On-Prem | campus, hall, ops, utility | 0 | 4/4 PASS |
| 합계 | **9개 경로** | **0** | **9/9 PASS** |

기준: 생성된 각 경로의 선분 샘플 검사에서 해당 navigator의 차단 셀/장비 장애물 영역 통과가 없어야 한다. 실기기 GPU 픽셀 확인이나 제품 CAD 정합성 테스트는 아니다.

### 3-B. 랙과 다른 보행자가 마주치는 별도 모의 시나리오 — 7/7 PASS

서로 반대 방향으로 이동하는 NPC 2명을 한 통로에 두고 중앙 랙을 배치하여 `0~64.5` 시뮬레이션 시간을 0.5 간격으로 검사했다.

- A* 우회 경로 2개 생성.
- 경로 전 구간의 시작·끝·선분 모두 장비 정적 충돌 없는지 확인.
- 각 시간 샘플에서 캐릭터의 장애물 통과 없음.
- 최소 직원 간 거리 **약 0.984 개념 단위**, 설정된 단일 인원 반경 0.39의 두 배인 0.78보다 큼.
- 샘플 간 최대 한 명 이동량 **약 0.43 단위** (원격 모의 시뮬레이션 조건).
- 근거리 경로 보정 카운트 **20회**. 이는 이 테스트에서만 얻은 수치로 실시간 성능 자료가 아니다.
- 시간 연속성과 갑작스러운 텔레포트 방지 검사 통과.

### 3-C. STEP 03~07 통합 Three.js API 모의 실행 — 16/16 PASS

실제 GitHub HTML에서 자산 74개 및 신규 A*/NPC/건축/설비/Scene Adapter 소스를 모두 재조회하여 모의 THREE API를 이용해 통합 렌더 프레임 실행.

| 핵심 확인 | 결과 |
|---|---|
| AI Campus 7구역, 자산 표시, NPC 6인 | PASS |
| A* Navigator 생성 및 경로 실패 없음 | PASS |
| 기존 STEP 06 관절 캐릭터 유지 | PASS |
| 렌더러 resize 매 프레임 반복 없음 | PASS |
| Hall 진입 후 NPC 1명, 문 통과 경유점 | PASS |
| 회사 서버실 단일 구역/직원 1명/외부 지형 제외 | PASS |
| AI Campus 복귀 시 NPC 6명 복원 | PASS |
| 원본 74개 자산 데이터 유지 | PASS |
| Adapter dispose 정상 | PASS |

**16개 세부 테스트 전부 PASS.** 위 모의 실행은 실제 WebGL2/브라우저 물리 렌더링이 아니다.

## 4. 실제 브라우저 검증 — 아직 수행되지 않음

새 페이지 링크:

https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=three

실제 Chrome/Edge에서 강력 새로고침 후 아래를 확인해야 최종 합격 판정을 내릴 수 있다.

1. `Three.js r186 · PBR` 표시 및 렌더 에러 0개.
2. Data Hall 작업자가 건물의 **정면 출입문**을 통과해 내부를 이동하는지 확인.
3. Hall 랙 사이에서 몸체가 랙을 뚫지 않고 우회하는지 확인.
4. 복수 NPC가 가까워질 때 서로 겹치지 않고 방향을 보정하거나 잠시 멈추는지 확인.
5. 회사 서버실에서 외부 직원·설비 없음, On-Prem에서도 직원 경로 정상.
6. 작업자 1/3인칭 시점·직원 HUD 위치가 Three NPC 모델과 일치하는지 확인.
7. 정전/냉각/화재 대피 시나리오·랙 도어·설비 상세·Legacy `?engine=legacy` 회귀 확인.
8. PC/모바일 Chrome의 Console 오류, 10분 안정성, FPS 및 메모리 측정.

브라우저 자동화 테스트 파일: [`step07.playwright.spec.mjs`](../ls3d-three/step07.playwright.spec.mjs).

```bash
npm install --no-save @playwright/test
npx playwright install chromium
npx playwright test DCI/DataCenter/ls3d-three/step07.playwright.spec.mjs
```

**스크립트를 GitHub에 저장하는 것과 Chromium에서 실행해 PASS 로그·스크린샷을 얻는 것은 별개**다. 현재는 스크립트가 저장됐으나 실제 Playwright 성공 로그는 확보하지 못했다.

## 5. 개발 경계 및 다음 단계

- STEP 07: A* 기반 직원 경로·설비/벽 장애물·출입 통과/근거리 타 직원 회피를 **코드로 구현, 정적·Mock 검증 통과**. 실제 브라우저 사용성·충돌·속도·GPU 품질 확인 필요.
- STEP 08: 점검 절차 이동→정렬→설비 관찰→태블릿 조작→이동, 직무별 대응 상태머신. **착수하지 않음**.
- 2D Tour `LS_Datacenter_Campus.html`, BOM `index.html` 및 관련 전력 계산/제품 데이터는 변경하지 않는다.

결론: **7번 프로그램 코드/모의 기능 검증 완료, 실제 브라우저 최종 검증 대기.**
