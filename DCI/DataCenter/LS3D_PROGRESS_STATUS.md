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
| 07~20/20 | 경로탐색/충돌회피·상태머신·보안·랙/MMR/최종 배포 | 미착수 | 사용자 별도 승인 전 진행 없음 |

**완전 검증 완료 2/20; STEP 03 수동 브라우저 6/6 PASS; STEP 04 Mock 12/12 PASS; STEP 05 Mock 13/13 PASS; STEP 06 NPC/애니메이션 통합 Mock 17/17 PASS. STEP 04~06의 실제 GPU/브라우저 회귀 및 성능은 대기.**

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

STEP 03 사용자 기본 동작 6/6 PASS이지만 일부 P0 회귀가 남아 있다. STEP 04 건축 Mock 12/12, STEP 05 설비/LOD Mock 13/13, STEP 06 NPC 관절 Mock 17/17 통과했다. STEP 04~06 신규 구현은 실제 WebGL2/PC·모바일 브라우저 동작과 GPU FPS/메모리/콘솔 오류 검증 대기. [STEP06 보고서](./architecture/STEP06_NPC_CHARACTERS_REPORT.md)에 따라 실제 캐릭터 움직임을 점검한 뒤 필요 결함만 보완한다. **STEP 07 NavMesh/A* 및 충돌 회피는 사용자 추가 승인 전 진행하지 않는다.**
