# LS Datacenter Campus 3D — 통합 개발 진행 현황

> 작성일: 2026-10-15 (KST)  
> 저장소: https://github.com/kimheeseo/LSCNS · 브랜치: `main`  
> 작업 원칙: 번호별 구현 → 실제 검증 → 완료 보고 → 사용자 확인 후 다음 번호

## 진행 상태

| 단계 | 업무 | 상태 | 검증 |
|---|---|---|---|
| 01/20 | 기존 3D 코드 분석·보호 범위 | 완료 | 정적 코드 분석 및 기준 SHA 기록 |
| 02/20 | Three.js 전환 구조 설계 | 완료 | 아키텍처·호환 API·롤백·검증 계획 기록 |
| 03/20 | Three.js 엔진·3D Scene·조명·카메라 | 개발 코드 GitHub 업로드 / **완료 판정 보류** | 구문·정적 연계·Mock 테스트 통과, **실제 WebGL2 브라우저/Playwright 테스트 미실행** |
| 04~20/20 | 고품질 건물·설비·NPC·랙·MMR·통합 배포 | 미착수 | 새 단계의 구현 및 브라우저 검증 없음 |

**완전 완료 2/20, 단계 03 실제 브라우저 검증 대기.**

## 개발 자료와 소스

- [20개 업무 로드맵](./LS3D_V4_20_TASK_ROADMAP.md)
- [02 전환 구조 설계](./architecture/STEP02_THREEJS_MIGRATION_ADR.md)
- [03 엔진 개발 및 검증 보고서](./architecture/STEP03_THREEJS_ENGINE_REPORT.md)
- [3D 웹페이지 HTML](./LS_Datacenter_3D.html)
- [Three.js Scene Adapter](./ls3d-three/scene-adapter.mjs)
- [Three.js Bootstrap/Legacy Fallback](./ls3d-three/bootstrap.mjs)
- [Playwright 브라우저 테스트 코드](./ls3d-three/step03.playwright.spec.mjs)

## 사용자 실행 링크

- 기본 WebGL: https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html
- 새 Three.js 시험 모드: https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=three
- 구형 WebGL 강제: https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=legacy

참고: 위 링크는 사이트 경로이며, 보고서 작성 현재 시점 실제 브라우저 구동 성공은 증명되지 않았다.

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

먼저 STEP 03의 실제 브라우저 검증을 완료하고 발생 오류를 수정한다. 그 결과를 사용자에게 보고한 후, 승인받을 때만 STEP 04 캠퍼스 건물 정밀 모델링으로 진행한다.
