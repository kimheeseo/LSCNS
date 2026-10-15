# STEP 18 / 20 — 카메라·모바일 UI·렌더링 성능 최적화

- 수행일: 2026-10-15 (KST)
- 저장소: `kimheeseo/LSCNS` → `main`
- 범위: **18단계만 수행.** 19단계 수용성 통합 QA 및 20단계 최종 배포 검증은 이번 작업에 포함되지 않는다.
- 판정: **개발 코드 및 소스 기반 JS 테스트 5/5 PASS / Chromium·WebGL·모바일 실화면 검증 대기**
- 접속: https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=three

## 1. 변경 전 문제

- 기존 `scene-adapter.mjs`는 low/medium/high 고정 DPR만 사용했고 프레임이 지속적으로 느려져도 해상도가 적응하지 않았다.
- 외벽/지붕·천장 객체가 근거리 장비 전면 시야에 겹칠 때 선택 장비가 가려질 수 있었다.
- Co-location/MMR의 좌상단 패널과 STEP 17 MEP 우상단 패널, 랙 조작 하단 패널이 작은 화면에서 겹칠 수 있었다.
- 모바일 Two-Finger Pointer 이동은 pinch zoom만 처리하고 두 손가락 평행 이동(pan)을 제공하지 않았으며 다중 접촉 직후 클릭 오인식 위험이 있었다.
- `clearGroup`은 캐시 Material 목록을 매 Mesh마다 Array 변환하여 검색, 씬 전환 시 불필요한 반복 작업이 있었다.

## 2. 신규 코드

`ls3d-three/scene-viewport-policy.mjs` (Three.js와 브라우저 런타임에 결합되지 않은 순수 상태/조작 모듈):

| 영역 | 구현 |
|---|---|
| 적응형 DPR | EMA frame interval 기반 sustained-slow 상태가 48프레임 누적되면 medium 모드 DPR 배율 0.77 적용, 빠른 상태가 210프레임 유지되면 복원. Slow 기준 EMA 38ms 초과, fast 기준 21ms 미만. |
| 품질 제어 | low cap=0.8, medium desktop=1.1/mobile=1.0, high desktop=1.55/mobile=1.3. 명시적 low/high 품질에 adaptive lower-tier 강제 적용하지 않음 |
| Mobile Shadows | low 모드 또는 모바일 medium의 지속적인 프레임 저하 때만 shadowMap 비활성화. 그 외 모드 유지 |
| Focus occlusion | 선택 장비가 있는 구역에서 시점–장비 간격이 24 이하일 때 직접 가리는 `arch:wall/roof/ceiling` 객체만 Raycaster로 검사하고 일시 숨김. 6프레임 간격으로 재검사하며 줌아웃·구역변경·단면·inspect·dispose 시 즉시 복원 |
| 근거리 near plane | `camera.near`를 0.055~0.15 범위에서 제한해 랙 근접 시 클리핑 안정성 확보. 멀리 떨어진 카메라 plane은 기존 `view.far` 유지 |
| UI | 640px 이하일 때 MMR/Co-location 패널 상단 높이 제한, MEP panel을 하단/랙패널 위로 이동, 랙 조작 스크롤 높이 제한. 사용자가 다시 펼친 MEP 메뉴는 계속 열 수 있도록 수동 제어 보존 |
| Multi-touch | 두 손가락 pinch zoom 및 중앙 이동에 따른 camera target pan, multi-pointer 해제 직후 잘못된 `pick()` 호출 억제 |
| Cleanup | `clearGroup`에서 material cache를 한 번만 `Set`으로 생성해 공유 Material을 반복 검색하지 않고 동적 객체 삭제 시 dispose |
| Diagnostics | `adapter.viewportPerformance`, `adapter.cameraOcclusion`, `diagnostic.viewport`, `diagnostic.occlusion`, `diagnostic.overlay`에 내부 품질·DPR·숨긴 객체·패널 상태 노출 |

**주의:** 여기서 평균 프레임 시간은 RAF 기반 `dt`의 이동평균으로 GPU frame time 실측값이 아니며, 프레임 향상률 수치나 최소 FPS를 보증하지 않는다. 건축물 투명화 대신 시점 가림 객체를 일시적으로 숨기는 **시각적 검토 용도**이며, 엔지니어링 Clash Detection 또는 법정 도면/피난 설계가 아니다.

## 3. 시행한 검증 및 판정

**GitHub에 저장된 모듈을 직접 불러 JS로 실행한 Mock/Unit 5/5 PASS**
1. 48프레임 slow hysteresis 이후 DPR 저하 / 210프레임 정상 회복 / low-high 고정 설정 보호
2. devicePixelRatio 제한, non-finite delta 입력 무시
3. 선택 장비 시점에서 지붕·벽만 일시 숨김, 바닥은 보존하고 6프레임 throttle 중 flicker 없음, 줌아웃 후 즉시 복원
4. 단면/inspect/회사 서버실/구역 전환 시 모든 occluder 복원
5. 모바일에서 zone·MEP·rack 패널 배치 / MEP 재개방 / 데스크톱 패널 스크롤 복원

코드 구현·검증 중 `room.add()` 반환 유형을 잘못 가정한 테스트 fixture 오류를 바로잡았고, 줌아웃 시 occluder가 한두 프레임 뒤 복원되는 실제 구현 문제도 수정했다.

**작성했지만 실제 실행하지 않은 Chromium Playwright 브라우저 테스트:** `ls3d-three/step18.playwright.spec.mjs`
- Data Hall 랙 가까이 보기, 카메라 clip plane 및 건물 가림 복원
- 390px 모바일, Co-location/MEP/rack 패널 표시 및 제어 도킹
- renderer DPR와 diagnostics, 7개 구역 회귀
- 전체화면 진입/복귀 및 Console 오류/스크린샷

현재 세션에서 사이트에 접속하는 독립 Chromium QA가 진행되지 않았으므로 **실브라우저 PASS·정량 FPS/메모리 개선·새 스크린샷을 주장하지 않는다**. 기존 STEP 03 사용자 수동 확인 6/6이 있었다고 해도 이 18단계 버전 수용성을 입증하지 않는다.

### 로컬에서 재실행

```bash
node --test DCI/DataCenter/ls3d-three/step18.unit.test.mjs
npm install --no-save @playwright/test
npx playwright install chromium
LS3D_BASE_URL=https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html npx playwright test DCI/DataCenter/ls3d-three/step18.playwright.spec.mjs
```

## 4. GitHub 반영 파일

- `ls3d-three/scene-viewport-policy.mjs` — 신규 viewport 품질·occlusion·반응형 UI 정책
- `ls3d-three/scene-adapter.mjs` — 렌더링 프레임/DPR/카메라/진단·패널 동작 연결
- `ls3d-three/bootstrap.mjs` — mobile multi-pointer 조작과 화면 이동, 새 adapter import 버전
- `LS_Datacenter_3D.html` — STEP 18 모듈 캐시 버전
- `ls3d-three/step18.unit.test.mjs` — 신규 독립 JS unit tests
- `ls3d-three/step18.playwright.spec.mjs` — Chromium 4개 실제 UI 수용 테스트 초안
- `architecture/STEP18_CAMERA_UI_PERFORMANCE_REPORT.md` — 본 보고서
- `LS3D_V4_20_TASK_ROADMAP.md`, `LS3D_PROGRESS_STATUS.md` — 진행 상황 갱신

원본 2D Tour `LS_Datacenter_Campus.html`와 BOM `index.html` 및 MMR/Colocation 핵심 모듈/STEP17 MEP 연결 모델은 의도적으로 변경하지 않았다.

## 5. 다음 단계 (진행하지 않음)

- **19단계:** Chromium/WebGL 브라우저, Desktop/Mobile, 7개 구역·NPC·출입·랙·MMR·Co-location·MEP, 카메라·Full Screen·모바일 터치, JS 예외/3D 화면 캡처, FPS/메모리 실측.
- **20단계:** 검증 결과 기반 마지막 오류 수정, 문서·버전·GitHub Pages 최종 배포/링크 확인.
