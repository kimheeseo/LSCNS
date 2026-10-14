# STEP 03 — Three.js 브라우저 검증 결과 및 잔여 조치 (2026-10-15)

> 저장소: `kimheeseo/LSCNS` · 브랜치: `main`  
> 사용자 수동 검증 결과에 따른 최신 기록. 본 기록은 STEP 04 구현 지시가 아니다.  
> 상태: **3번 기능 구현 + 기본 수동 브라우저 6/6 통과; 자동 회귀/성능 및 일부 기능 검증 대기**

## 1. 수동 브라우저 검증 결과

다음 6개 PASS는 **사용자가 실제 브라우저에서 직접 확인했다고 보고한 사항**이다. 제3자 CI/Playwright 로그 또는 화면 캡처가 첨부된 결과는 아니다.

| No. | 사용자 확인 내용 | 결과/근거 수준 |
|---|---|---|
| M01 | 3D 데이터센터 건물과 설비 표시 | PASS (사용자 직접 확인) |
| M02 | 화면의 `Three.js r186 · PBR` 표시 | PASS (사용자 직접 확인) |
| M03 | 7개 구역 이동/전환 | PASS (사용자 직접 확인) |
| M04 | 랙 클릭 → 설비 상세 정보 | PASS (사용자 직접 확인) |
| M05 | 마우스/터치로 카메라 회전·확대 | PASS (사용자 직접 확인) |
| M06 | 랙 도어 열기, 서버 인출 | PASS (사용자 직접 확인) |

**판정:** STEP 03의 기본 3D 엔진 초기화·3D 표시·필수 조작 동작을 사용자 환경에서 확인하였다. PBR 라벨은 초기화의 관찰 증거이며, 실제 장치의 FPS나 GPU 소비량, 모든 PBR 설정이 정확하다는 성능 증거는 아니다.

## 2. 독립 브라우저 실행 시도 및 한계

- 검증 방식: 테스트 실행 환경의 Linux Chromium / Python Playwright에서 실제 GitHub Pages `?engine=three` URL로 `page.goto(...,wait_until='domcontentloaded',timeout=15000)` 요청.
- 관찰: `net::ERR_BLOCKED_BY_ADMINISTRATOR`, 브라우저 URL 접속 단계에서 차단.
- 판정: **NOT RUN / ENV BLOCKED**. 사이트의 404, 코드 예외, Three.js 렌더 오류로 판정하면 안 된다.
- Playwright 테스트 구현 파일 [`step03.playwright.spec.mjs`](../ls3d-three/step03.playwright.spec.mjs)은 저장되어 있으나 해당 원격 사이트에 대한 **성공한 자동화 실행 로그·스크린샷은 현재 없다**.
- GitHub HTML/JS 문법 검사 및 Three API Mock 검사 PASS는 기존 [STEP03 보고서](./STEP03_THREEJS_ENGINE_REPORT.md)를 따른다. 실제 GPU 독립 실행 성공을 대체하지 않는다.

## 3. 추가 검증 우선순위

| ID | 우선도 | 테스트 항목 | 구체적 합격 기준 | 상태 |
|---|---|---|---|---|
| R01 | P0 | Console / Network | Three.js 렌더 실행 중 예외·404/CORS·모듈 로드 실패 0건. 수동 콘솔 로그 확인 | 미확인 |
| R02 | P0 | Legacy fallback | `?engine=legacy`에서 기본 장면/클릭/시뮬레이션 정상. `?engine=three`에서 fallback 유도 시 기존 canvas로 돌아오고 중복 RAF/입력 없음 | 미확인 |
| R03 | P0 | Server-room 격리 | `facilityMode=server-room`, `hall`: 외부 건물/설비/광·전력·냉각 경로 0건, 작업자 표시 범위 확인 | 미확인 |
| R04 | P0 | 기존 데이터와 BOM | `T.exportConfig().schemaVersion === 'lsdc-twin-bom/1.0'`; 동일 입력에 legacy vs three BOM 스냅샷(시간/생성일 제외) 및 주요 수치 동일 | 미확인 |
| R05 | P1 | 랙 반복 동작 | 문 열고/닫고, 서버·카세트 인출/복귀, 후면 시점, 모달 종료 후 원위치, 반복 20회 입력 오류 없음 | 부분 확인(M06만) |
| R06 | P1 | 모드·시나리오 회귀 | `aidc`/`server-room`/`onprem`, 정전/냉각/네트워크 장애·복구 및 NPC 작업자 시점·방문객 절차 동작 | 미확인 |
| R07 | P1 | 안정성 | 모바일·데스크톱에서 10분 연속 실행 후 프레임·메모리/장비 수 지속 관찰, freeze·증가형 자원 누수 없음 | 미확인 |
| R08 | P1 | Playwright 실제 실행 | 7구역·설비·서버실·레거시·모바일 테스트, screenshot·에러 로그 확보 | 스크립트 준비/실행 환경 차단 |
| R09 | P2 | 품질/성능 벤치마크 | 동일 설정·브라우저·장치에서 legacy vs three 중간 FPS, p95 frame time, draw calls, GPU/메모리 비교표 | 미확인 |
| R10 | P2 | 가용성 | GitHub Pages·Three.js CDN 모듈 요청 200 응답과 캐시 버전 일치 | 사용자 기본 3D 사용은 확인, 개별 네트워크 로그 없음 |

P0 = STEP 03 최종 완료 판정 전 우선 확인; P1 = 기능 회귀/안정성; P2 = 정량 성능·유지보수.

## 4. 소스 정적 조사로 찾아낸 **수정 후보** (확정 장애와 구분)

1. **P1 잠재적 메모리 누수** — `ls3d-three/scene-adapter.mjs`의 `personMesh()`는 NPC별 `SphereGeometry`를 생성한다. `refreshPeople()`가 NPC를 화면에서 빼는 과정에서 `peopleRoot.remove(g)`만 호출하고 그 Geometry를 즉시 `dispose()`하지 않는다. 모드/구역 반복 전환 시 GPU 리소스 증가 여부를 R07로 측정 후 필요 시 제거 이벤트에 dispose 적용. **정적 코드상 관리 누락 확인, 실측 메모리 누수는 미확정**.
2. **P1 복구 검증 부족** — `bootstrap.mjs`는 import 실패 및 렌더 예외 fallback을 처리하지만 WebGL context-lost 이벤트에 대한 명시적인 `webglcontextlost`/복구 처리가 보이지 않는다. 인위적 context loss 테스트 후 필요 시 listener 및 복귀 UI 보완.
3. **P1 확장 형상 호환성 공백** — 기존 `LS3D_RESPONSE_VISUALS` 및 Phase2B `LS3D_PHASE1_RENDER`의 `ctx.box/line/cylinder` 관련 부가 3D 연출은 현재 새 Three renderer에서 동일 수준으로 재생성되는지 검증되지 않았다. 광/전력/냉각/대피 경로가 사라지지 않았는지 R06에서 비교. 미구현이면 STEP 03 호환성 패치 또는 관련 후속 단계에 명시적으로 이관.
4. **P2 모델 정밀도는 별도 범위** — 기본 박스·랙·문·인출 표현의 실제 CAD 치수, 현실 건물 벽·출입문 모델, NPC 경로 충돌 회피 등은 기존 로드맵 STEP 04~17 대상이다. 이번 3번 테스트에서 미구현을 무조건 결함으로 간주하지 않는다.

**이번 기록 작성 시 사용자가 기능 오류를 보고한 항목은 없다.** 따라서 확인 없는 본 프로그램 수정은 수행하지 않고 우선 재현·원인 확인 후 STEP 03 범위에 한정해 수정한다.

## 5. 현장 콘솔 확인 스니펫 (읽기 전용)

다음 코드를 `?engine=three` 페이지의 Chrome F12 → Console에 붙여넣어 P0 일부 검증 정보를 수집할 수 있다. 설비 상태나 시나리오를 변경하지 않는다.

```js
(() => {
  const S = window.LS3D_THREE;
  const T = window.__LS3D_TEST__;
  let schema = 'unavailable';
  try { schema = T?.exportConfig?.()?.schemaVersion ?? 'unavailable'; } catch(e) { schema = 'error: ' + e.message; }
  return {
    engine: S?.status, threeRevision: S?.adapter?.THREE?.REVISION,
    frames: S?.renders, diagnostics: S?.diagnostics,
    zones: T?.Z?.length, currentZone: T?.currentZone,
    facilityMode: T?.facilityMode, assetCount: T?.visibleAssets?.().length,
    bomSchema: schema,
    pageError: T?.error || null
  };
})();
```

필요 시 위 결과와 Console 빨간 오류의 첫 10줄, 테스트 화면 스크린샷을 별도 보존한다. 공개 저장소에는 사내 계정/개인정보/접근 토큰/기밀 설비 수치가 포함되지 않도록 주의한다.

## 6. 단계별 승인 상태

- STEP 01: 완료.
- STEP 02: 완료.
- STEP 03: **구현 완료, 사용자 핵심 수동 테스트 6/6 PASS, P0 회귀 및 독립 브라우저 자동 시험 대기.**
- STEP 04: **착수하지 않음**. 사용자 승인 이전 신규 고품질 건물 모델링을 시작하지 않는다.

개발 진행 종합: [`LS3D_PROGRESS_STATUS.md`](../LS3D_PROGRESS_STATUS.md), [20단계 로드맵](../LS3D_V4_20_TASK_ROADMAP.md).
