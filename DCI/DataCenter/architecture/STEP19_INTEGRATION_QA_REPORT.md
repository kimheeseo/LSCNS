# STEP 19 / 20 — Three.js 전체 통합 검증 및 회귀 QA 결과

- 작성일: 2026-10-15 KST
- 저장소: `kimheeseo/LSCNS` / `main`
- 기준 페이지: `DCI/DataCenter/LS_Datacenter_3D.html?engine=three`
- **판정: 소스 기반 재검증 완료, 브라우저 자동 테스트 준비/CI 등록 완료. 실제 Chromium/WebGL2 실행 PASS·스크린샷 결과 검증은 아직 미확인.**
- 20단계 최종 배포 완료 판정을 하지 않음. 2D Gold Pixel Tour와 BOM은 수정하지 않음.

## 1. 이번 세션에 직접 수행한 검사

GitHub 커넥터에서 `main`에 저장된 실제 `.mjs` 구현 파일과 Node 호환 테스트 소스를 다시 가져와, JavaScript 테스트 함수를 Three.js API 호환 Mock 객체로 실제 호출하였다.

| 회귀 테스트 | 실행 결과 |
|---|---|
| STEP 14 — ToR/NIC/후면 PDU A/B | 3/3 PASS |
| STEP 15 — MMR, ODF, DWDM 및 광 회선 | 4/4 PASS |
| STEP 16 — Co-location A/B/C 독립 구획 | 5/5 PASS |
| STEP 17 — 광·전력·냉각 캠퍼스 경로 | 5/5 PASS |
| STEP 18 — Camera occlusion·DPR·모바일 HUD | 5/5 PASS |
| **합계** | **22/22 PASS** |

추가로 Three.js 실행 진입부와 핵심 구현/검증 파일 **20개의 JavaScript 구문 파싱 PASS**, 검색한 상대 경로 import의 **누락 0건**을 확인하였다.

**주의:** 위 22/22는 Node/JavaScript Mock 기반 모듈 및 상태 검증이다. 물리 GPU에서 렌더링, 실제 사용자 입력, 모바일 브라우저, 메모리/FPS를 계측한 결과가 아니다. 새로운 `step19.contract.test.mjs`는 GitHub Actions용으로 추가했으나 현재 세션에서 실제 Node CLI 실행 여부는 확인되지 않았다.

GitHub에서 원본 변경 여부를 확인한 보호 대상:
- `DCI/DataCenter/LS_Datacenter_Campus.html` Git blob SHA `9c84f24f2101043d7c113ce2d0da50e7245bc318`
- `DCI/DataCenter/index.html` Git blob SHA `692767f0c234df3b8310642e3acd032137c68baa`

## 2. 신규 Chromium/WebGL2 종합 시나리오

`ls3d-three/step19.integration.spec.mjs`에 실제 Playwright 브라우저 검증 5종을 구현했다.

| ID | 테스트 목적 | 캡처/근거 |
|---|---|---|
| S19-01 | 7개 구역 각각 Chrome Three.js r186 WebGL2 화면, 설비·렌더 draw call/triangles, 캠퍼스 보안 출입검색 및 NPC | `step19-zone-*.png` 7장, `zone-render-metrics.json` |
| S19-02 | 실제 Rack 선택, ToR-NIC·전원 A/B 3D 경로, 전면문 열기 및 서버 인출 후 상태 비교 | `step19-rack-open-nic-cables.png`, `rack-mechanics.json` |
| S19-03 | MMR/Carrier/ODF/DWDM 광회선, Co-location 고객 A/B/C, 3D 전력·냉각·광 라우팅 및 선택 | `step19-network-mmr.png`, `step19-hall-colocation.png`, `step19-campus-mep-fiber.png`, `mmr-colo-mep.json` |
| S19-04 | 폭 390px 모바일 장치 레이아웃, 조작 메뉴 노출 및 회사 서버실 외부 경로 격리 | `step19-mobile-390.png`, `mobile-layout.json` |
| S19-05 | 실 브라우저 requestAnimationFrame 90프레임 표본으로 FPS 근사값, p95 프레임 시간, 렌더 콜/메모리/지오메트리 기록 | `step19-performance-hall.png`, `chromium-performance-90-frames.json` |

측정값은 실제 브라우저에서 생성된 경우에만 해석한다. `requestAnimationFrame` 기반 FPS는 특정 호스트 성능의 참고값이며 GPU 하드웨어 시간이나 실제 현장 클라이언트의 FPS를 보증하지 않는다. 테스트의 프로그램 상 시나리오 존재와 **실제 브라우저 PASS**는 서로 다른 상태다.

## 3. GitHub Actions 도입 및 테스트 반복성

- 워크플로: `.github/workflows/ls3d-step19-qa.yml`
- Node 소스 검증 Job: STEP 14~18 기존 `*.unit.test.mjs` 및 STEP 19 원본 무결성/상대 import/7구역 계약 테스트.
- Browser Job: Playwright Chromium + WebGL2 (SwiftShader) + Three.js **0.186.1**. 테스트 서버는 `localhost:8765`, 테스트하는 HTML과 3D 모듈은 `main` Checkout의 실제 소스다.
- CDN 네트워크 외부 의존성을 줄이기 위해 **Three.js 0.186.1** CDN import 경로만 로컬 설치된 동일 npm 버전으로 Playwright `route.fulfill()` 처리하며, 실제 배포된 HTML은 변경하지 않는다.
- 실패하면 traceback/trace/video 등, 성공하면 스크린샷/JSON 측정치 등을 `ls3d-step19-qa-<run_number>` Actions Artifact(14일 보관)에 수집하도록 구성.
- `main`에 워크플로 변경 시 push trigger 및 `workflow_dispatch` 수동 실행을 허용한다.

Actions 링크: https://github.com/kimheeseo/LSCNS/actions/workflows/ls3d-step19-qa.yml

**실행 상태:** 원본 공개 GitHub Pages에 대한 이 작업 환경의 curl 접근은 DNS 조회 실패(`Could not resolve host: kimheeseo.github.io`)였다. GitHub Actions 실행 결과를 조회할 수 있는 완전한 push-workflow listing 액션도 연결 도구에서 제공되지 않아, 이 문서 작성 시점에 브라우저 Job의 시작/종료/결과 및 아티팩트 실제 생성 유무를 확인하지 못하였다. 따라서 **CI 등록은 확인됐지만 CI PASS는 기록하지 않는다.**

### 로컬 실행

```bash
node --test \
  DCI/DataCenter/ls3d-three/step14.unit.test.mjs \
  DCI/DataCenter/ls3d-three/step15.unit.test.mjs \
  DCI/DataCenter/ls3d-three/step16.unit.test.mjs \
  DCI/DataCenter/ls3d-three/step17.unit.test.mjs \
  DCI/DataCenter/ls3d-three/step18.unit.test.mjs \
  DCI/DataCenter/ls3d-three/step19.contract.test.mjs

npm install --no-save @playwright/test@^1.56.0 three@0.186.1
npx playwright install --with-deps chromium
npx playwright test --config=DCI/DataCenter/ls3d-three/playwright.step19.config.mjs
```

## 4. 이번 단계 신규 파일

- `ls3d-three/step19.integration.spec.mjs` — 실제 브라우저 E2E 테스트 5종
- `ls3d-three/playwright.step19.config.mjs` — Chromium·로컬 서버·JSON/HTML 보고서·trace 스크린샷 설정
- `ls3d-three/step19.contract.test.mjs` — 2D/BOM blob 고정, 7 구역/핵심 설비, ES import graph 검사
- `.github/workflows/ls3d-step19-qa.yml` — GitHub Actions 자동화
- `architecture/STEP19_INTEGRATION_QA_REPORT.md` — 이 문서
- `LS3D_V4_20_TASK_ROADMAP.md`, `LS3D_PROGRESS_STATUS.md` — 상태 동기화

**기존 시뮬레이터 구현을 변경하지 않았으며 STEP 18 결과물을 보존했다.**

## 5. 남은 결함 확인 / 20단계 진입 전 조건

- GitHub Actions `source-mock` Job 실제 완료와 `chromium-webgl` Job 최종 결과를 확인하고 실패 시 구체적인 stack trace 또는 스크린샷을 기반으로 수정.
- WebGL2 초기화, GPU buffer 생성, 화면 가림·랙 클릭·문·카세트·모바일 패널의 **실제 픽셀/상호작용** 확인. 7개 zone screenshot과 FPS/heap 데이터 유효성 판단.
- 보안 방문자 애니메이션, 실제 touch pinch/pan, 전체화면 복귀, 브라우저별 성능에 대한 추가 수동/자동 검사 검토. 단계 19 시나리오만으로 모든 1~18단계 기능의 완전한 수용성을 주장하지 않음.
- GitHub Pages 배포 상태와 최종 문서/SHA/사용 설명서 갱신은 20단계 작업으로 분리.

현 단계는 **19/20: 테스트 인프라 및 확인 가능한 소스 검증 완료 / 실제 브라우저 최종 수용성 판정 대기**이다.

## 6. 보완 통합 시나리오 및 별도 워크플로

GitHub 저장소의 최신 19단계 아티팩트를 다시 조사해, 초기 5종의 실브라우저 종합검사 외에 다음 **8개 독립 시나리오**를 별도 보완했다.

- `step19.integration.playwright.spec.mjs`: 실제 Chromium으로 7구역 렌더·draw call, NPC/8단계 보안, 랙 Door/NIC/ToR/PDU A/B/카세트, MMR/ODF/DWDM, A/B/C Co-location, 3D MEP 계통별 회선, 모바일 Two-Finger Pan/Pinch 및 서버실/legacy 격리를 검증하도록 구성.
- `step19.playwright.config.mjs`: 실패 시 trace/video/screenshot, 성공 시 구역별/모바일/설비 캡처와 JSON 진단을 보관하도록 구성.
- `step19.source.test.mjs`: 모듈 간 기본 데이터 계약·7구역·지정 UI 셀렉터, 14~18단계 단위테스트 및 2D/BOM Git Blob 보호를 Node에서 재검사하도록 구성.
- `.github/workflows/ls3d-step19-quality.yml`: Node STEP 14~18 + STEP 19 원본 정합성 검사, localhost 기반 headless Chromium/WebGL 테스트, `actions/upload-artifact@v4` 수집을 위한 별도 GitHub Actions 워크플로.
- Three.js CDN 요청 중 `three.module.js`와 `three.core.js`는 모두 동일 버전의 로컬 npm 패키지로 fulfillment하며, ESM cross-origin CORS 허용 헤더를 포함. 외부 CDN 접근이 실패해도 핵심 WebGL 모듈 로딩에 의존하지 않도록 테스트 구성.

**직접 수행한 독립 소스 검증**: GitHub의 14~18단계 모듈과 저장된 테스트를 다시 가져와 JS/Three API Mock 실행 **22/22 PASS**(3+4+5+5+5). 수정된 전용 Playwright 테스트 구문 검증, 모듈·UI 셀렉터 및 버전 경로 일치, 보호 Git Blob SHA와 Actions 설정을 포함한 **별도 정적 정합성 검사 15/15 PASS**.

**최종 브라우저 결과는 여전히 미확인**: `ls3d-step19-qa.yml` 및 `ls3d-step19-quality.yml` 두 워크플로는 모두 GitHub `main`에 등록되어 있지만, 이 세션에서는 push-triggered GitHub Actions 실행 로그·스크린샷 아티팩트·PASS 판정을 직접 읽지 못했다. 두 워크플로의 구성을 성공한 실제 브라우저 수용성 검증으로 간주하지 않는다.

- [추가 8개 WebGL E2E 테스트](../ls3d-three/step19.integration.playwright.spec.mjs)
- [추가 5개 원본 정합성 테스트](../ls3d-three/step19.source.test.mjs)
- [추가 CI 워크플로](https://github.com/kimheeseo/LSCNS/actions/workflows/ls3d-step19-quality.yml)

20단계 전에는 위 Actions 결과에서 **실제 실패/통과 수, WebGL 콘솔 및 캡처, FPS/메모리 측정값**을 확인하고 오류를 수정해야 한다.
