# 디지털 트윈 작업 화면 개선 · 2026-11-07

기준: `main` / `d34bfd0ec626313a1550bea5da9b114e21de93c6`. 구현은 기존 Three.js r186.1 및 v4.7 계산 코어 위에 추가했다. 장비 ID, 포트/케이블 데이터, BOM과 계산식은 변경하지 않았다.

## 코드 조사 및 설계

| 기능 | 조사 결과 | 처리 |
| --- | --- | --- |
| Three.js / WebGL2, 레거시 WebGL 폴백 | 유지 | 부트스트랩 계약 유지, 실제 폴백 테스트 |
| sRGB / ACES / PCFSoft 그림자 | 유지·개선 필요 | 중복 도입 없이 노출 1.28→1.10, 조도 균형 조정 |
| 랙·서버·ODF·커넥터 생성, 실제/개념 경로 표기 | 유지 | 기존 factory, rack kit, 케이블 선택 재사용 |
| 전면/후면 힌지, 서버 레일, 카세트 인출 | 유지 | 기존 기계 상태와 애니메이션 루프 재사용 |
| 카메라 보간, 회전/이동, 휠, 키보드·터치 | 유지 | 기존 host API 연결, 장비 클릭 초점 이동 |
| 투시 모드 | 미구현 | 선택 랙의 문 재질만 복제, 해제/랙 변경/종료 시 복원 |
| 좌우 패널 실제 접기, 장비 계층 탐색 | 개선 필요 | CSS grid track 축소, 구역 내 장비 목록, 탐색 경로, 기존 부품 선택기 유지 |
| 저/표준/고 품질, 캐시·InstancedMesh·LOD | 유지·개선 필요 | 기존 DPR·적응형 정책 유지, 그림자 0/1024/2048, 비선택 랙 LOD 55/96/135 |
| 추가 AO/후처리 | 미구현 | 소프트웨어 렌더링 비용을 고려해 추가하지 않음 |
| 실제 CAD/현장 토폴로지/실측 광학 파라미터 | 확인 불가 | 데이터 생성 없이 기존 개념 모델 및 설명 유지 |

배경은 네이비/차콜, 강조는 청록색, 기존 Pretendard와 한국어 시스템 폰트 폴백을 유지한다. 운영 설정을 접어 3D 공간을 확보하되 버튼으로 모든 기존 정보를 다시 연다. 패널 접기는 해당 영역을 inert 처리해 키보드 포커스가 숨겨진 패널로 들어가지 않는다. 모바일은 탐색 서랍으로 제공한다.

선택 강조는 선택된 부품의 재질 복제만 변경한다. emissive 강도를 낮추고 선택 해제 시 원본을 복원한다. 문은 기본 불투명성을 유지한다. 투시는 힌지 각도·문 상태·가시성을 바꾸지 않으며, 공유 재질의 이웃 장비도 바꾸지 않는다. 새 독립 렌더링 루프는 없다.

주요 병목은 캠퍼스의 많은 draw calls와 매 프레임 그림자 렌더링이다. 기존 geometry/material 캐시·반복 객체 인스턴싱은 유지했다. 고품질은 픽셀 및 그림자 비용이 증가하며 저사양에서 성능을 보장하지 않는다. 성능 HUD의 G/T는 Three.js geometry/texture **객체 수**이며 GPU 메모리 바이트가 아니다. FPS는 실제 경과시간과 렌더 프레임 수로 계산한다.

## 실행 및 검증 명령

저장소 루트에서 `python3 -m http.server 8765 --bind 127.0.0.1` 실행 후 브라우저에서 `/DCI/DataCenter/LS_Datacenter_3D.html?engine=three` 경로를 연다. `?engine=legacy`는 원래 렌더러를 연다. HTML/CSS/ES module의 상대경로를 유지하므로 GitHub Pages의 `/LSCNS/` 하위 경로와 호환된다. 라이브 Pages 배포는 수행하지 않았다.

브라우저 주소: `http://127.0.0.1:8765/DCI/DataCenter/LS_Datacenter_3D.html?engine=three`.

검증 의존성: Node 24.22.1, Python 3.12.14, `@playwright/test@1.56.1`, `three@0.186.1`.

Node/Python 외 별도 애플리케이션 서버나 유료 API가 필요하지 않다. 테스트 도구는 저장소 밖 prefix에 설치했고 루트 node_modules는 그 경로에 대한 로컬 symlink다. 의존성 선언/lockfile은 추가·변경하지 않았다.

```bash
node --test DCI/DataCenter/ls3d-three/rack-xray.test.mjs
node --test DCI/DataCenter/ls3d-three/step{14,15,16,17,18}.unit.test.mjs DCI/DataCenter/ls3d-three/step19.source.test.mjs DCI/DataCenter/ls3d-three/step20{.release,.rack-panel.unit}.test.mjs
# 기존 파일명은 step19.playwright.config.mjs이다.
LS3D_CHROMIUM_EXECUTABLE=/usr/bin/chromium LS3D_NO_VIDEO=1 LS3D_VERIFIED_CDN=1 \
  npx playwright test --config DCI/DataCenter/ls3d-three/step19.playwright.config.mjs --timeout=120000
```

`LS3D_VERIFIED_CDN=1`은 클라우드 Chromium이 프록시 CA를 신뢰하지 못할 때만 사용하는 테스트 전용 전송이다. 시스템 CA로 검증하는 curl로 공개 CDN 자산을 읽어 동일 바이트를 브라우저에 전달한다. 인증서 검증을 끄거나 사용자 브라우저 신뢰 저장소를 변경하지 않는다. 기본 CI에는 적용되지 않는다. 시스템 Chromium 대신 Playwright bundled Chromium을 사용하면 executable 변수를 생략한다. video 옵션은 bundled ffmpeg가 없을 때만 끈다.

## 브라우저 비교 자료

1440×900, 고품질(적응형 DPR 변경 없음), 동일 시설 데이터·기준 화면에서 저장한 카메라 위치/시야, 시뮬레이션 정지 조건의 실제 Chromium 스크린샷이다. UI 변경으로 캔버스 영역은 커졌다. 픽셀 단위 그래픽 성능 비교로 해석하지 않는다. 원본 파일은 `before-*.png`, 수정 후는 `after-*.png`다.

| 화면 | 변경 전 | 변경 후 |
| --- | --- | --- |
| 캠퍼스 | [before](before-campus.png) | [after](after-campus.png) |
| 서버실 | [before](before-hall.png) | [after](after-hall.png) |
| 랙 정면 | [before](before-rack.png) | [after](after-rack.png) |
| 문 개방·내부 | [before](before-open.png) | [after](after-open.png) |
| 케이블 선택 | [before](before-cable.png) | [after](after-cable.png) |
| 장비 정보 | [before](before-selected.png) | [after](after-selected.png) |

`capture.cjs before|after`는 로컬 실행 서버와 설치된 Playwright를 이용한 촬영 절차다. `benchmark.cjs before|after`는 표준 초기 상태로 Data Hall에서 각 품질의 실제 8초 측정 구간을 수집한다. benchmark before는 기준 SHA의 HTML/bootstrap/adapter 원본을 read-only Git 조회하여 브라우저에 제공한다. 카메라 원본은 before-camera.json에 기록한다. 성능 측정은 전후 모두 캔버스를 850×500으로 고정한다.

## 변경 파일

| 파일 (`DCI/DataCenter/` 기준) | 상태 / 내용 |
| --- | --- |
| `LS_Datacenter_3D.html` | 수정: stylesheet·캐시 버전, 기존 루프의 UI 동기화 호출 |
| `ls3d-three/bootstrap.mjs` | 수정: 작업 패널 초기화, adapter 캐시 버전 |
| `ls3d-three/scene-adapter.mjs` | 수정: 조명·선택·품질·LOD·FPS·투시 통합 |
| `ls3d-three/workspace.mjs` | 신규: 탐색, 접이식 패널, 경로, 카메라 도구, HUD |
| `ls3d-three/workspace.css` | 신규: 산업용 작업 화면 및 모바일 서랍 |
| `ls3d-three/rack-xray.mjs` | 신규: 기계 상태와 분리한 재질 복제/복원 |
| `ls3d-three/rack-xray.test.mjs` | 신규: 공유 재질, material 배열, 복제 해제 검증 |
| `ls3d-three/workspace.playwright.spec.mjs` | 신규: 3 해상도, 투시/연결 보존, 폴백, Pages 하위 경로 |
| `ls3d-three/browser-fixtures.mjs` | 신규: 선택적인 시스템 CA 기반 공개 CDN 전송 |
| `ls3d-three/step19.integration.playwright.spec.mjs` | 수정: fixture 재사용, 제거된 2D 링크 대신 실제 레거시 기하 렌더링 검증 |
| `ls3d-three/step19.playwright.config.mjs` | 수정: 새 테스트 포함, 선택적 시스템 Chromium·video 환경 옵션 |
| `ls3d-three/step20.release.test.mjs` | 수정: 문서 버전은 고정된 4.8.1 대신 현재 4.8.x 검증 |
| `architecture/workspace-20261107/` | 신규: 보고서·실제 스크린샷·카메라 상태·측정 및 재현 스크립트 |

과거 테스트는 root README의 버전, 이미 삭제된 `#back2D` 링크, 현재 기본 회전 모드를 반영하지 못했다. 현재 버전·실제 레거시 기하 렌더링·기본 회전 확인 후 명시적 이동 모드 선택으로 갱신했다. 하역장 세계 좌표를 읽기 전 전체 world matrix를 명시적으로 갱신해 숨겨진 LOD 부모 행렬도 동기화했다. 기존 비교 assertion은 유지했다. 링크를 새로 만들거나 테스트를 생략하지 않았다.

## 검증 범위와 제한

기본 재질, 문·레일·카세트·GPU, NPC/견학, MMR, 광/전력/냉각 경로, 시설 전환과 레거시 동작은 기존 STEP19 테스트로 검사한다. 새 테스트는 패널 접기 후 실제 canvas 폭 증가, 모바일 가로 넘침 방지, 품질별 실제 그림자 크기, 투시 후 문 상태 및 flowPaths/cableItems 동일성, 자동 폴백을 검사한다. `/LSCNS/` 프리픽스 경로는 로컬 정적 서버를 통해 로딩 검증한다.

실제 GitHub Pages 서비스에 배포하지 않았으므로 배포 CDN 캐시 및 라이브 Pages 결과는 검증 범위 밖이다. 실물 제조사 CAD 정확도, 실측 현장 연결, 실기기 모바일 GPU/FPS, 네이티브 전체화면 전환의 실기기 UX는 검증하지 않았다. 추가 AO/후처리도 도입하지 않았다. 실무 설계·현장 검증을 대체하는 모델로 표현하지 않는다.

브랜치: `feat/digital-twin-workspace` (`main` 직접 커밋/PR 병합/Pages 배포 없음). PR 생성용 GitHub API는 네트워크 정책에서 차단되어 `api.github.com` 도메인을 설정 초안에 추가했다. Git HTTPS push는 성공했으며 API 접근 허용과 별개다. API를 허용한 뒤 Draft PR 생성을 재시도해야 한다.

## 성능 실측

Linux x86_64, Xeon Platinum 8573C 5 vCPU, Chromium 156.0.8078.79, headless + ANGLE SwiftShader 실행 인자. 다른 테스트 브라우저와 동시에 실행하지 않았다. 브라우저 1440×900, 캔버스 CSS 850×500(테두리 제외 내용 848×498), Data Hall, 동일 데이터 및 시뮬레이션 정지. 각 모드 3초 준비 후 실제 8초 이상 경과시간/렌더 프레임 수로 FPS를 구했다. 장비의 LOD를 포함한 전체 렌더 비용이며 실기기 GPU 성능으로 일반화하지 않는다.

| 상태 | 품질 | 실제 DPR | FPS | Draw calls | Triangles | Geometries | Textures | console/page 오류 |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| before | low | 0.80 | 1.10 | 19293 | 346058 | 949 | 204 | 0 |
| before | medium | 1.00 | 1.07 | 21439 | 398482 | 949 | 204 | 0 |
| before | high | 1.00 | 1.04 | 21439 | 398482 | 949 | 204 | 0 |
| after | low | 0.80 | 11.37 | 2492 | 64788 | 126 | 43 | 0 |
| after | medium | 0.77 | 1.22 | 21439 | 398482 | 947 | 204 | 0 |
| after | high | 1.00 | 1.07 | 21439 | 398482 | 947 | 204 | 0 |

원시 값: [before-metrics.json](before-metrics.json), [after-metrics.json](after-metrics.json). Draw calls/triangles/memory는 측정 구간 마지막 실제 프레임의 renderer.info 값이다. FPS는 구간 평균이다. G/T는 메모리 바이트가 아닌 객체 수다. 반복 객체의 draw call 부담은 여전히 크며, 고품질 그림자 비용과 함께 추가 성능 작업이 필요한 지점이다.

표준 모드는 기존 적응형 정책을 유지하므로 수정 후 측정에서 DPR이 0.77로 낮아졌다. 표준 FPS 차이를 동일 픽셀 수의 그래픽 개선으로 해석하지 않는다. 저사양 DPR은 전후 모두 0.8이며 draw calls는 약 87.1% 감소했다. 고품질 DPR은 전후 1.0이다.

## 최종 검증 결과

- 소스/단위 테스트: **39 passed, 0 failed**.
- 브라우저 전체 21개 실행: 처음 **18 passed, 3 failed, 0 skipped**. 실패는 위 레거시 링크/월드 행렬/기본 이동 모드 검사였다.
- 갱신 후 해당 세 항목만 재실행: H **1 passed**; K 및 마지막 J **2 passed**. 최종적으로 **21개 고유 시나리오 모두 통과**했다. 한 번의 전체 실행이 21/21로 끝났다고 주장하지 않는다.
- 새 6개 브라우저 테스트에는 1440×900, 1920×1080, 390×844, 투시/도어/연결 데이터, 명시적 레거시 및 강제 폴백, `/LSCNS/` 하위 경로가 포함된다.
- 전후 6종 **12장** 실제 스크린샷, 해상도별 작업 화면 **3장**, 카메라 상태 및 전후 성능 JSON을 보관했다.

실행 집계: [validation-summary.json](validation-summary.json). 화면 크기별 결과: [1440](workspace-1440.png), [1920](workspace-1920.png), [390](workspace-390.png).
