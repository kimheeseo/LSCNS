# Industrial Twin Studio · v4.8.9

기준 코드는 PR #16이 병합된 main `dd45db8765752f1b19d9096e9e38d7f2dc73678c`이다. 이 브랜치는 첫 화면의 정보 구조와 실제 WebGL 재질/조명/구조 표면을 개선한다. 건물·장비의 실제 수량, ID, 연결, 계산식과 BOM은 유지한다. 실제 제조사 CAD 교체나 센서 연동을 추가한 것은 아니다.

## 화면에서 보이는 변화

- 넓은 3D 작업 영역: 큰 상단 KPI 카드와 운영 정보가 차지하던 공간을 정리했다. KPI는 하단 상태 바로, 운영 정보·시나리오는 펼칠 수 있는 영역으로 옮겼다. 기존 요소와 이벤트 처리는 유지한다.
- 탐색/속성 정돈: 압축된 구역 목록, 장비 이름·ID 검색, 선택 상태, 일관된 속성 패널, 별도 설정 메뉴와 렌더링 품질 선택기를 제공한다. 기존 모델 안내 그림은 펼치기 영역에 보존한다.
- 재질과 조명: 외부 HDR 파일 없이 만든 128px 스튜디오 반사 환경, 중성 바닥/구조물 색, 금속 프레임 마감, 닫힌 문의 불투명 표면 패턴 및 엠보싱, 구조 프레임의 둥근 모서리를 적용한다.
- 접촉부 음영: 현재 보이는 장비의 기존 바닥 면적에서만 음영을 만들고 하나의 InstancedMesh로 그린다. 새 장비나 연결을 생성하지 않는다. 피킹 대상에서도 제외한다.
- 선택·투시·도어·서버 레일·카세트·케이블 선택은 기존 구현을 재사용한다. 전체화면 버튼, 명시적 PC/모바일 화면 선택, 설정 메뉴도 유지한다.

## 품질과 자원 사용

| 모드 | 반사 환경 | 거칠기/도어 범프 맵 | 접촉 음영 | 기존 그림자 |
| --- | --- | --- | --- | --- |
| 저사양 | 끔, 단색 배경 | 끔 | 끔 | 끔 |
| 표준 | 128px 절차적 스튜디오 | 거칠기·범프 맵 끔, 문 색상 패턴 유지 | 1회 배치 | 1024 |
| 고품질 | 128px 절차적 스튜디오 | 브러시 거칠기·문 범프 맵 | 1회 배치 | 2048 |

sRGB, ACES 및 기존 DPR/LOD 정책을 유지한다. 새 SSAO/후처리 라이브러리나 외부 HDR·3D 모델 다운로드는 없다. 재질 맵, 환경 맵, 구조 모서리 geometry는 한 번 생성하여 공유하고 adapter 종료 시 해제한다. 새 독립 렌더링 루프를 생성하지 않는다. 모서리 geometry는 구조 표면 일부에만 사용하고 단위 크기의 선택 경계를 유지한다.

## 실제 화면 비교

생성 이미지가 아닌 실제 Chromium/WebGL 스크린샷이다. 1440×900, 고품질, 동일 시설·장비 선택·시뮬레이션 정지 조건이다. before-camera.json의 카메라 좌표를 after에도 그대로 적용한다. UI가 재배치되므로 캔버스 크기는 달라지며, 첫 진입의 자동 맞춤 시점과 비교 카메라를 구분한다.

| 화면 | 기존 main | 개선 코드 |
| --- | --- | --- |
| 전체 캠퍼스 | [before](before-campus.png) | [after](after-campus.png) |
| 서버실 | [before](before-hall.png) | [after](after-hall.png) |
| 닫힌 랙 정면 | [before](before-rack.png) | [after](after-rack.png) |
| 문 개방/내부 | [before](before-open.png) | [after](after-open.png) |
| 케이블 조작 | [before](before-cable.png) | [after](after-cable.png) |
| 장비 속성 | [before](before-selected.png) | [after](after-selected.png) |

before-layout.json/after-layout.json은 실제 viewport 크기, 품질, 선택 장비, 구역, DPR을 기록한다. `capture.cjs before`는 위 기준 SHA의 HTML/bootstrap/adapter/workspace 모듈을 읽기 전용으로 제공해 원래 UI를 재현한다. `capture.cjs after`는 현재 코드를 촬영한다. 공개 사이트 배포 결과로 표시하지 않는다.

## 실행

정적 HTML/CSS/ES module 구조이며 기존 GitHub Pages `/LSCNS/` 경로와 호환된다. 저장소 루트에서:

```bash
python3 -m http.server 8765 --bind 127.0.0.1
```

주소는 `http://127.0.0.1:8765/DCI/DataCenter/LS_Datacenter_3D.html?engine=three`이다. `?engine=legacy`는 기존 렌더러를 사용한다. 새 UI는 레거시에서도 사용할 수 있고 스튜디오 3D 재질은 Three.js 경로에 적용된다.

검증 의존성은 기존과 같은 `@playwright/test@1.56.1`, `three@0.186.1`이다. 선언/lockfile 및 BOM 의존성은 변경하지 않는다.

```bash
node --test DCI/DataCenter/ls3d-three/industrial-visuals.test.mjs DCI/DataCenter/ls3d-three/rack-xray.test.mjs
# Pages 하위 경로 검사는 8766 포트의 /LSCNS/ 경로도 필요하다.
# .github/workflows/ls3d-step19-quality.yml의 서버 설정과 준비 검사를 참고한다.
LS3D_CHROMIUM_EXECUTABLE=/usr/bin/chromium LS3D_NO_VIDEO=1 LS3D_VERIFIED_CDN=1 \
  npx playwright test --config DCI/DataCenter/ls3d-three/step19.playwright.config.mjs --timeout=120000
```

`LS3D_VERIFIED_CDN=1`은 관리형 브라우저에서 시스템 CA로 검증한 curl을 통해 공개 CDN 자산을 전달하는 기존 테스트 옵션이다. TLS 검증은 유지한다. 일반 CI에서는 생략한다. 선택적인 시스템 Chromium 경로와 video 옵션도 기존과 같다.

## 변경 파일

- 새 UI: `ls3d-three/industrial-studio.mjs`, `industrial-studio.css`.
- 새 표면/조명/음영: `ls3d-three/industrial-visuals.mjs`, 단위 검사 `industrial-visuals.test.mjs`.
- 통합: `workspace.mjs`, `scene-adapter.mjs`, `architecture.mjs`, `rack-components.mjs`, `bootstrap.mjs`, 진입 HTML의 스타일/캐시 버전 및 v4.8.9 표시.
- 브라우저: 기존 `workspace.playwright.spec.mjs`에 영역 크기·검색 검사를 보강하고 `studio.interaction.playwright.spec.mjs`에서 native fullscreen, 명시적 모바일 선택, 설정 및 범례 메뉴를 검증한다.
- CI: 기존 통합 QA에 새 단위·브라우저 검사를 포함하고 이 작업 브랜치에서도 실행한다. 실제 배포 job은 추가하지 않는다.

## 한계

기존 공간·장비 형상은 공유하며 표면/조명/일부 구조 모서리와 UI를 개선했다. 실측 CAD/BIM 정합성, 제조사 재질 규격, 실제 센서 수집 및 운영 제어는 검증 대상이 아니다. 실제 GPU·모바일 하드웨어의 FPS를 보장하지 않는다. 표준/고품질의 반사와 표면 효과는 렌더링 비용이 있으므로 저사양 모드와 측정 결과를 함께 제공한다. 원래 N-XI Preview는 별도 렌더러이며 이번 작업에서 교체하지 않는다.

## 기능 검증 결과

소스/단위 41개, 기존 브라우저 전체 21개가 통과했다. 새 스튜디오 상호작용 검사 1개를 추가해 고유 브라우저 시나리오는 22개다. 최종 품질 정책 적용 후 UI·품질·검색·투시·폴백·Pages·native fullscreen 7개를 재검사하여 모두 통과했고, 이후 셰이더 품질 전환 변경 후 관련 3개도 다시 통과했다. 접촉 음영 추가 후 7구역 전환 검사도 통과했다. 최초 22개를 한 번에 실행한 결과라고 표시하지 않는다.

집계는 [validation-summary.json](validation-summary.json)에 보관했다. 실제 작업 화면은 [1440×900](workspace-1440.png), [1920×1080](workspace-1920.png), [390×844](workspace-390.png) 및 [데스크톱에서 명시적 모바일 모드 선택](studio-mobile-choice.png)으로 확인할 수 있다.

## 표시 영역과 성능 실측

동일 1440×900 비교에서 캠퍼스 viewport는 865×440.7px에서 926×641.4px로 바뀌었다. 표시 면적은 약 **55.8% 증가**했다. 카메라 좌표와 시설 데이터는 동일하게 재생했다.

성능은 Linux x86_64, Xeon Platinum 8573C 5 vCPU, Chromium 156/ANGLE SwiftShader 소프트웨어 경로에서 측정했다. 전후 모두 캔버스 내용 크기 850×500, 동일 Data Hall·선택·시뮬레이션 정지 조건이다. 각 품질에서 3초 및 최소 6프레임 준비 후 실제 약 8초 경과시간과 프레임 수를 기록했다. 다른 테스트 브라우저와 동시에 실행하지 않았다.

| 상태 | 품질 | 실제 DPR | FPS | Calls | Triangles | Geometries | Textures | 오류 |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| before | low | 0.80 | 12.49 | 2492 | 64788 | 126 | 43 | 0 |
| before | medium | 0.77 | 1.20 | 21439 | 398482 | 947 | 204 | 0 |
| before | high | 1.00 | 1.14 | 21439 | 398482 | 947 | 204 | 0 |
| after | low | 0.80 | 5.00 | 2492 | 66068 | 129 | 43 | 0 |
| after | medium | 0.77 | 0.73 | 21441 | 426548 | 950 | 204 | 0 |
| after | high | 1.00 | 0.58 | 21441 | 426548 | 950 | 205 | 0 |

**소프트웨어 렌더링 FPS는 기존보다 낮게 측정됐다. 성능 향상이나 실기기 GPU 준비 완료로 표시하지 않는다.** 저사양에서도 실제 GPU 브라우저에서의 추가 측정이 필요하다. 표준/고품질 반사와 표면 효과의 비용을 확인한 뒤 사용 환경에 맞춰 선택해야 한다.

[before-metrics.json](before-metrics.json)과 [after-metrics.json](after-metrics.json)에 실제 프레임 수, 경과시간, canvas 크기, 사용자 에이전트, DPR, 품질 및 오류를 기록했다. G/T는 GPU 메모리 바이트가 아닌 객체 수다. Calls/triangles/memory는 구간 마지막 실제 프레임 값이며 FPS는 구간 평균이다.

`benchmark.cjs before|after`는 이 조건을 재현한다. 클라우드의 소프트웨어 결과를 실기기 GPU 성능으로 일반화하지 않는다. 기능 검증과 시각 개선은 완료했으며, 성능 검토를 포함해 승인 전 작업 브랜치로 제공한다. 병합/공개 사이트 배포는 수행하지 않았다.
