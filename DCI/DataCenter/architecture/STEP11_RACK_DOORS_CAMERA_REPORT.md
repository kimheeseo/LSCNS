# STEP 11 / 20 — 전면·후면 랙 도어 및 독립 카메라 개발 보고서

> 2026-10-15 KST · GitHub `kimheeseo/LSCNS` / `main`
> **구현 완료 / Three.js API 모의 검증 통과 / 실제 Chromium·WebGL2 시각 회귀 미검증**

## 목적 및 구현 결과

STEP 10의 `rack-components.mjs`에는 이미 `front-door-hinge`, `rear-door-hinge` 두 실제 THREE.Group가 존재했지만, 후면 도어의 독립 회전 동작과 실제 후면 시점 조작이 부족했다.

- [`rack-door-controls.mjs`](../ls3d-three/rack-door-controls.mjs) 새로 작성. `createRackDoorControls()`가 각 랙 ID별 후면 개폐 target/current를 관리한다. `tick(id,moving,legacy,dt)`는 **기존 전면 도어 상태 `T.mechanical(a).door`를 그대로 반영**하면서 후면만 독립적으로 부드럽게 회전시킨다. 추가 RAF 없음.
- [`scene-adapter.mjs`](../ls3d-three/scene-adapter.mjs): `toggleRearRackDoor(assetId)`, `toggleRackFrontDoor(assetId)`, `focusRackSide(assetId,side)` 공개.
- 실 카메라 목표값은 기존 호스트 `T.cameraDesired`의 target/yaw/pitch/distance에 반영: **전면 = yaw π, 후면 = yaw 0**, 타깃은 해당 선택 랙 좌표. 새 후면 카메라 선택 중 랙 전체를 뒤집지 않아 위치 혼동을 줄인다. 기존 legacy `m.rear` 동작은 새 카메라가 없을 때 보존한다.
- 3D 화면 우하단 `#ls3dRackMotionControls` 패널의 **전면/후면 카메라, 전면/후면 도어** 버튼으로 조작할 수 있다.
- 랙이 시야/장면에서 삭제되면 해당 도어 상태를 정리한다.

## 검증

실제 GitHub에서 업데이트한 42U(스토리지)·48U(GPU) 모델을 가져와 가상 THREE.Group API로 시험하였다.
- 두 랙의 독립 전·후면 경첩 객체 존재.
- 전면 도어와 후면 도어가 별도로 회전하며 한쪽을 닫아도 다른 쪽이 열린 상태 유지.
- 두 도어 모두 닫힘 값으로 돌아옴.
- Side-camera API 호출 시 `cameraDesired.yaw` 및 target이 전면/후면으로 변경.
- STEP 11~13 **부품 모의 테스트 24/24 PASS**, Scene Adapter 결합 **17/17 PASS** 중 관련 검사 통과.

단, API 검증은 브라우저 렌더링 결과가 아니다. **실제 도어 메쉬의 물리 간섭, 카메라 클리핑, 후면 PDU 가시성, 모바일 터치 버튼 반응은 Chrome/Android 브라우저 확인 필요**. 도어/시점은 교육용 시뮬레이션이며 실제 설치 환경의 시야·안전 간격을 인증하지 않는다.

## 실기기 브라우저 확인

1. [Three 모드](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=three)에서 Hall → 랙 선택.
2. 우하단 **전면 카메라** 버튼을 누르면 전면을 바라보는지 확인.
3. **후면 카메라** 버튼으로 후면 PDU/후면 도어를 확인.
4. 후면 도어 열기 → 전면 도어 열기 → 전면만 닫기 → 후면만 닫기 테스트.
5. Rack 검사 모드/회사 서버실/기존 `?engine=legacy` 회귀, F12 Console 오류·클리핑 확인.

[실기기 Playwright 시험](../ls3d-three/step11to13.playwright.spec.mjs) 코드는 작성/구문 검사만 수행했으며 **실행 로그·스크린샷은 없음**.
