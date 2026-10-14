# STEP 12 / 20 — 서버 레일 인출 및 NIC/포트 독립 3D 모델 보고서

> 2026-10-15 KST · GitHub `kimheeseo/LSCNS` / `main`
> **코드 구현 및 모의 API 검증 통과 / 실브라우저·GPU 픽셀 검증 대기**

## 문제 및 해결

기존 STEP 10의 `server-tray` 공통 THREE.Group에 다수 서버가 자식으로 들어 있었기 때문에 `T.mechanical(a).slide`을 적용하면 **여러 서버가 함께 당겨지는 현상**이 발생할 수 있었다.

- [`rack-components.mjs`](../ls3d-three/rack-components.mjs) 빌드 반환값에 `serverGroups`, `slideGroup` 등 **개별 서버 참조**를 추가한다. 이전 `moving.hinge/slide/cassette/reach` API 보존.
- [`rack-server-rails.mjs`](../ls3d-three/rack-server-rails.mjs) 새로 구현. 서버별 고정 guide rail, 이동 carriage, 3D NIC 카드, 전면 슬롯, 독립적인 4개 NIC cage/래치 포트와 LED를 생성.
- `rackRails.update(id,moving,m)`: 공통 parent `server-tray.position.z=0`으로 고정하고 **선택된 서버 group에만** `-m.slide*moving.reach`를 적용한다. 나머지 서버의 group.position.z는 0. 이를 통해 원래 슬라이더 UI의 `T.setMechanical(a,'slide')` 상태를 재사용하면서 개별 인출을 실현한다.
- 클릭한 서버의 기존 STEP 10 `userData.rackPart` ID를 통해 `setRackServerSelection(id,partId)` 또는 `cycleRackServer(id,±1)`으로 변경 가능.
- 우하단 랙 패널에 **이전/다음 서버 / 선택 서버 인출·복귀** 버튼을 추가. 랙을 다시 선택하면 해당 서버를 구분할 수 있다.

## 모의 실행 검증

- 48U GPU 랙 + 42U 스토리지 랙 2종에서 각각 **7개 독립 서버 그룹**, 합계 14개 3D NIC 카드·56개 개념 NIC 포트 생성.
- 첫 번째 서버 인출 시 두 번째 서버 위치 0 유지; 선택을 두 번째 서버로 바꾸면 첫 번째 0 복귀, 두 번째 서버만 인출되는 실제 THREE.Group 위치 수치 검사 PASS.
- Scene Adapter 결합 시 선택된 서버만 움직이고 다른 서버가 움직이지 않는지 검사 PASS.
- STEP 11~13 **부품 24/24**, **통합 Mock 17/17 PASS**. GPU 전송, 물리 레일/충돌 제한과 온도·성능은 실측하지 않음.

**주의:** 이 NIC 포트는 실제 제품의 정확한 QSFP-DD/OSFP/PCIe 물리 핀·전력·광 특성, 자사 BOM 또는 제조사 CAD를 재현한 것이 아닌 **개념적인 외형 모델**이다. 14번 ToR/NIC 링크 경로 전체 정밀 모델은 미착수.

## 브라우저 체크

1. Hall에서 GPU 랙 선택 → **다음 서버** 버튼 클릭.
2. **선택 서버 인출/복귀** 클릭 시 **해당 서버만** 밖으로 나오는지 확인.
3. 후면 카메라로 전환하고 독립 NIC 카드·4개 포트·레일 형태 확인.
4. 같은 방법으로 회사 서버실 Storage Rack에서도 인출을 재현.
5. 기기별 클릭/선택 UX, F12 오류, 프레임 안정성 확인.

자동화 초안: [STEP11~13 Playwright](../ls3d-three/step11to13.playwright.spec.mjs) (**작성됨·실행 미확인**).
