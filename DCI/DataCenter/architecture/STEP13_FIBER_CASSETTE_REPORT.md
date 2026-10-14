# STEP 13 / 20 — 광 카세트 인출·커버·MT 페룰/광 커넥터/광섬유 3D 모델

> 2026-10-15 KST · GitHub `kimheeseo/LSCNS` / `main`
> **모델·선택/인출 코드 구현 및 모의 검증 통과, 실제 브라우저 시각·수치 검증 대기**

## 범위와 기존 문제

기존 STEP 10 광 카세트에는 `fiber-cassette` 그룹과 단순한 드로어가 있었지만, 사용자가 인출해도 내부 어댑터/커넥터 및 광섬유 경로를 구별하기 어려웠다. 이번 단계는 이 그룹에 **실제 깊이 있는 독립 Mesh/TubeGeometry**를 추가했다.

- [`rack-optical-cassette.mjs`](../ls3d-three/rack-optical-cassette.mjs) 신규 모듈의 `decorate(asset,built)`가 기존 인출형 `cassetteDrawer` 자식으로 3D 트레이 바닥, 측면 rail, **8개 adapter port**, **8개 MT 페룰 형상**, **8개 커넥터 boot**를 생성.
- 각 포트에 `id=<rack ID>:fiber-port:<0..7>` 메타데이터를 부여하고 기존 `partMap`에 등록. Mesh 자체에 `assetId/partId` 태그를 부여해 클릭 판별 가능.
- 8개의 내부 광섬유 서비스 루프는 `THREE.CatmullRomCurve3` + `THREE.TubeGeometry`로 생성되므로 **납작한 그림이 아닌 3D 곡선 경로**다.
- `rack-optical:cover-hinge` 그룹의 독립 3D 커버는 인출 시 열림 각이 커진다.
- `update(id,moving,m)`은 기존 `T.mechanical(a).cassette` 상태에 따라 카세트 전체를 `-open*reach*.70` 방향으로 실제 이동시키고 커버를 `-1.18*open` 회전. 인출량 0에서는 모두 원위치.
- [`scene-adapter.mjs`](../ls3d-three/scene-adapter.mjs)의 raycast 클릭 처리에서 **하위 포트 자체 `partId`를 상위 Drawer 메타데이터보다 우선**해 선택한다. 우하단 패널에 **다음 광 커넥터** 버튼 및 **광 카세트 인출/복귀** 버튼 제공.
- 기존 광 카세트 정렬 및 전력·냉각/BOM 수치에는 손대지 않았다.

## 테스트 증거

GitHub 최신 42U Storage Rack/48U GPU Rack 소스를 읽어 THREE API 가짜 객체로 실제 build/update 수행.

- 각 랙별 8개의 adapter/페룰/커넥터/3D fiber = **두 랙 합계 16개 섬유/포트**.
- 카세트 인출 `position.z<0`, 덮개 `rotation.x<0`; 복귀 때 각각 0 확인.
- `partMap`에서 포트의 `kind:'fiber-port'` / `portIndex:0..7` 식별.
- 통합 Scene Adapter raycast에 **커넥터 페룰 Mesh 충돌 결과를 주입**해, 상위 카세트가 아니라 정확한 `fiber-port:0`가 선택됨을 확인.
- STEP 11~13 **부품 24/24**, **통합 Mock 17/17 PASS**.

**중요한 한계:** 포트·섬유 수는 교육용 예시이며 실제 MPO/MTP 광 채널 수 또는 MT 페룰 내 물리적 코어 배열을 의미하지 않는다. 광섬유 손실·굽힘반경, 접속 공차, 실제 MT 커넥터 기구 동작은 분석하지 않았다. 향후 14번 네트워크/케이블 연결 단계와 별개다.

## 실제 화면 점검

1. [Three 모드](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=three) → Hall → Rack 클릭.
2. **광 카세트 인출/복귀** 선택 후 Drawer가 앞쪽으로 움직이고 반투명 커버가 열리는지 관찰.
3. 가까이에서 8개 커넥터와 내부 3D 광섬유가 보이는지 확인.
4. 특정 커넥터를 클릭했을 때 부품 패널에 해당 포트가 뜨는지 확인.
5. 복귀 후 형태·위치·클리핑·다른 포트 간 겹침 확인.
6. 회사 서버실·모바일 터치·`?engine=legacy` 회귀 확인.

[STEP11~13 Playwright](../ls3d-three/step11to13.playwright.spec.mjs)는 저장했으나 실제 Chromium/WebGL2 성공 로그·스크린샷은 아직 없음.
