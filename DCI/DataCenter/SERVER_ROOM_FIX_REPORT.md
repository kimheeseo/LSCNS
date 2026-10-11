# Server Room 외부 객체 / MEP 수정 보고서

## 원인과 조치

- 서버실 렌더링 바닥은 Data Hall zone 중심에서 폭 26 m, 깊이 24 m이며 천장 구조는 약 7 m 높이입니다. 기존 Phase 6 경로는 hall 중심 기준 좌우 18 m/17 m로 생성되어 실내 폭을 초과했고, AIDC/캠퍼스 좌표의 광 흐름 경로도 공유 데이터에 남았습니다.
- `server-room-geometry.js`가 실제 hall 원점과 현재 renderer의 26×24 m 바닥 및 0~7 m 실내 높이로 `RoomBounds`를 만듭니다.
- `LS_Datacenter_3D_Extensions.js`의 server-room 렌더링 필터는 설비 외형 치수까지 bounds에 포함되는 자산만 표시합니다. BOM/설계 데이터에서는 asset을 삭제하지 않습니다.
- server-room에서는 실내 폭에 맞춘 22 m Busway만 생성합니다. 현재 server-room 냉각 preset은 공랭이며 CDU가 없어 DLC 공급/환수관을 만들지 않습니다. 향후 DLC와 CDU를 둘 다 구성했을 때만 실내 냉각선을 추가합니다.
- server-room optical 흐름 경로는 Data Hall bounds 안의 짧은 실내 경로로 바꾸고, 캠퍼스 링크를 그리지 않습니다. AIDC 및 On-Premise 경로는 유지합니다.

## 검증

- `tests/server_room_geometry.test.cjs`: 6 assertions 통과. 좌표 bounds, 기존 서버실 8개 설비, 외부 asset, 천장 초과, 내부/외부 route를 확인했습니다.
- `tests/phase8_browser_regression.cjs`에는 server-room 전환, 실내 자산 0건 외부 검사, MEP/optical route 검사와 스크린샷 캡처를 추가했습니다. 현재 실행 환경에는 Playwright 패키지는 있으나 Chromium 실행 파일이 없어 실제 스크린샷과 브라우저 결과는 미검증입니다.
