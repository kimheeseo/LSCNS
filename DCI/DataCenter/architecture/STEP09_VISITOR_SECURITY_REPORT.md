# STEP 09 / 20 — 방문객 출입 보안 시퀀스 구현 보고서

> 기준: 2026-10-15 KST · `kimheeseo/LSCNS` main  
> 상태: **코드 구현·보안 시퀀스 및 통합 Mock 검사 통과, 실브라우저/실제 보안 인증 검증 미실행**.

## 개발 범위
신규 [`visitor-security.mjs`](../ls3d-three/visitor-security.mjs)에서 기존 원본 자산 `entry-desk`, `xray`, `metal-detector`, `gate`의 좌표·규격을 읽어 단일 방문객의 순환형 입장 장면을 만들었다.

| 번호 | 방문 절차 | 구현 |
|---|---|---|
| 1 | 접수 데스크 접근 | 3D 방문객 이동 |
| 2 | 방문자 신원 접수 연출 | 정지·태블릿 자세 |
| 3 | X-ray 검사 대기 | 본체 우측 옆 통로로 이동 |
| 4 | 수하물 검색 | 여행가방 독립 Group이 컨베이어를 따라 이동; Scanner 램프 표시 |
| 5 | 금속 탐지기 통과 | 방문객 보행과 단계 전환 |
| 6 | 카드 리더 확인 | 카드 리더 램프 대기/승인 전환 |
| 7 | 보안 게이트 통과 | 진입 바리케이드 Pivot 회전 |
| 8 | 입장 완료 | 열린 게이트 유지 후 다음 사이클 초기화 |

기존 경비요원 2명(`guard-a`/ `guard-b`)은 경비 모형과 위치를 유지하며, X-ray 및 승인 단계에서 관절 점검 포즈를 보인다. 입장 방문객은 [`npc-characters.mjs`](../ls3d-three/npc-characters.mjs)에 추가한 별도 `visitor` 민간인 복장/머리 모델로 출력한다.

## 기술 및 검증
- 별도의 `requestAnimationFrame`을 만들지 않고 기존 `T.simTime` 주기에 연동한다.
- 기존 입구/캐리어/회사 서버실 운영 범위를 보존한다. **AI Campus / Campus 구역 + 4개 필수 자산이 모두 보일 때만** 동작한다. Server Room/Network 등 다른 모드에서는 보안 Group 자체를 숨긴다.
- `window.LS3D_THREE.adapter.security`에 단계명·게이트 회전 각도·가방 위치 등 진단 제공.
- 실제 74개 자산 데이터로 검사한 **8개 절차의 전 단계 전환, X-ray·접수데스크 footprint 침범 회피, 승인 전 닫힌 게이트, 승인 후 보행 중 열린 상태** 확인.
- 통합 Three API 모의 실행에서 Visitor 객체/보안 그룹/랙·NPC 분리와 Server Room 격리 확인.
- 방문자 개인 식별·실제 방사선 검사·보안 경보 인증 또는 의료/공공 시설용 안전 모델 아님. 시간·위치·승인 결과는 교육용 연출 값이다.
- Chrome/Android WebGL2 화면·Playwright 실행 로그는 **미확보**. [`step09.playwright.spec.mjs`](../ls3d-three/step09.playwright.spec.mjs)만 준비.

## 후속 과제
방문객 다중 큐, 접근 권한·ID 실데이터, 실제 카드 판독, 보안 대기 인원 물리 충돌/동선, 경비 인력의 상세 상호작용은 이번 시뮬레이션에 포함되지 않는다.
