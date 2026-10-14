# STEP 08 / 20 — 작업자 설비 점검 상태머신 개발 보고서

> 기준: 2026-10-15 KST · `kimheeseo/LSCNS` main  
> 상태: **코드 구현·통합 Three API Mock 검사 통과, 실브라우저 Playwright/FPS 미실행**.

## 목표와 구현
기존 `T.npcPose()`/STEP07 A*의 보행 좌표를 그대로 사용하면서 작업자가 정지한 시간에 5단계 **이동(move) → 정렬(align) → 설비 점검(inspect) → 태블릿 기록(tablet) → 이탈(depart)** 을 실행한다. 독립 모듈 [`inspection-fsm.mjs`](../ls3d-three/inspection-fsm.mjs)의 `createInspectionFSM(T)` 구현이다.

- 직원별 상태와 완료 시각은 Map으로 관리한다. 움직이던 인원이 점검 위치에서 멈추면 정렬 0.50, 시각 점검 0.76, 태블릿 0.64 개념 시뮬레이션 시간 단계를 실행한다.
- 직원 역할/지역에 따라 가까운 설비를 타겟팅하고, 장애 시나리오 `power/cooling/network/fire` 등에서 관련 설비를 우선한다.
- 점검 위치에서는 해당 설비를 향해 상체/머리의 방향을 조정한다. `x,z`는 기존 A* 경로 값으로 고정해 충돌 검사 결과를 덮어쓰지 않는다.
- 태블릿은 `showTablet` 플래그가 참인 *tablet* 구간에만 표시한다. STEP06 관절 캐릭터 구현을 재사용한다.
- 현재 상태/대상/완료·경보 카운트는 `window.LS3D_THREE.adapter.inspections`에서 확인 가능하다.
- 점검 종료 중 장시간 정지하면 처음부터 반복되는 문제와 완료 횟수가 중복될 가능성을 코드 검사 중 발견하여 수정했다.

## 검증
- 실제 GitHub 소스의 JavaScript Three API Mock 통합 실행: STEP08/09/10의 핵심 연계 테스트 **21/21 통과** (수정 전 기준) 및 별도 FSM 단계 전환 시험 완료.
- 구역 7개, 기존 74개 설비 데이터, NPC 6명, A* Navigator와 랙/설비 모델 공존 확인.
- **보완 필요**: 설비 타겟은 기존 경로 정지점의 *가까운* 설비이다. 실제 설비 정면으로 접근하도록 A* 순찰 경유점 자체를 재설계한 것은 아니다. 전력/냉각 계측 데이터 기록이나 실제 태블릿 API 저장도 없다.
- 실 Chrome GPU/동작·Playwright 실행 결과는 **미확보**. [`step08.playwright.spec.mjs`](../ls3d-three/step08.playwright.spec.mjs)에 단계 전환과 장애 응답 회귀 코드만 제공.

## 범위 제한
직무별 점검은 개념 시각화이며 실제 안전점검 기록, 출입 승인, 정비 작업의 자동 수행 기능이 아니다. BOM/전력 수식, 2D Tour 및 기본 Legacy 렌더러는 수정하지 않았다.
