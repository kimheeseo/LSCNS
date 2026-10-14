# STEP 10 / 20 — 42U/48U 랙 구성품의 독립 3D 객체화 보고서

> 기준: 2026-10-15 KST · `kimheeseo/LSCNS` main  
> 상태: **모듈 구현·선택 API 및 통합 Mock 검사 통과, 실브라우저 GPU/클릭 회귀 검증 대기**.

## 핵심 결과
신규 [`rack-components.mjs`](../ls3d-three/rack-components.mjs)에 `createRackComponentKit(THREE,{cube})`를 구현하고 기존 Scene Adapter의 단순 랙 Mesh를 독립 부품 Group 구조로 교체했다.

- **42U/48U:** `asset.type/id` 기반으로 GPU/ToR 성격의 가상 랙을 48U, 그 외 일반 랙을 42U 시각 배치로 구분. 외함 높이는 기존 `a.h`를 사용하며 내부 단위 피치로 나눠 시각화한다. 실제 19인치/랙 제조사 도면의 밀리미터 치수에 대한 적합성은 검증하지 않았다.
- **독립 선택 부품:** 프레임, 마운팅 레일, 상부 광패치패널, ODF/광카세트, ToR 스위치, 다수의 서버/스토리지/네트워크 섀시, 전원 PDU A/B, 수직 케이블 관리자, 전·후면 도어.
- 각 부품 Group에 `userData.rackPart={id,assetId,kind,label,uStart,uHeight}`를 정의하고 내부 Mesh에 `partId`를 부여하였다. 실제 Three Raycaster 클릭 시 가장 가까운 부품 메타데이터를 찾아 아래 HUD에 랙 ID·부품명·유닛 위치를 표시한다.
- 공개 선택 API: `window.LS3D_THREE.adapter.selectRackPart(assetId,partId)`, `adapter.selectedPart`, `adapter.rackComponents`.
- 이전 STEP03 문 열기, 서버 레일 인출, 카세트 인출, 후면 보기 입력을 유지한다. **서버 개별 레일 조작과 케이블/포트 연결은 STEP11~14**로 분리하여 이번 단계에 추가하지 않았다.
- 기존 Rack LOD(상세/원거리 간략 모델), 2D/BOM 데이터 및 고유 설비 id 보존.

## 정적 및 API Mock 검증
- HTML의 실제 74개 자산과 함께 실행한 통합 Three API Mock **21/21 통과**(STEP08/09/10 연계, 결함 수정 전 수행된 통합 기준).
- 모의 Campus 장면에서 **23개 가시 랙**에 대한 별도 `rackPart` 구조 생성, **42U 랙 11개 / 48U 랙 12개**, **독립 선택 대상 401개** 관찰. 이는 테스트 시점의 샘플 상태이며 전체 운영 데이터/실제 제조사 BOM 수량으로 해석하면 안 된다.
- 선택 API를 통해 ToR 부품과 유닛 메타데이터 정상 조회, 선택 정보 HUD 생성을 위한 DOM 분기 코드 검사.
- Hall 변경, 회사 서버실 전환, 원래 STEP07 A* 직원·STEP06 NPC와 기능 연결 유지.
- **실제 WebGL2 부품 클릭, 전·후면 피킹, 랙 안쪽 Mesh 가림 현상, 모바일 탭, FPS·메모리**는 미검증이다.
- 브라우저 자동 검사: [`step10.playwright.spec.mjs`](../ls3d-three/step10.playwright.spec.mjs) 작성, 실행 로그/스크린샷은 아직 없음.

## 사용법
1. [Three.js 시험 페이지](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=three)를 새로고침.
2. Data Hall의 GPU 랙 또는 Network 스위치 랙을 선택한다.
3. 기존 상세조작으로 랙 문을 연 뒤 서버/상부 스위치/광패널/PDU를 클릭한다.
4. 화면 아래쪽 `선택 부품: <명칭> · <랙ID> · U <시작>-<끝>` HUD를 확인한다.
5. Console에서 `window.LS3D_THREE.adapter.rackComponents`로 건수 확인.

**실제 랙 수와 점유 U, 실물 서버 구동/전력/냉각 계수는 교육용 배치값이며 BOM 엔진 수치와 자동 결합하지 않았다.**
