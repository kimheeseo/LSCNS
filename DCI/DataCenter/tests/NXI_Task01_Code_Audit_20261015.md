# N-XI 업그레이드 업무 1/20 — 현재 코드 및 기능 분석

- 완료일: 2026-10-15 (한국시간)
- 분석 기준 커밋: da61829d1b391f07e207d371da32c9f7a4df1d42
- 범위: 구조·기능·기존 테스트 증거 분석. 이번 업무는 런타임 변경 및 본 페이지 전환을 포함하지 않는다.
- 보호 대상: LS_Datacenter_Campus.html, DataCenter BOM 설계 앱 및 데이터. 이번 커밋은 이 보고서만 추가한다.

## 1. 현재 실행 상태

| 구분 | 파일 | 확인 결과 |
|---|---|---|
| 본 페이지 | LS_Datacenter_3D.html | v4.7.0 자체 WebGL. Three.js NXI 모듈을 로드하지 않음 |
| 신규 후보 | LS_Datacenter_3D_NXI_Preview.html | Three.js NXI 모듈과 NXI CSS를 로드함 |
| 렌더러 | LS_Datacenter_3D_NXI.js | 로컬 Three.js 0.160.0, 조명·그림자·반사 환경·선택 및 조작 |
| 모델 | assets/datacenter-3d/models.js | 독립 부품 기반 절차형 3D 모델 |
| 이동 | assets/datacenter-3d/navigation.js | FloorNav A* 및 StaffAgent 상태 전환 |
| 기존 효과 연결 | assets/datacenter-3d/response.js | 기존 장애 대응 효과의 Three.js 표현 연결 |

신규 후보를 본 페이지에 연결하는 작업은 업무 2에서 수행할 대상이다. 기존 코드를 모두 새로 작성할 필요는 없다.

## 2. 보존해야 할 기존 기능

- Campus / Utility Yard / Cooling Plant / Gray Zone / Data Hall / Network-MMR / NOC-Safety의 7개 구역.
- AI Data Center, 회사 서버실, On-Premises 운영 모드와 구역별 표시.
- 시뮬레이션 재생·일시정지·속도, 정상 및 장애 시나리오, 가정값 및 이벤트/추세 기능.
- 카메라 회전·이동·줌·탐방, 지붕·단면, 설비 선택·검색 및 랙 검사.
- 데이터 내보내기/가져오기, 교육 오버레이, 기존 BOM/호환성 엔진 연결.
- 기존 방문자 안내 투어 및 초미세 교육용 검사 기능. NXI 보안 시연과 기존 방문자 투어의 상호작용은 후속 검증 대상이다.

본 페이지는 UI·상태·카메라·기존 렌더러가 큰 HTML에 통합되어 있다. 신규 후보는 window.__LS3D_TEST__를 상태 연결 인터페이스로 사용하고 render3D 및 pickAt에서 NXI 렌더링/선택으로 위임한다. 독립 앱으로 전면 재작성하는 방식보다 이 연결을 재사용하면 기존 기능 보존 범위를 좁혀 검증할 수 있다.

## 3. 재사용할 구현과 업무별 검증 항목

| 업무 | 재사용할 코드 | 확인/보완 대상 |
|---|---|---|
| 2–3 | NXI 렌더러 및 CSS, 로컬 vendor/three.module.js | 모듈 연결, 실패 시 표시, 조명·그림자·가독성 |
| 4–6 | buildingModel / equipmentModel | 7개 구역 형상, 내부/외부 구분, 치수·배치 |
| 7–9 | personModel / animatePerson / FloorNav / StaffAgent | 보행, 점검, 회전, 장애물·사람 간 충돌, 고장 출동 |
| 10–11 | X-ray·게이트 모델, startVisitor | 가방 적재·컨베이어·검색·승인·문 개방 순서 |
| 12–16 | rackModel / serviceControls / 기계 상태 | 개별 도어·서버·카세트 실제 변환, ToR/NIC 포트 및 후면 배선 |
| 17–18 | transport 모델, Customer cage 모델 | MMR 식별, 고객 A/B/C 분리·문·접속 |
| 19 | buildRoutes / inspectCable / 카메라·레이어 UI | 계통 강조, 메타데이터, 모바일 및 탐방 통합 |
| 20 | 기존 NXI·호환성·서버실 테스트 | 최종 버전 회귀검증, 스크린샷, 배포 및 실제 실행 확인 |

재사용 가능은 해당 업무의 사용자 확인까지 완료되었다는 의미가 아니다. 각 업무에서 실제 동작 및 화면을 확인하고 별도로 완료 보고한다.

## 4. 기존 테스트 증거

기준 커밋의 GitHub Actions 상태를 API로 확인했다.

- NXI browser verification: completed / success
  https://github.com/kimheeseo/LSCNS/actions/runs/38589485737
- Pages build and deployment: completed / success
  https://github.com/kimheeseo/LSCNS/actions/runs/38589484041

검증 워크플로는 CI 체크아웃에서 Preview.html을 LS_Datacenter_3D.html로 복사한 뒤 검사한다. 따라서 CI 성공은 저장소 본 페이지가 이미 NXI로 전환되었다는 증거가 아니다.

워크플로 포함 검사:
- NXI 모델·A* 단위 검사 및 실제 캠퍼스 배치 직원 이동 검사.
- 기존 Phase7 계산·호환성 및 server-room geometry 검사.
- 고정 Playwright 1.56.1 Chromium으로 데스크톱/모바일 렌더링·조작 검사.
- 스크린샷 및 JSON 보고서 아티팩트 업로드.

브라우저 테스트 소스에서 7개 구역 렌더링, 실제 도어 회전·후면 도어·서버/카세트 조작, MMR/고객 공간, 컨베이어·보안 절차, 직원 행동·고장 출동, 서버실 외부 표시 차단, 모바일 및 오류 수집을 확인했다. 이번 업무에서 브라우저 테스트를 새로 실행하거나 기존 아티팩트의 스크린샷을 육안 검토하지는 않았다. 정확한 케이스 수·성능 수치를 임의로 보고하지 않는다.

## 5. 후속 보완 및 제한

1. 본 페이지 연결은 미적용: 업무 2에서 신규 렌더러 연결과 기본 동작 확인.
2. 기존 기능 호환성은 CI에 포함된 범위를 넘어 사용자 흐름별 확인이 필요.
3. 42U/48U 및 장비 크기는 명시된 모델링 가정이며 제조사 CAD/현장 실측이 아니다.
4. 층 선택은 단일 운용층의 IT/상부 서비스 표시 분리이며 실제 다층 BIM이 아니다.
5. 통신 속도·길이·커넥터는 개념/RFQ 정보. 검증되지 않은 제품 사양으로 표시하면 안 된다.
6. 하드웨어별 FPS와 실제 전력·냉각·통신 성능은 이번 구조 분석으로 검증되지 않는다.
7. 포트 검사, 케이블 곡률, 모바일 카메라 프레이밍 및 두 방문자 모드 병행은 후속 업무에서 실제 화면 확인.
8. GLB 없이 절차형 Mesh를 재사용한다. 추가 자산은 필요한 형상 개선에 한정한다.

## 6. 업무 1 완료 기준

- 본 파일과 프리뷰의 렌더링 차이 확인: 완료.
- 기존 기능과 상태/입력 연결 구조 분석: 완료.
- 모델·이동·보안·랙·연결 코드의 재사용 위치 정리: 완료.
- 기존 자동검증 구성 및 기준 커밋 성공 상태 확인: 완료.
- 재사용 범위와 미확인 항목 구분: 완료.

업무 1/20 완성. 사용자 확인 후 업무 2로 진행한다.
