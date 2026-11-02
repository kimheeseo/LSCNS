# LSCNS | Data Center Digital Twin & Optical Research

데이터센터 3D 디지털 트윈, 초기 BOM/고객용 설계 도구, 광 네트워크 교육, 광통신·EDFA 데이터 기반 연구를 정리한 저장소입니다.

> **업데이트:** 2026-11-02 (KST) · 브랜치 `main`  
> 데이터센터 3D HTML의 표시 버전은 **v4.8.8**입니다. 최신 구현은 버전 문자열뿐 아니라 실제 소스와 Git 커밋을 기준으로 확인합니다. 3D 설비·링크·BOM은 학습 및 초기 검토용 개념 모델이며, 실측 CAD·확정 설계·검증된 견적을 대체하지 않습니다.

## 바로 실행

| 도구 | 실행 / 소스 | 설명 |
|---|---|---|
| **데이터센터 3D 투어 (Three.js)** | **[3D 디지털 트윈 실행](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=three&v=20261102-main-readme)** | 현행 Three.js / WebGL2 3D 투어, 랙 단독 보기·광부품·전력·냉각·네트워크 |
| 기존 렌더링 모드 | [Legacy 3D](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=legacy) | 기존 렌더러 경로; 새 3D 엔진과 구분 |
| **DataCenter BOM / 설계** | [DataCenter 메인](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/) | 전력·냉각·네트워크·케이블·제품 카탈로그 기반 계획 |
| 고객용 간편 설계 | [온프레미스 / 서버실 / AI DC](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/customer-designer.html) | 고객 환경별 핵심 입력 → 계획 수량과 구성 후보안 |
| 데이터센터 기초 OJT | [단계별 OJT](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/data-center-ojt.html) | 시설, MMR·Colocation, 랙 구성 학습과 3문항 복습 |
| **OFC 2026 ML Challenge** | **[EDFA Gain 예측 연구](26OFC_ML_challenge/README.md)** | 물리 기반·Transformer·FiLM·COSMOS 전이학습 비교 |

## 데이터센터 3D 투어

**실행 파일:** [`DCI/DataCenter/LS_Datacenter_3D.html`](DCI/DataCenter/LS_Datacenter_3D.html)  
**엔진:** Three.js (WebGL2, r186 계열) / 기존 Legacy 렌더러 선택 가능

### 주요 기능

- **시설 투어:** 데이터센터의 전력(발전기·UPS·배전), 냉각(칠러·CDU·펌프), Network/MMR·ODF, 서버·GPU·Storage 랙, 케이블 경로를 단계별로 확인합니다.
- **랙 단독 보기:** 전면/후면 시점, 도어 열기·닫기, 서버 레일·광 카세트 인출, 장비 및 구성품의 이름·역할 표시를 제공합니다.
- **광배선 부품:** 19인치 광 패치패널 내부, MPO/MTP 카세트, LC 어댑터 패널, 소형 ODF 단자함을 개별 관찰하고, 일부 포트·광경로를 강조할 수 있습니다.
- **광 네트워크 전체화면:** 일반 아키텍처 개념도와 디지털 트윈의 설계 토폴로지/링크 정보를 구분하여 표시합니다. 광 예산·링크 경로와 연계하되, 실제 시공 측량 결과로 오인하지 않도록 안내합니다.
- **설계/운영 시나리오:** 전력·냉각·네트워크 장애 및 기본 계산 가정을 학습용으로 확인합니다.

### 2026-11-02 기준 최신 변경

| 항목 | 반영 내용 |
|---|---|
| **랙 전면·후면 도어** | 닫힌 상태에서 내부 장비가 보이지 않도록 불투명 패널 적용. 문을 열면 내부 확인 가능 |
| **랙 자유 시점** | 랙 단독 보기에서 왼쪽 드래그를 **기본 회전**으로 변경. 좌우 360° 및 위·아래 관찰 범위 확대 |
| 화면 이동 / 확대 | 오른쪽 드래그 또는 Shift+왼쪽 드래그로 이동, 휠로 확대·축소, **랙 전체 맞춤**으로 시점 복귀 |
| **3D 화면 2×** | 화면 2× 버튼으로 3D 뷰포트의 **표시 면적을 약 2배**로 확장, 재클릭/닫기로 복귀 |
| 랙 부품 상세 | 광 패치패널·MPO/MTP·LC 어댑터·소형 ODF, GPU/스위치 부품 선택 및 내부 설명 고도화 |
| 광 네트워크 | **일반 구조**와 **디지털 트윈 설계 버전** 탭 분리 |
| 화면 정리 | 기존 2D 투어·운영/계산 Word 가이드북 등 이전 메뉴 링크를 3D 화면에서 정리 |

**랙 조작 방법:** 3D 투어에서 랙을 선택하고 **랙 단독 보기**에 들어갑니다. 왼쪽 드래그는 랙을 둘러보는 **회전**, 오른쪽/Shift+드래그는 **화면 이동**입니다. 상단의 **회전·이동 버튼**으로 기본 동작을 바꿀 수도 있습니다. 터치 기기에서는 한 손가락 드래그와 두 손가락 제스처를 사용합니다.

**주의:** 각 구성품의 형상, 장착 수량, 포트·케이블 경로와 카메라 확대 기능은 시각화·학습용입니다. 실제 제조사 CAD, 현장 전력·냉각 성능 검증, 통신 링크 실측, 정식 시공 도면과 다릅니다.

## 고객용 간편 설계 및 BOM

고객용 간편 설계는 상세 BOM을 작성하기 전에 구축 규모와 설계 방향을 검토하는 **초기 계획 화면**입니다. 결과의 **상세 엔지니어링 BOM 열기**를 이용해 기존 DataCenter 설계 도구로 이동할 수 있습니다.

| 고객 유형 | 중점 검토 | 주요 계획 출력 |
|---|---|---|
| 온프레미스 | 보안·망 분리, TCO, 증설성 | 랙, IT/시설 전력, UPS, GPU, 기본 이중경로 케이블 |
| 서버실 | 공간, UPS 백업시간, 냉각, 연속 운영 | 랙·공간 여유, UPS 에너지, GPU, 기본 이중경로 케이블 |
| AI 데이터센터 | GPU 랙, 고밀도 전력, 수랭, Fabric | GPU/서버/랙, 전력·냉각, Leaf/Spine, 고속 광·케이블 |

- **온프레미스·서버실:** 고객 입력 GPU 수를 사용합니다. 초기 케이블 계획은 `랙 수 × 2`의 이중 네트워크 경로를 가정하며, 매체·길이·커넥터·심수는 상세 검토가 필요합니다.
- **AI 데이터센터:** GPU·서버 조건을 `engineering-v82.js` 계산 엔진에 전달하고, Cable/AOC/DAC/Trunk/Patch 등 산출된 BOM 항목을 우선 사용합니다. 누락된 항목은 초기 링크 계획값으로 표시할 수 있습니다.
- **제품 DB 기반 1·2·3안:** ① 성능·확장 우선, ② 균형·표준화, ③ 조달 대안. 등록된 제품군 후보의 비교용이며, 확정 SKU·가격·납기·호환성 보증이 아닙니다.

단선결선도, 보호협조, UPS·냉각 실제 운전점, 광손실, 포트/FEC 규격, 시공 경로와 RFQ는 별도 엔지니어링 검토가 필요합니다.

## OFC 2026 ML Challenge · EDFA Gain Profile 예측

**[연구 모듈·재현 방법·검증 결과 보기](26OFC_ML_challenge/README.md)**

95채널 EDFA Gain Profile 예측을 대상으로 물리 기반 모델과 Spectral Attention Transformer를 비교하고, Residual-FiLM·SHB-aware loss·COSMOS 사전학습/미세조정을 실험합니다.

| 연구 모듈 | 내용 |
|---|---|
| [COSMOS Transfer Learning](26OFC_ML_challenge/COSMOS_Transfer_Learning/) | PI-EDFA / Transformer 사전학습 → 미세조정, Scratch 대조 |
| [FiLM Residual Transformer](26OFC_ML_challenge/FiLM_Residual_Transformer/) | 근사 물리 기준선 + FiLM 조건부 Attention + SHB-aware loss |
| [PI-EDFA / Spectral Transformer](26OFC_ML_challenge/PI_EDFA_Transformer/) | 채널별 입력 파워·mask·위치의 관계를 Attention으로 학습 |
| [Official Data Benchmark](26OFC_ML_challenge/results_2026-11-02/) | 공개 TestGround 사후 검증 및 오차 기록 |

동일 공개 TestGround 데이터 사후 비교에서 PI-EDFA **0.382980**, Spectral Transformer **0.333409**의 종합 점수를 기록했습니다(낮을수록 좋음). 이는 **비공개 대회 리더보드 점수가 아니며**, Aging·SHB에서는 Transformer가 PI-EDFA보다 악화되었습니다. 다른 COSMOS/FiLM 예비 실험은 조건이 달라 직접 순위 비교하지 않습니다. 상세 근거와 재현 조건은 각 연구 README를 확인하세요.

## 사용 설명서 · 변경 이력 · 검증

- [DataCenter 3D 사용 설명 및 날짜별 변경 이력](DCI/DataCenter/README.md) — 여러 시점의 변경 기록이 함께 수록되어 있으므로 **현재 도어·카메라 동작은 위 최신 변경 표와 실제 소스**를 우선합니다.
- [3D 개발 로드맵](DCI/DataCenter/LS3D_V4_20_TASK_ROADMAP.md) · [진행 현황](DCI/DataCenter/LS3D_PROGRESS_STATUS.md)
- [버전 변경 이력 TXT](DCI/DataCenter/LS3D_VERSION_UPGRADE_HISTORY.txt) · [Step 20 릴리스 보고서](DCI/DataCenter/architecture/STEP20_RELEASE_REPORT.md) — 초기 v4.8.0~4.8.1 단계의 역사적 문서입니다.
- [GitHub Actions 검증 내역](https://github.com/kimheeseo/LSCNS/actions) · [3D 관련 QA 워크플로](https://github.com/kimheeseo/LSCNS/actions/workflows/ls3d-step19-quality.yml)

**검증 범위:** 2026-10-16의 기존 자동화 이력에는 통합 Chromium QA **9/9 PASS**, 별도 Three.js QA **5/5 PASS**가 기록되어 있습니다. 이는 **당시 커밋**에 대한 결과이며, 2026-11-02 이후 수정된 최신 코드의 전체 실기기/모바일 FPS·WebGL2 품질을 보증하는 결과가 아닙니다. 최신 Actions 상태와 실제 화면 조작을 함께 확인해야 합니다.

## 관련 제품·기술 자료

- **[CPO Supply Chain](https://kimheeseo.github.io/cpo-supply-chain/)** — 광 인터커넥트/공급망 관련 외부 실행 사이트
- **[OCP Global Summit 2026 DataCenter 제품 DB 반영 기록](DCI/DataCenter/OCP_2026_PRODUCT_CATALOG_UPDATE_20261024.md)** — 전시·데모·프로토타입은 상용 구매 가능한 확정 SKU와 구분

---

문의나 추가 개발 시에는 [Issues](https://github.com/kimheeseo/LSCNS/issues) 및 해당 프로젝트 하위 README의 재현·검증 기록을 참고하세요.
