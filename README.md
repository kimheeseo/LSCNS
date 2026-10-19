# LS Cable Digital Twin and Optical Tools

## 바로 실행

| 도구 | 링크 | 버전 및 상태 |
|---|---|---|
| **데이터센터 3D — Three.js 디지털 트윈** | [최신 3D 시뮬레이터 열기](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=three) | **v4.8.1** · WebGL2 / Three.js r186.1 · 안내 투어·가정값·광 네트워크 개선 |
| 기존 3D 렌더 | [Legacy 3D 열기](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=legacy) | 기존 v4.7.0 계산·운영 코어 유지 |
| 원본 2D Gold Pixel Tour | [2D 탐방](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_Campus.html) | 원본 보호 |
| 데이터센터 BOM 및 제품 툴 | [DataCenter 메인](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/) | 기존 BOM 도구 보존 |
| **고객용 간편 설계** | [온프레미스 · 서버실 · AI 데이터센터 간편 설계](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/customer-designer.html) | 고객 유형별 핵심 조건 → 랙·전력·냉각·GPU·케이블 계획 요약 |\n| **데이터센터 기초 OJT** | [시설 전체 · MMR/Colocation · 랙 구성 학습](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/data-center-ojt.html) | 초보자용 단계별 설명·선택형 도식·3문항 복습 |

## 고객용 간편 설계

고객이 복잡한 엔지니어링 입력 전에 구축 방향을 검토하도록 만든 2D 계획 화면입니다. 기존 상세 BOM 엔진을 대체하지 않으며, 결과 화면의 **상세 엔지니어링 BOM 열기**에서 DataCenter 메인 도구로 이동합니다.

| 고객 모드 | 강조하는 의사결정 | 주요 산출 |
|---|---|---|
| 온프레미스 | 보안·망 분리, TCO, 증설성 | 계획 랙, IT/시설 전력, UPS, GPU, 기본 이중경로 케이블 |
| 서버실 | 공간, UPS 백업시간, 냉각, 운영 연속성 | 설치 랙, 공간 여유, UPS 에너지, GPU, 기본 이중경로 케이블 |
| AI 데이터센터 | GPU 랙, 고밀도 전력, 수랭, Fabric | GPU·서버·랙, IT/시설 전력, UPS·발전기, Leaf/Spine, 고속 광·케이블 |

### GPU·케이블 표시 기준

- **온프레미스·서버실:** 고객이 입력한 GPU 수를 표시합니다. 케이블은 초기 계획값으로 `랙 수 × 2`를 사용하며, 랙별 이중 네트워크 경로를 의미합니다. 실제 광/동 케이블 종류·길이·커넥터·심수는 상세 BOM에서 산정합니다.
- **AI 데이터센터:** 목표 GPU 수 및 서버 프로파일을 기존 `engineering-v82.js` 계산 엔진에 전달합니다. 광/고속 케이블은 엔진 BOM의 Cable/AOC/DAC/Trunk/Patch 수량을 우선 사용하며, 항목이 아직 산출되지 않은 경우 GPU-Fabric 링크 기반의 계획 수량을 표시합니다.
- 모든 수량은 초기 사업성·공간·전력 검토용이며, 실제 단선결선도, 보호협조, 냉각/배관, 광 예산, 시공 경로, SKU 적합성 및 RFQ 검토를 대체하지 않습니다.

### 제품 DB 기반 1·2·3안

고객용 간편 설계는 산출된 GPU·서버·랙·스위치·고속 케이블·UPS 수량을 품목별 요구사항으로 변환하고, 기존 `product_catalog` DB의 등록 제품 후보를 조회해 세 가지 구성안을 표시합니다.

- **1안 · 성능·확장 우선:** 고속 AI/Fabric 확장성과 등록 제품 적합성을 우선 검토합니다.
- **2안 · 균형·표준화:** 다중 제조사 조달 가능성과 운영 표준화를 함께 검토합니다.
- **3안 · 조달 대안:** 동일 제품 분류의 다른 제조사 후보를 제시합니다.

세 안은 DB 기반 **제품 후보안**이며 가격 비교나 확정 SKU 제안이 아닙니다. 단가·납기·호환성·광손실·포트/FEC·전력 규격은 RFQ 및 상세 BOM 검토로 확정해야 합니다.

## 사용 설명서 · 개발 이력

- **[3D 시뮬레이터 사용 방법](DCI/DataCenter/README.md)** — 시작, PC/모바일 조작, 랙 서버 인출/문/카세트, 광·전력·냉각, MMR/Co-location, Troubleshooting.
- **[v4.8.1 버전업 상세 이력 TXT](DCI/DataCenter/LS3D_VERSION_UPGRADE_HISTORY.txt)** — 1~20단계의 변경 내용, 모듈 캐시 정책, 버전 승격 및 롤백 방법.
- **[20단계 최종 릴리스 보고서](DCI/DataCenter/architecture/STEP20_RELEASE_REPORT.md)** — 범위, 보호된 기존 파일, 테스트 근거, 실제 브라우저 확인 필요 항목.
- [20단계 개발 로드맵](DCI/DataCenter/LS3D_V4_20_TASK_ROADMAP.md) · [개발/검증 진행 현황](DCI/DataCenter/LS3D_PROGRESS_STATUS.md)
- [STEP 19 GitHub Actions 브라우저 QA](https://github.com/kimheeseo/LSCNS/actions/workflows/ls3d-step19-quality.yml)

### v4.8.1 개선 버전 상태

v4.8.1은 첫 진입 자동 가이드 투어, 전원→냉각→광 단선 추천 시나리오 재생, 편집 가능한 계산 가정/출처, 광 네트워크 빠른 보기 기능을 추가했습니다. 20단계 구현 및 문서 정리를 포함한 코드가 GitHub main에 반영되어 있습니다. Three.js/랙/MMR/고객별 배선/모바일 관련 JavaScript 모의 회귀 테스트 22/22는 19단계 재검증에서 PASS였습니다. **최신 빌드의 독립 Chromium/WebGL2 실화면·모바일·FPS 테스트와 GitHub Pages 배포 결과는 아직 확인되지 않았습니다.** 따라서 v4.8.1도 실기기에서 검증된 설계값이나 정식 엔지니어링 계산기로 취급하지 말고, GitHub Actions와 실제 접속 화면 검증을 함께 확인하세요.

### 이전 Phase 7 v4.7.0 기능

기존 시뮬레이터는 공유 설계 패킷 lsdc-twin-bom/1.0의 정보를 요약하고 구성 호환성 및 근거 부족을 표시하며, 회사 서버실을 모델링된 공간 범위에 한정하는 기능을 포함합니다. 이번 Three.js 업데이트는 기존 데이터·계산/BOM 호환성을 유지하면서 시각 탐방 레이어를 추가한 것입니다.
\n\n## 데이터센터 기초 OJT\n\n복잡한 BOM/설계 화면과 분리된 입문 교육 페이지입니다. 데이터센터 주요 구역과 전력·냉각 흐름을 익힌 뒤, 통신사 회선이 고객 랙까지 연결되는 과정(MPOE → MMR → MDA → Cage/Rack), Colocation의 역할 분담, 랙 내부 네트워크·전력 구성을 학습합니다. 구성요소를 클릭해 설명을 보고 마지막 3문항으로 핵심 개념을 확인할 수 있습니다. 도식은 일반적인 개념 예시이며 시설별 실제 설계는 다를 수 있습니다.\n