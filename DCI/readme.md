# AI Data Center BOM Engine v4.7.1
AI 데이터센터 물리 인프라/BOM 설계 프로토타입의 **개발 이력 및 검증 개요**를 정리한 문서입니다.

> **Source policy:** 실행 가능한 HTML/JavaScript 소스는 이 공개 저장소에 배포하지 않습니다.  
> 실제 실행본과 제품 DB는 비공개 저장소 또는 제한된 파일 저장소에서 관리합니다.

## 현재 개발 버전

**v4.7.1 Stable Excel / Multilingual / Korea Channels / Juniper**

설계 흐름:

**Requirements → Compute/GPU → Network Fabric → Rack → Optical Connectivity → Power/Cooling → Facility/Storage → Product Mapping → Estimated BOM**

## 주요 기능

- GPU / AI system profile 기반 system 수량 및 rack 산정
- Leaf / Spine / Core / Rail topology physical solver
- 실제 사용 포트 기준 Pod egress oversubscription 계산
- Rack RU / 전력 / 냉각 / airflow 제한 기반 iterative placement
- DAC / AOC / DR4 / FR4 / MPO / LC media selection
- 8F / 16F / 24F / 32F / 48F / 72F / 96F / 144F structured trunk sizing
- Patch panel / housing / connector / cable manager BOM
- A/B PDU 및 power-cord sizing
- In-band / OOB auxiliary network
- Facility 옵션: UPS / Generator / Transformer
- Storage 옵션: Storage system / Storage fabric
- 제품 DB 기반 Manufacturer / Product / Model / Qty / Evidence 예상 영수증
- Performance-first / Balanced / Cost-sensitive 조달 정책
- 설계 계산 옆 compact action bar: **Export Excel / Email Excel**
- Excel 출력 4개 시트: **요구조건 / 설계 결과 / Generic BOM / 제품매칭 예상 영수증**
- 제품매칭 예상 영수증 Alternative에 확인 가능한 공식 datasheet/spec URL 표시, 확인 어려운 경우 "-"
- 한국 구매/기술 문의 채널 패널: Corning / Sumitomo Electric / Supermicro / Juniper / Cisco 등
- Juniper QFX5240-64OD / QFX5240-64QD verified switch candidate 추가
- 선택 UI 언어에 따른 설계 예시 및 이메일 제목/본문 번역
- Conceptual BIM-ready floor plan / rack schedule
- CSV BOM / Validation CSV / Design JSON export
- 한국어 / 日本語 / English / 中文 / Deutsch 내장 UI 번역
- GPU 규모별 NVIDIA DGX SuperPOD reference-size class 표시
- Optical connectivity interface 구체화: Port / Optic / Connector / Fiber / Fibers-per-link / Trunk
- LS Cable & System / Hengtong / YOFC / Lightera / Sumitomo Electric / Corning / ZTT 다심 광케이블 후보
- Data Hall용 pre-terminated trunk와 Backbone용 high-count cable 분리
- 총 Fiber 수와 equivalent FP(Fiber Pair) 동시 표시
- 모바일/스마트폰 화면용 responsive layout 및 긴 문자열 줄바꿈
- Cisco Nexus N9364E-SG2-Q/O 및 N9364E-SP2R-Q/O switch candidate catalog
- MPO product/interface 세부화: Auto / verified MPO-8 / verified MPO-12 / verified MPO-16
- MPO-8 / 12 / 16은 제조사 공식 제품 또는 공식 제품군이 확인된 경우에만 추천
- 설계 계산 아래 **설계 예시** 셀 추가: Performance first / Balanced / Cost-sensitive 비교
- MPO ferrule / installed fiber / typical application / breakout / direct-channel utilization 표시

## Version history

### v2.x — 초기 통합 엔진
- GPU/Server → Fabric → Rack → Optical → Power → Cooling → BOM 흐름 통합
- Rack 사이 실제 연결 도식 및 BIM-ready summary 추가
- Power / thermal sizing 및 runtime 오류 수정
- Reference-blind validation 구조 도입

### v3.1–v3.2 — Solver hardening
- Fabric policy와 validation fixture 분리
- Server-group 기반 Pod 구성
- Pod egress oversubscription 명시
- Infeasible 설계 fail-closed 처리
- Compute-side panel/housing BOM 보완
- Core/Aux rack을 BIM rack count에 포함
- Single topology phantom Spine 제거
- Auxiliary switch design-power 기준 정리
- Fixed-point 안정성 검증 강화

### v3.3 — Physical connectivity visualization
- Rack / switch / panel / connector / cable 상세 SVG 구성
- Server → Leaf → Spine 연결 흐름 및 switch-face mapping 시각화

### v3.4 — Fiber-count policy
- Server↔Leaf / Leaf↔Spine에 8F~144F 심수 정책 추가
- 광심 수에 따른 trunk cable 수 자동 계산
- Optical loss budget / attenuation / connector loss / margin 입력 추가

### v3.5 — Dynamic physical connectivity
- Physical Connectivity 상세도를 입력값 기반으로 동적 생성
- P2P / Structured, fiber count, 거리, optics, switch model, panel/trunk 수량 연동
- Spine이 없는 설계에서 관련 구성 자동 제외

### v3.6 — Product DB & validation hardening
- 부품별 업체 리스트 및 curated product DB 추가
- Manufacturer / Product / verified Model·SKU / Qty / Evidence 기반 예상 영수증
- Pod oversubscription을 실제 사용 포트 기준으로 수정
- Spare 정책 양끝 대칭화 및 rack/boundary 단위 trunk 산정
- REVIEW 진입 시 이전 결과 전면 무효화
- DR4 / FR4 application별 loss budget 및 MPO/LC connector-loss 분리
- Placeholder/internal label의 SKU 출력 방지
- Auxiliary network profile 및 mixed-speed physical-cage 검사 개선

### v3.7 — Facility / Storage / tiered procurement
- UPS / Generator / Transformer 초기 sizing
- Storage system / Storage fabric 요구조건 및 BOM 확장
- Performance-first / Balanced / Cost-sensitive 제품 후보 정책 추가
- 실제 가격은 RFQ 단계에서 확인하도록 분리

### v3.8 — Multilingual / neutral release
- 조직 특정 표현 제거 및 중립형 UI 정리
- 한국어 / 日本語 / English / 中文 / Deutsch 5개 언어 지원
- 제품명·SKU·DR4/FR4/MPO/OSFP 등 기술 식별자는 원문 유지


### v4.0–v4.1 — Private deployment / product-aware advisor
- 실행 가능한 설계 엔진을 Private Git repository + Render 구조로 분리
- 공개 GitHub에는 개발 이력/검증 문서만 유지
- Contact 안내 문구 및 Render 배포용 patch 구조 추가
- GPU 수와 시스템 입력을 기반으로 NVIDIA DGX SuperPOD reference-size class 표시
- Optical connectivity advisor 추가
- 400G DR4 / FR4 / SR8, NVIDIA NDR 400G MMF, 800G parallel, short-reach DAC/AEC/AOC application 구분
- Port → Optic → Connector → Fiber → Trunk chain 표시
- Required fiber 및 equivalent FP(Fiber Pair) 계산
- Data Hall / Backbone trunk scope 구분

### v4.2.0 — Optical cable catalog / mobile layout
- 광케이블 vendor catalog 확대
  - LS Cable & System: Micro Array / Lock'n Roll™ Ribbon
  - Hengtong: High-density MPO / pre-terminated data-center cable
  - YOFC: 12–144F MPO/MTP pre-terminated trunk
  - Lightera: DuctSaver® / AccuTube®+ / R-Pack rollable-ribbon 계열
  - Sumitomo Electric: FREEFORM RIBBON™ UHFC / 3,456F pre-connectorized MPO / 6,912F high-density cable
  - Corning: EDGE8® MTP® trunk / RocketRibbon® high-fiber-count cable
  - ZTT: 17,280F ultra-high-density flexible-ribbon cable
- 각 vendor에 대표 제품군과 공식 Product / Homepage 링크 구분
- exact SKU 또는 fiber-count가 공개자료로 확인되지 않는 경우 임의 생성하지 않고 RFQ로 표시
- 스마트폰 화면에서 긴 제품명·URL·interface 문자열이 화면 밖으로 벗어나지 않도록 responsive wrapping 보완
- 좁은 화면에서 주요 card / KPI / form을 1열로 자동 배치
- VERSION / CHANGELOG 기반 버전 관리 정책 도입
- 앞으로 사용자에게 보이는 기능, 제품 DB, UI 변경 시 버전을 갱신하고 이 README의 Version history에도 함께 기록


### v4.3.0 — Cisco switch catalog / MPO Base architecture
- Cisco 데이터센터/AI fabric switch 후보 추가
  - Nexus N9364E-SG2-Q: 64 × 800G QSFP-DD, 2RU, 995W typical / 2,270W max
  - Nexus N9364E-SG2-O: 64 × 800G OSFP, 2RU
  - Nexus N9364E-SP2R-Q/O: 64 × 800G, 3RU, 16GB HBM deep-buffer spine/DCI 후보
- Cisco N9364E-SG2는 2×400G / 8×100G breakout을 지원하고, SP2R 계열은 2×400G / 4×200G / 8×100G 등 더 다양한 breakout mode를 지원
- Cisco 항목은 공식 datasheet 기반 verified candidate로 표시하며, 기존 타사 switch profile에 임의 매핑하지 않음
- MPO architecture selector 추가 *(v4.4.0에서 verified product 방식으로 대체)*
  - **Base-8**: 일반적으로 MPO-12 ferrule에서 8 active fibers를 사용하는 4-lane parallel 구조. SR4 / PSM4 / DR4 및 다수 400G parallel/breakout 구성에 활용
  - **Base-12**: 12-fiber trunk 구조. 8F parallel optic을 직접 연결하면 4심이 미사용될 수 있어 cassette/harness packing 검토 필요
  - **Base-16**: MPO-16 / 16 active fibers. 400GBASE-SR8, SR8/DR8(PSM8), 일부 800G parallel optic 등에 적용
- MPO-16 ↔ 2 × MPO-8 Y-harness 등 migration/breakout 경로 표시
- MPO base 선택은 transceiver의 실제 active fiber count를 임의 변경하지 않으며, polarity / cassette / harness / physical-link graph를 별도 검토
- 실행 화면 제목을 **AI 데이터센터 BOM 설계 엔진 v4.3.0**으로 갱신
- 앞으로 버전 증가 시 실행 화면 제목과 본 README 제목/현재 개발 버전/Version history를 함께 갱신

### v4.4.0 — Verified MPO products / procurement design examples
- MPO connector 추천을 generic Base architecture 중심에서 **실제 제품 확인 중심**으로 변경
  - MPO-8 / 8F MTP: Corning EDGE8® 8F MTP® trunk 제품 확인
  - MPO-12 / 12F MTP: Corning EDGE™ 12F MTP® trunk 제품 확인
  - MPO-16 / 16F: SENKO MPO-16 및 US Conec MTP®-16 제품군 확인
- 800G라는 속도 정보만으로 MPO-8/12/16을 임의 선택하지 않으며, exact transceiver/interface가 없으면 MPO 추천을 보류
- Optical connectivity 화면에 verified product와 official product link를 함께 표시
- 설계 계산 버튼 아래 **설계 예시** 셀 추가
  - Performance first: 800G-ready, 1:1 우선, 높은 spare/headroom, 고밀도 cooling 우선
  - Balanced: 400G server-facing + 필요 시 800G uplink, 2:1 초기 검토, 중간 spare
  - Cost-sensitive: 400G 중심, workload 허용 시 높은 oversubscription 검토, short-reach DAC/AEC/AOC 우선
- 설계 예시는 확정 BOM이 아니라 현재 GPU/system/topology 입력을 기준으로 한 advisory configuration이며, 실제 수량은 설계 계산을 다시 수행하여 확정
- Private build의 `VERSION`을 단일 source of truth로 사용하도록 version patch 개선
  - Webpage title
  - API/package metadata
  - Optical advisor version
  - Design example version
  를 build 시 같은 버전으로 동기화

### v4.5.0 — Multilingual design examples / Excel export & email
- **설계 예시** 영역을 한국어 / English / 日本語 / 中文 / Deutsch UI 선택에 따라 동적으로 번역
- 설계 계산 버튼 옆에 작은 action bar 추가
  - **Export Excel**: 현재 입력/요약/설계 데이터를 실제 `.xlsx` 파일로 생성
  - **Email Excel**: 수신자 이메일을 입력하면 동일한 `.xlsx` 설계안을 첨부하여 발송
- Excel workbook 구성
  - Summary
  - Inputs
  - Design Data
- 이메일 발신자는 `harrykim9463@gmail.com`으로 고정
- 이메일 제목과 안내 문구는 현재 선택한 UI 언어를 따름
  - KO: 이 설계안이 당신의 업무에 도움이 되면 좋겠습니다.
  - EN: I hope this design proposal helps you with your work.
  - JA: この設計案があなたの業務に役立てば幸いです。
  - ZH: 希望这份设计方案能对您的工作有所帮助。
  - DE: Ich hoffe, dass dieser Entwurf Sie bei Ihrer Arbeit unterstützt.
- 메일 발송 endpoint는 same-origin으로 제한하고 client별 10분당 3회 rate limit 적용
- Gmail App Password는 소스/GitHub에 저장하지 않고 Render Secret Environment Variable로만 설정
- `VERSION` 단일 기준으로 webpage/API/package/Optical Advisor/Design Examples/Action Menu 버전 동기화

### v4.6.0 — Four-sheet Excel / receipt datasheets / Korea channels / Juniper
- Excel 설계안 구성을 다음 4개 시트로 재구성
  1. 요구조건
  2. 설계 결과
  3. Generic BOM
  4. 제품매칭 예상 영수증
- Generic BOM / 제품매칭 예상 영수증은 화면에 렌더링된 현재 계산 결과 table을 기준으로 Excel에 반영
- 제품매칭 예상 영수증의 Alternative 항목에 exact product/family가 확인되는 경우 공식 datasheet/spec 링크를 추가하고, 확인이 어려우면 `-` 표시
- **설계 예시** 언어 선택 추적 강화: 한국어 / English / 日本語 / 中文 / Deutsch 버튼 클릭 상태를 공통 UI language state로 전달
- 설계 예시 / Export Excel / Email Excel 아래에 **한국 구매 / 기술 문의 채널** 추가
  - Corning: 남해이엔지 취급 채널 + 한국코닝 공식 법인
  - Sumitomo Electric: Sumitomo Electric (Korea) Electronics
  - Supermicro: Korea Office
  - Juniper: 공식 총판 인성디지탈 / 시엔스 / 투케이엠시스템즈
  - Cisco: Cisco Systems Korea
- Juniper 제품 기업/스위치 candidate 추가
  - QFX5240-64OD: 64 × 800GbE OSFP, 2RU
  - QFX5240-64QD: 64 × 800GbE QSFP-DD, 2RU
  - 공식 Juniper QFX5240 datasheet 링크 포함
- Private `VERSION` / 실행 화면 / README / CHANGELOG를 v4.6.0으로 동기화

### v4.7.0 — Stable reimplementation
- v4.4.0 안정 배포본을 기준으로 이전 요청 기능을 다시 구현
- Backend는 기존 server.js에 문자열 치환을 누적하지 않고, 검증된 전체 server 파일을 build 시 생성하도록 변경
- Excel 설계안 4개 시트
  1. 요구조건
  2. 설계 결과
  3. Generic BOM
  4. 제품매칭 예상 영수증
- 설계 예시 / Export Excel / Email Excel을 한국어 / English / 日本語 / 中文 / Deutsch 선택 상태와 연동
- Email Excel:
  - sender: `harrykim9463@gmail.com`
  - 현재 선택 언어의 제목/본문 사용
  - Excel 첨부
  - Gmail App Password는 Render Secret Environment Variable에서만 사용
- 제품매칭 예상 영수증 Alternative:
  - 정확한 제품/제품군을 확인할 수 있으면 공식 datasheet/spec 링크 표시
  - 확인이 어려우면 `-`
- 한국 구매/기술 문의 채널 패널
  - Corning: 공식 Optical Communications distributor 및 한국코닝
  - Sumitomo Electric: Sumitomo Electric (Korea) Electronics
  - Supermicro: Korea Office
  - Juniper: 인성디지탈 / 시엔스 / 투케이엠시스템즈 공식 총판
  - Cisco: Cisco Systems Korea
- Juniper QFX5240-64OD / QFX5240-64QD를 switch candidate 및 receipt alternative에 추가
- 복구용 안정 branch `dc-bom-v44-stable`을 별도 보존

### v4.7.1 — Email removed for deployment stability
- 홈페이지 안정성을 위해 Email Excel 기능을 제거
- `nodemailer`, Gmail runtime dependency, `/api/email-design` endpoint 제거
- Render의 `MAIL_USER` / `MAIL_APP_PASSWORD`가 없어도 실행 가능
- **Export Excel**은 유지하며 4개 시트 출력
  1. 요구조건
  2. 설계 결과
  3. Generic BOM
  4. 제품매칭 예상 영수증
- 설계 예시 다국어, datasheet/spec 링크, 한국 구매/기술 문의 채널, Juniper QFX5240 후보는 유지
- Private deploy branch: `dc-bom-v4-deploy`
- Recovery branch: `dc-bom-v44-stable`

## 30-Case Independent Validation Ledger

30-Case Ledger는 개별 고객 설계 결과가 아니라 **엔진 자체의 검증 이력 관리표**입니다.

검증 원칙:

1. 엔진에는 요구조건과 제품 catalog / physical constraints만 입력
2. Reference topology/BOM 값은 계산 입력에서 분리
3. 계산 이후 공개 Reference와 비교
4. 파생값은 독립 validation field로 중복 점수화하지 않음
5. 공개·독립 evidence가 부족한 Case는 높은 validation grade를 부여하지 않음

Golden fabric case에서는 Leaf / Spine 등 독립 비교 가능한 항목을 중심으로 오차를 계산합니다.

## Product DB policy

제품 DB는 전 세계 모든 SKU를 완전 수집한 데이터베이스가 아니라, 데이터센터 설계에 필요한 주요 제품/제품군을 공식 제조사 자료 위주로 정리한 **curated engineering DB**입니다.

- 확인된 exact SKU/model만 SKU/model 필드에 사용
- 제품군만 확인 가능한 경우 family/RFQ로 표시
- 지역별 공급성, firmware compatibility, support contract는 RFQ 단계에서 재확인
- 조달 정책은 실제 최저가 순위가 아니라 요구조건을 만족하는 제품 후보의 우선 검토 전략


## Update / versioning policy

앞으로 데이터센터 BOM 설계툴에 사용자에게 보이는 기능, 제품 catalog, validation, UI/UX, 계산 정책이 새롭게 반영될 때마다 다음 항목을 함께 갱신합니다.

1. Private 실행본의 VERSION
2. Private CHANGELOG
3. GitHub commit message
4. 본 공개 문서 `DCI/readme.md`의 **현재 개발 버전 / 주요 기능 / Version history**

즉, 실행본 변경과 공개 개발 이력이 서로 어긋나지 않도록 동일 버전 기준으로 관리합니다.

## Source / execution policy

이 저장소는 **개발 이력과 검증 개요만 공개**합니다.

- 실행 가능한 HTML/JavaScript: 공개 저장소에 미배포
- Product DB 원본: 공개 저장소에 미배포
- 실행본: 비공개 Google Drive 또는 Private Git repository에서 관리
- Public HTML Preview / GitHub Pages: 사용하지 않음

## 제한사항

- Conceptual / pre-RFQ 설계 도구이며 최종 시공도, IFC, 전기 보호협조, 구조검토를 대체하지 않습니다.
- 실제 광손실은 transceiver application, connector grade, polarity, patching route, vendor datasheet를 기준으로 재검토해야 합니다.
- UPS/Generator/Transformer는 초기 sizing이며 실제 전기설계에는 계통전압, 역률, 고조파, 보호협조, 연료/배기/법규 검토가 추가로 필요합니다.
- 제품 추천은 호환성/용량/성능 조건 기반 후보 제시이며 최종 구매 승인이나 가격 보장은 아닙니다.
