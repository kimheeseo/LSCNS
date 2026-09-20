# AI Data Center BOM Engine

AI 데이터센터 물리 인프라/BOM 설계 프로토타입의 **개발 이력 및 검증 개요**를 정리한 문서입니다.

> **Source policy:** 실행 가능한 HTML/JavaScript 소스는 이 공개 저장소에 배포하지 않습니다.  
> 실제 실행본과 제품 DB는 비공개 저장소 또는 제한된 파일 저장소에서 관리합니다.

## 현재 개발 버전

**v3.8 Multilingual / Neutral**

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
- Conceptual BIM-ready floor plan / rack schedule
- CSV BOM / Validation CSV / Design JSON export
- 한국어 / 日本語 / English / 中文 / Deutsch 내장 UI 번역

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
