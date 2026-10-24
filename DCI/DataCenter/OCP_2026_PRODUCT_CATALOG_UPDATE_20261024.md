# OCP Global Summit 2026 제품 카탈로그 반영

확인일: 2026-10-24  
행사: OCP Global Summit 2026, San Jose, 2026-10-12~15

## DataCenter BOM 제품 DB

아래 항목을 `product_catalog`의 업체별 구조화 카탈로그에 추가하거나 갱신했습니다.

| 업체 | 제품/기술 | 카탈로그 분류 | 반영 상태 |
|---|---|---|---|
| Broadcom | Tomahawk 6–Davisson 102.4T CPO | CPO / Data Center Switching | OCP 파트너 부스 데모. 정확한 시스템 SKU와 공급 상태 확인 필요 |
| Marvell | Ara DSP · 1.6T transmit-retimed optics | Optical DSPs | OCP 전시 항목. 세부 모듈 SKU/호환성 확인 필요 |
| Marvell | Teralynx T100 102.4T switch | Data Center Switching | 제조사 전시 발표. 포트 구성·전원/냉각·SKU 확인 필요 |
| Delta | 800 VDC 12 kW PDB, eFuse, DC-DC module | Data Center Power | 공개 전시 사양 등록. SKU와 조달 상태 확인 필요 |
| Delta | HVDC 25 kW electrical pump | Data Center Cooling | 공개 운전점 정보 등록. CDU 수량/성능으로 자동 환산하지 않음 |
| Astera Labs | Scorpio Smart Fabric Switch | AI Rack Interconnect | Lumentum·Corning 공동 데모 구성품. 포트/P/N 및 공급 확인 필요 |
| Lumentum | 1060 nm VCSEL PCIe-over-optics | CPO | AI 랙 규모 공동 데모 |
| Corning | 멀티모드 광섬유 연결 | CPO | 공동 데모 연결 구성. 정확한 케이블 SKU/등급 미공개 |
| Phasic Energy | 1 MW (2×500 kW) modular heat exchanger | Data Center Cooling | OCP FTS 데모. 판매 SKU/성능 곡선 확인 필요 |
| Apheros | metal-foam direct-to-chip cold-plate substrate | Data Center Cooling | 프로토타입 시연. 열/유압 사양 미공개 |
| LightXcelerate | microVCSEL + multicore fiber CPO engine | CPO | OCP FTS 라이브 데모. 제품번호/성능/공급 상태 미공개 |

새 카탈로그 경로는 다음 6개입니다.

- `Marvell/Data Center Switching/catalog.json`
- `Delta/Data Center Cooling/catalog.json`
- `Astera Labs/AI Rack Interconnect/catalog.json`
- `Phasic Energy/Data Center Cooling/catalog.json`
- `Apheros/Data Center Cooling/catalog.json`
- `LightXcelerate/CPO/catalog.json`

카탈로그 manifest revision은 `dc-products-ocp-20261024`이며 구조화 카탈로그 경로는 499개입니다. 네트워크/냉각 역할 인덱스도 새 경로를 포함하도록 갱신했습니다.

## BOM 반영 기준

OCP 행사에서 발표·전시·시연된 사실만으로 주문 가능 제품이나 확정 SKU로 취급하지 않습니다. 행사 데모·프로토타입 중심 레코드는 `referenceOnly: true`로 표시해 자동 BOM 수량 산출과 대체품 매칭을 제한했습니다. 정확한 모델 번호, 포트 구성, 정격, 호환성, 제조사 공급 상태가 확인되면 구매 후보 카탈로그로 별도 승격할 수 있습니다.

## 공식 출처

- [OCP Global Summit 2026 Recorded Content](https://www.opencompute.org/events/ocp-summit/2026-ocp-global-summit/recorded-content/)
- [Broadcom OCP 2026 networking demonstrations](https://investors.broadcom.com/news-releases/news-release-details/broadcom-redefines-ai-infrastructure-industry-leading-networking)
- [Marvell OCP 2026 product showcase](https://investor.marvell.com/news-events/press-releases/detail/1036/marvell-to-showcase-open-ai-infrastructure-solutions-at-ocp-global-summit-2026)
- [Lumentum · Astera Labs · Corning PCIe-over-optics demo](https://investor.lumentum.com/financial-news-releases/news-details/2026/Lumentum-Astera-Labs-and-Corning-to-Demonstrate-1060-nm-VCSEL-Based-PCIe-over-Optics-for-AI-Rack-Scale-Connectivity-at-the-2026-OCP-Global-Summit/default.aspx)
- [Delta OCP Global Summit 2026 recap](https://www.delta-americas.com/en-US/landing/2026-OCP-Summit)
- [OCP Innovation Village / Future Technologies Symposium demo lineup](https://www.opencompute.org/events/ocp-summit/2026-ocp-global-summit/innovation-village/)
