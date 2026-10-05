# AI Data Center DCI

AI 데이터센터의 **Compute / Network / Optical / Rack / Power / Cooling / BOM**을 설계·검토하기 위한 정적 웹 도구입니다.

## 바로 실행

- [▶ DataCenter Tool](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/)
- [LS Datacenter Campus · Gold Pixel Tour](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_Campus.html?reset=1)
- [CPO Supply Chain](https://kimheeseo.github.io/cpo-supply-chain/)

> DataCenter 화면은 **LSCNS GitHub Pages**에서 실행되고 핵심 계산·BOM 제품 매칭은 Private backend에서 실행됩니다. CPO는 별도 `cpo-supply-chain` GitHub Pages에서 관리합니다.

## 주요 기능

- GPU / Rack / Leaf-Spine-Core 네트워크 설계
- Optical / Fiber / Connector / Transceiver BOM 산정
- Power / Cooling sizing 및 2D/3D 시각화
- Supply Chain / Product Mapping
- Excel / CSV / Design JSON export

## Repository

- `DataCenter/` — 공개 UI / 시각화 / runtime assets. 핵심 계산 엔진과 제품 매칭 source는 private `kimheeseo/others`에서 관리
- CPO Supply Chain — 별도 repository: `kimheeseo/cpo-supply-chain`
- `current_version/` — 현재 개발 버전
- `updates/` — 날짜별 업데이트 이력

Validation / reference case는 private `kimheeseo/others/DCI/`에서 관리합니다.

- [현재 개발 버전](./current_version/README.md)
- [업데이트 내역](./updates/README.md)


## Source protection

- Private source: `kimheeseo/others/DCI/DataCenter/`
- Production calculation API: Railway private service
- Public Pages에는 핵심 `dc-engine.min.js`와 `catalog-design-match.js`를 두지 않습니다.
- 브라우저는 입력 JSON을 backend로 보내고 결과 JSON만 받아 표시합니다.
- 제품 카탈로그 자체는 제조사 공개 자료이므로 public reference data로 유지할 수 있습니다.

> 주의: 과거 public Git commit에는 이전 소스가 남아 있을 수 있습니다. 완전한 과거 이력 제거는 별도의 history rewrite 또는 runtime-only public repository 재생성이 필요합니다.
