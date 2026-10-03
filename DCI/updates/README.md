# DCI Development Updates

## Current runtime

**v7.3.2**

현재 GitHub Pages 실행본은 `DCI/DataCenter/index.html`의 `data-dc-bom-version="7.3.2"` 기준입니다.

- Runtime: `DataCenter/index.html`
- Latest UI layer: `v73-20260930.css / v73-20260930.js`
- Public URL: https://kimheeseo.github.io/LSCNS/DCI/DataCenter/
- Deployment: **GitHub Pages only**

## 주요 버전 흐름

### v7.3.2 — Current
- 2026-09-30 기준 UI / engineering workflow 통합
- Rack / Cooling / Optical / Supply Chain / structured cabling 기능을 현재 실행본에 통합
- GitHub Pages 정적 실행 구조 유지

### v7.0.x — Diagram rework
- Physical / optical diagram 표현 개선
- 데이터센터 topology 및 연결 구조의 시각화 정리

### v6.8.0 — Visualization Hub
- Visualization Hub
- Animated Cooling
- Logic Diagram
- Structured Cabling

### v6.x — Engineering UI expansion
- Rack / cooling / compute silicon / CPO link
- 1P2T aggregation / colocation
- UI hardening 및 시각화 기능 확장

### v5.x — Dynamic BOM Supply Chain
- Generic BOM과 Product Mapping 결과를 공급망 관점으로 재구성
- Compute / Network / Optical / Power / Cooling / Rack / Storage / Facility 영역 연결

### v4.x — Product-aware advisor
- GPU 규모 기반 reference-size class
- Optical connectivity advisor
- Port → Optic → Connector → Fiber → Trunk 연결 모델

### v3.x — Solver hardening
- Fabric policy와 validation fixture 분리
- Physical connectivity / fiber-count policy / product DB / facility / storage 확장

### v2.x — Initial integrated engine
- GPU / Server → Fabric → Rack → Optical → Power → Cooling → BOM 설계 흐름 통합

## Archive

기존 DCI README에 기록되어 있던 상세 변경 이력은 아래에 보존합니다.

- [legacy_full_history.md](./legacy_full_history.md)

향후 개발 버전의 상세 변경사항은 이 `updates/` 폴더에 계속 추가합니다.
