# AI Data Center DCI

AI 데이터센터의 **Compute / Network / Optical / Rack / Power / Cooling / BOM**을 한 화면에서 설계·검토하기 위한 정적 GitHub Pages 도구입니다.

## 바로 실행

- **DC BOM Designer**  
  https://kimheeseo.github.io/LSCNS/DCI/DataCenter/

- **LS Datacenter Campus · Gold Pixel Tour**  
  https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_Campus.html

> 공개 실행 페이지는 Render가 아니라 **GitHub Pages**에서 직접 제공합니다.

## 주요 기능

- GPU / AI system 수량과 Rack 배치 산정
- Leaf / Spine / Core / Rail 기반 네트워크 fabric 설계
- Optical interface, transceiver, connector, fiber / trunk BOM 계산
- Rack 전력, A/B PDU, UPS / Generator / Transformer 초기 sizing
- Air / Liquid cooling, CDU / manifold / cold-plate 구조 시각화
- Rack 2D/3D, Physical Connectivity, Cooling / Power / Optical diagram
- 제품·제조사 기반 Supply Chain / Product Mapping
- Excel / CSV / Design JSON export
- 한국어 / English / 日本語 / 中文 / Deutsch 및 모바일 반응형 UI

## 저장소 구조

- `DataCenter/` — GitHub Pages 실행용 HTML / CSS / JavaScript
- `CPO/` — DCI와 연계되는 CPO 정적 리소스
- `updates/` — 현재 개발 버전 및 변경 이력
- Validation / reference case 원본은 private `kimheeseo/others`의 `DCI/`로 이동하여 관리

## 개발 이력

현재 버전과 상세 업데이트 내역은 아래 문서에서 관리합니다.

- [Development Updates](./updates/README.md)
- [Legacy Full History](./updates/legacy_full_history.md)
