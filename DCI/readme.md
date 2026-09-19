# AI Data Center BOM Engine

AI 데이터센터의 **GPU/Server → Network Fabric → Rack → Optical Connectivity → Power → Cooling → BOM**을 계산하는 브라우저 기반 설계 프로토타입입니다.

## ▶ 실행

### GitHub Pages
**[AI Data Center BOM Engine 실행하기](https://kimheeseo.github.io/LSCNS/DCI/)**

GitHub Pages가 아직 활성화되지 않았다면 아래 링크로 즉시 실행할 수 있습니다.

### HTML Preview
**[index.html 바로 실행하기](https://htmlpreview.github.io/?https://raw.githubusercontent.com/kimheeseo/LSCNS/main/DCI/index.html)**

## 파일

- `index.html` — 최종 통합 v2.6.1 웹 애플리케이션
- `readme.md` — 실행 링크 및 안내

## 주요 기능

- GPU / Server Profile DB
- Switch / Port Mode DB
- Leaf / Spine / Rail topology sizing
- Rack iterative placement
- Optical transceiver / DAC / AOC / cable / connector BOM
- Patch Panel / ODF sizing
- A/B Power Connectivity / PDU sizing
- Compute / Storage / In-Band / OOB Network 분리
- Cooling / Thermal / CDU sizing
- Golden Case validation
- BOM CSV / Design JSON export

> 비용/RFQ 기능은 현재 버전에서 제외되어 있습니다.

## GitHub Pages 설정

Pages 링크가 열리지 않는 경우 저장소에서 다음만 한 번 설정하면 됩니다.

`Settings → Pages → Build and deployment → Deploy from a branch → main / (root)`

설정 후 실행 주소:

`https://kimheeseo.github.io/LSCNS/DCI/`


## v2.6 업데이트

- **9. Rack 사이 실제 연결 도식 수정**
  - SVG 단독 의존을 제거하고 HTML Rack–Cable–Rack 도식 + 상세 SVG를 함께 표시
  - Compute Rack → Network Rack → Spine Rack 연결 수량, media, connector, fiber 수 자동 반영
- **상단 최종 설계도 + BIM-ready Summary 추가**
  - Conceptual floor plan
  - Compute / Network / Spine Rack schedule
  - Rack 수, IT 전력, 열부하, Cooling, PDU A/B, Optical path 요약
  - BIM planning assumption: rack 600×1200 mm, cold aisle 1200 mm, hot/service aisle 1000 mm
  - Design JSON에 BIM 기본 데이터 포함

> BIM 화면은 개념설계 및 Revit/IFC 입력 준비용 요약입니다. 실제 시공 BIM/IFC 모델은 프로젝트 건축·MEP 기준으로 별도 확정해야 합니다.


## v2.6.1 Runtime Fix

- STEP 1의 **계산 / 도식 업데이트** 버튼이 동작하지 않던 JavaScript runtime 오류 수정
- Rack side-reference에서 scope 밖의 `shown` 변수를 참조하던 오류 수정
- BIM-ready Summary가 초기 로드 및 재계산 후 정상 갱신되도록 복구
- Browser smoke test:
  - 초기 64 GPU → Compute Rack 2 / Total Rack 4
  - GPU 128로 변경 후 버튼 실행 → Compute Rack 4 / Total Rack 6
  - BIM KPI / floor-plan / power / thermal 값 동시 갱신 확인
