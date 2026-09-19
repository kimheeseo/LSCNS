# AI Data Center BOM Engine

AI 데이터센터의 **GPU/Server → Network Fabric → Rack → Optical Connectivity → Power → Cooling → BOM**을 계산하는 브라우저 기반 설계 프로토타입입니다.

## ▶ 실행

### GitHub Pages
**[AI Data Center BOM Engine 실행하기](https://kimheeseo.github.io/LSCNS/DCI/)**

GitHub Pages가 아직 활성화되지 않았다면 아래 링크로 즉시 실행할 수 있습니다.

### HTML Preview
**[index.html 바로 실행하기](https://htmlpreview.github.io/?https://raw.githubusercontent.com/kimheeseo/LSCNS/main/DCI/index.html)**

## 파일

- `index.html` — 최종 통합 v2.5 웹 애플리케이션
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
