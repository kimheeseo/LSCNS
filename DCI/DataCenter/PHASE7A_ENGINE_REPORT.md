# Phase 7-A 공통 BOM 엔진 보고서

## 구현

- 입력은 기존 `T.exportConfig()`가 반환하는 `lsdc-twin-bom/1.0` packet 하나입니다.
- `phase7-bom-engine.js`가 서버 랙, GPU 서버, GPU, NIC, Leaf/ToR, Spine, Core, 활성 광 링크와 전체 광 길이를 공통 `rows` 및 `counts`로 반환합니다.
- Phase 7 화면은 별도 랙/uplink/거리 산식 입력을 없애고 snapshot을 계산해 표시합니다. 엔진 결과와 3D 렌더링용 자산 수는 분리되어 있습니다.
- SKU·물리 cage·광 코어 등의 정보가 packet에 없으므로 엔진은 이를 구매 BOM 수량으로 만들어내지 않습니다.

## 계산 예시

| 입력 packet | 랙 | GPU 서버 | GPU | NIC | Leaf / Spine / Core | 광 링크 | 광 길이 |
|---|---:|---:|---:|---:|---:|---:|---:|
| 검증 packet: 100 racks, GPU ratio 0.5, 서버당 8 GPU, NIC 1개 | 100 | 50 | 400 | 50 | 2 / 2 / 1 | 2 | 150 m |
| 200 racks, GPU ratio 0.5, 기본 가정 | 200 | 100 | 800 | 100 | packet 토폴로지 기준 | packet 링크 기준 | packet 길이 합 |
| 현재 AIDC 기본 설정: 20 racks, GPU ratio 0.8 | 20 | 16 | 128 | 16 | 2 / 2 / 1 | 9 | 1,355 m |

GPU는 packet에 명시 수량이 없을 때 서버당 8개, NIC는 NIC/server 값이 없을 때 GPU 서버당 1개 가정으로 표시합니다. 기본 가정은 각 행의 `source`와 `basis`에 남깁니다.

## 검증

- `tests/phase7_engine.test.cjs`: 6 assertions 통과. 랙 100→200 변화와 서버/GPU/NIC 수량 변화, active 광 링크 길이 합산 및 스키마 거부를 검증했습니다.
- 제품 SKU, GPU 서버당 가속기 구성, NIC 수량이 packet에 제공되면 가정 대신 해당 입력값을 사용합니다.
