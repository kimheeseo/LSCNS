# Phase 7-B 제품 호환성 보고서

## 구현

`phase7-compatibility.js`는 protocol, 링크 속도, 매체, CPO/Pluggable 모드, endpoint, SKU와 reach를 검사합니다. 반환값은 `compatible`, `nonBlocking`, `status`, `blocking[]`, `warnings[]`, `info[]`입니다.

`compatible: true`는 PASS일 때만 반환합니다. 명시적 차단 없음은 `nonBlocking: true`로 구분합니다. 공식 제품 근거가 부족하면 `status: UNKNOWN`과 경고를 함께 표시하므로 이를 제품 호환 PASS로 보지 않습니다. 제품 reach가 packet에 없으면 임의의 DAC/AEC/MMF 길이 한계를 만들어 차단하지 않습니다.

## 검증 예시

| 사례 | validator 결과 |
|---|---|
| SKU·공식 reach 누락 | `UNKNOWN`, reach/SKU 추가 확인 경고 |
| packet 설계가 InfiniBand인데 endpoint가 Ethernet | `FAIL`, blocking |
| 800G 설계에 400G 링크 | `FAIL`, 속도 blocking |
| 링크 100 m, 명시된 제품 reach 50 m | `FAIL`, reach blocking |
| CPO switch endpoint와 별도 pluggable 모듈 동시 표기 | `FAIL`, 이중 계상 blocking |

현재 v1.0 packet에는 제조사별 orderable SKU, QSFP-DD/OSFP cage, MPO polarity, FEC, 제품 reach가 충분히 실리지 않습니다. 해당 필드가 확인되기 전에는 제품별 `PASS`로 승격하지 않습니다.

## 검증

- `tests/phase7_compatibility.test.cjs`: 기본 및 보완 사례 통과.
- 공식 제조사 카탈로그는 기존 BOM 페이지의 catalog UI와 별도로 존재합니다. 이 validator는 그 파일들을 자동으로 크롤링해 규격을 추정하지 않으며, packet에 명시된 사양에 대해서만 판정합니다.

## 2026-10-11 보완

compatible은 PASS일 때만 true이며 차단 없음은 nonBlocking으로 구분합니다. 빈 products 배열은 제품 근거로 인정하지 않습니다. 활성 링크마다 SKU·데이터시트 URL·reach, 노드별 포트·프로토콜, 광 connector/FEC 정보를 검사합니다. 포트 초과와 endpoint/제품 상충은 FAIL, 근거 부족은 UNKNOWN입니다. 1.6T/1600G와 IB/InfiniBand 동의 표기를 지원합니다. PASS는 제공된 근거의 규칙 검사이며 URL 원문의 독립 검증을 의미하지 않습니다.
