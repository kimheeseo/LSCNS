# LS DataCenter Twin ↔ BOM 공유 스키마

- 스키마 ID: `lsdc-twin-bom/1.0`
- JSON Schema: [lsdc-twin-bom.schema.json](./lsdc-twin-bom.schema.json)
- 값은 모두 **교육·설계 검토용 가정값**이며 실제 제품 사양이나 시공 승인값이 아닙니다.
- 3D 트윈은 JSON 내보내기를 지원하고 BOM 도구는 JSON 파일 가져오기로 입력 초안을 채웁니다. 기존 BOM 산출 엔진은 변경하지 않습니다.

## 필드와 단위

| 필드 | 단위 | 뜻 |
|---|---:|---|
| `facility.profile` | enum | `aidc`, `server-room`, `onprem` |
| `facility.rackCount` | 대 | 전체 랙 수 |
| `facility.rackPowerKw` | kW/랙 | 평균 랙 IT 부하 |
| `facility.gpuServerRatio` | 0–1 | GPU 서버 비율 |
| `facility.cooling.mode` | enum | `air`, `rear`, `dlc`, `mixed` |
| `facility.cooling.capacityKw` | kW | 가정 냉각 용량 |
| `facility.redundancy` | enum | `N`, `N+1`, `2N` |
| `facility.targetPue` | ratio | 목표 PUE |
| `facility.outdoorC` | °C | 외기 온도 |
| `facility.generatorKw`, `upsKw` | kW | 발전기·UPS 가정 용량 |
| `facility.batteryMinutes` | 분 | 배터리 백업 가정 시간 |
| `optical.topology.nodes[].layer` | — | Rack, Leaf, Spine, Core, MMR, External 계층 |
| `optical.topology.links[].lengthM` | m | 링크 길이 가정 |
| `optical.topology.links[].fiber` | — | SMF/MMF 매체 |
| `optical.topology.links[].cores` | 코어 | 링크 코어 수 |
| `optical.topology.links[].speed` | — | 링크 속도 예시/가정 |
| `optical.topology.links[].budgetDb` | dB | 예상 링크 마진 |
| `optical.cpoMode` | enum | `pluggable`, `cpo` |
| `optical.assumptions` | mixed | 전력·밀도 등 모드 가정값 |

## BOM 가져오기 매핑

BOM 브리지는 랙 수, 랙 전력, 냉각 유형, 냉각 용량, 이중화, 광 코어 수, 대표 링크 길이를 해당 입력 필드에 채웁니다. GPU 수 초안은 **GPU 대상 랙당 GPU 서버 1대, 서버당 GPU 8개**라는 별도 변환 가정으로 산출합니다. 가져온 뒤 값과 단위를 검토하고 사용자가 BOM 계산 버튼을 눌러 산출합니다.

랙 배치 미리보기는 3D 성능을 위해 최대 20랙까지 표시하며, IT 부하 계산에는 전체 랙 수를 사용합니다. JSON 파일에는 전체 설정과 광 링크 목록이 포함됩니다.

이 버전은 3D 트윈에서 BOM으로 보내는 단방향 가져오기를 구현합니다. BOM 산출 결과를 다시 3D 트윈에 반영하는 역방향 동기화는 현재 스키마를 공유하되 별도 구현 범위입니다.
