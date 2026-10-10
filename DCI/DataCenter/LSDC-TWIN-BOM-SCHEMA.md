# LS DataCenter Twin ↔ BOM 공유 스키마

- 스키마 ID: `lsdc-twin-bom/1.0`
- 검증 정의: [lsdc-twin-bom.schema.json](./lsdc-twin-bom.schema.json)
- 값은 모두 **교육·설계 검토용 가정값**이며 실제 제품 사양이나 시공 승인값이 아닙니다.
- 데이터 흐름: 3D 트윈에서 JSON 파일 또는 해시 공유 링크를 생성하고, BOM 도구가 이를 입력 초안으로 매핑합니다. 3D 트윈은 같은 스키마 JSON을 다시 불러올 수 있습니다. BOM 계산 결과의 역동기화는 구현하지 않았습니다.

## 단위 및 필드

| JSON 경로 | 단위 / 형식 | 뜻 |
|---|---|---|
| `schemaVersion` | 문자열 | `lsdc-twin-bom/1.0` |
| `units.power / length / temperature / loss / time` | kW / m / degC / dB / min | 숫자 필드의 단위 표기 |
| `assumptionNotice` | 문자열 | 가정값 고지 |
| `facility.profile` | enum | `aidc`, `server-room`, `onprem` |
| `facility.rackCount` | 대 | 전체 랙 수 |
| `facility.rackPowerKw` | kW/랙 | 평균 랙 IT 전력 |
| `facility.gpuServerRatio` | 0–1 | GPU 서버 비율 |
| `facility.cooling.mode` | enum | `air`, `rear`, `dlc`, `mixed` |
| `facility.cooling.*Pct` | % | 냉각 방식별 구성 비율 |
| `facility.cooling.capacityKw` | kW | 가정 냉각 용량 |
| `facility.redundancy` | enum | `N`, `N+1`, `2N` |
| `facility.targetPue / pueRange` | 비율 | 목표 PUE 및 참고 범위 |
| `facility.outdoorC` | °C | 외기 온도 가정값 |
| `facility.generatorKw / upsKw` | kW | 발전기·UPS 용량 가정값 |
| `facility.batteryMinutes` | 분 | 배터리 백업 시간 가정값 |
| `optical.topology.type / radix / portSpeed` | enum / 포트 / 문자열 | Spine-Leaf 또는 Fat-tree 토폴로지 설정 |
| `optical.topology.nodes[].layer` | enum | Rack, Leaf, Spine, Core, MMR, External 계층 |
| `optical.topology.links[].from / to` | ID | 양 끝 노드 ID |
| `optical.topology.links[].type / fiber` | 문자열 | 링크/케이블 유형 및 SMF/MMF 매체 |
| `optical.topology.links[].cores` | 코어 | 광 코어 수 가정값 |
| `optical.topology.links[].lengthM` | m | 링크 길이 가정값 |
| `optical.topology.links[].connectors / splices` | 개 | 손실 계산 입력 개수 |
| `optical.topology.links[].speed` | 문자열 | 링크 속도 예시/가정 |
| `optical.topology.links[].budgetDb` | dB | 링크 손실/마진 참고값 |
| `optical.cpoMode` | enum | `pluggable`, `cpo` |
| `optical.assumptions` | 객체 | 포트 전력·스위치 전력 등 사용자 가정값 |
| `layoutData.zones[] / assets[]` | 객체 배열 | 3D 캠퍼스 구역·표시 설비의 내보내기 스냅샷 |

## BOM 입력 매핑

BOM 브리지는 랙 수, 랙 전력, 냉각 방식과 용량, 이중화, 대표 케이블 길이, 트렁크 코어 수를 BOM 입력 초안에 채웁니다. 광 topology는 BOM의 Rail/3-tier 선택지에 매핑합니다. GPU 수 초안은 **GPU 대상 랙당 GPU 서버 1대, 서버당 GPU 8개**라는 별도 계산 가정입니다. CPO/플러거블과 원본 링크 목록은 요약에 보존되지만, BOM 도구의 기존 산출 엔진이 이 모든 필드를 제품 단위 수량으로 직접 계산하지는 않습니다. 가져온 입력을 검토한 뒤 사용자가 BOM 계산을 실행합니다.

3D 배치 미리보기는 성능상 최대 40랙까지만 표시할 수 있으며 전력 계산에는 설정된 전체 랙 수를 사용합니다. JSON에는 현재 표시 중인 구역/설비 요약과 전체 광 링크가 포함됩니다.

## 호환성

BOM 가져오기는 공유 링크(`#twin=<base64url JSON>`)와 JSON 파일을 지원합니다. 3D 트윈도 이 스키마의 JSON 파일을 불러옵니다. 잘못된 버전과 범위를 벗어난 주요 BOM 입력은 거부합니다. 브라우저 저장소를 구성 공유 수단으로 사용하지 않습니다.
