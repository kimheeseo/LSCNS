# Case 29 — OLCF Frontier

## 검증 목적
Frontier 공식 User Guide의 **rack → node → accelerator/GCD → Slingshot NIC** hierarchy를 공통 HPC 엔진으로 재구성한다.

- Validation class: **A- — HPC architecture derivation**
- Source: https://docs.olcf.ornl.gov/systems/frontier_user_guide.html

## Reference 원문/표 근거

현재 Frontier User Guide의 System Overview는 **77 Olympus rack cabinets**, rack당 **128 AMD compute nodes**, 총 **9,856 compute nodes**라고 명시한다.

Compute Node section은 다음 구조를 공개한다.

| Published item | Official value |
|---|---:|
| Compute nodes | 9,856 |
| Racks | 77 |
| Nodes / rack | 128 |
| MI250X / node | 4 |
| GCD / MI250X | 2 |
| Visible GPUs(GCD) / node | 8 |
| Slingshot NIC / node | 4 |
| NIC speed | 200 Gbps |
| Node injection bandwidth | 800 Gbps |

짧은 원문 표현: “4x HPE Slingshot 200 Gbps” 및 “node-injection bandwidth of 800 Gbps.”

## 코드 입력
```json
{
  "nodes": 9856,
  "accelerators_per_node": 4,
  "visible_gpus_per_node": 8,
  "nics_per_node": 4,
  "nic_speed_gbps": 200,
  "racks": 77
}
```

## Reference vs Engine

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| Compute nodes | 9,856 | 9,856 | 0.00% |
| MI250X accelerators | 39,424 | 39,424 | 0.00% |
| Visible GCD GPUs | 78,848 | 78,848 | 0.00% |
| Slingshot NICs | 39,424 | 39,424 | 0.00% |
| Injection BW / node | 800 Gbps | 800 Gbps | 0.00% |
| Nodes / rack | 128 | 128 | 0.00% |

**MAPE: 0.00% · PASS.**

## Source conflict note
OLCF의 별도 allocation 안내 페이지에는 9,408 nodes라고 적힌 시점의 문서도 존재한다. 본 Case는 **현재 Frontier User Guide의 System Overview(9,856 nodes)**를 기준 source로 고정했다. 따라서 source version/date를 함께 보존해야 한다.

## 검토 결론
이 Case는 GPU 서버 vendor profile이 아니라 실제 HPC system hierarchy 전체를 검증한다. 다만 Slingshot switch 총수와 실제 cable count는 User Guide에서 이 표 수준으로 공개되지 않아 scoring하지 않는다.
