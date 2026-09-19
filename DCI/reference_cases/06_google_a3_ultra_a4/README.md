# Case 06 — Google A3 Ultra H200 / A4 B200

## 검증 목적

동일한 공통 machine-profile 로직이 서로 다른 GPU 세대(H200/B200)에 대해 NIC, memory, bandwidth BOM을 올바르게 구성하는지 검증한다.

- 검증 유형: **C — Profile consistency validation**
- 공식 machine table: https://docs.cloud.google.com/compute/docs/gpus
- network architecture: https://docs.cloud.google.com/compute/docs/gpus/gpu-network-bandwidth

## Reference table

| Machine | GPU | Physical NIC | Max network | GPU count | GPU memory |
|---|---|---:|---:|---:|---:|
| a3-ultragpu-8g | H200 | 10 | 3,600 Gbps | 8 | 1,128 GB |
| a4-highgpu-8g | B200 | 10 | 3,600 Gbps | 8 | 1,440 GB |

Networking 문서는 두 플랫폼에 대해 **8×CX-7 + 2×gVNIC** 구조를 설명한다.

> “The eight CX-7 NICs deliver a total network bandwidth of 3,200 Gbps.”

또한 2개의 gVNIC가 추가 400 Gbps를 제공하므로 총 3,600 Gbps가 된다.

## Reference vs Code

| Profile | Metric | Reference | Engine | Error |
|---|---|---:|---:|---:|
| A3 Ultra | GPUs | 8 | 8 | 0.00% |
| A3 Ultra | Physical NICs | 10 | 10 | 0.00% |
| A3 Ultra | GPU memory | 1,128 GB | 1,128 GB | 0.00% |
| A3 Ultra | GPU NIC/GPU | 1.0 | 1.0 | 0.00% |
| A3 Ultra | Max network | 3,600 Gbps | 3,600 Gbps | 0.00% |
| A4 | GPUs | 8 | 8 | 0.00% |
| A4 | Physical NICs | 10 | 10 | 0.00% |
| A4 | GPU memory | 1,440 GB | 1,440 GB | 0.00% |
| A4 | GPU NIC/GPU | 1.0 | 1.0 | 0.00% |
| A4 | Max network | 3,600 Gbps | 3,600 Gbps | 0.00% |

### Review result

- MAPE: **0.00%**
- Coverage: **10/10**
- Result: **PASS**
- Validation strength: **C / consistency**

## 검토 결론

Case 06은 장비 catalog parser/profile 계산을 검증하는 데 유효하다. 다만 vendor table을 그대로 profile로 입력한 값이 많아 predictive accuracy의 강한 증거는 아니다. 향후 이 profile을 기반으로 rack density, switch, cable, power 수량을 예측하는 Case가 필요하다.
