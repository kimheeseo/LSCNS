# Case 01 — Google TPU v4

## 검증 목적

Google TPU v4 공개 논문에서 제시한 **4×4×4 building block, rack, optical link, OCS 구조**를 공통 BOM 엔진이 수식적으로 재구성하는지 검증한다.

- 검증 유형: **B — Reference-input 기반 파생 검증**
- 주의: 논문에서 그대로 입력한 값은 오차율 평가에서 제외하고, **코드가 계산한 파생값만 scored metric**으로 사용한다.
- Shared engine: `../engine/multi_arch_bom_engine.js`

## Reference

Jouppi et al., *TPU v4: An Optically Reconfigurable Supercomputer for Machine Learning with Hardware Support for Embeddings*, ISCA 2023.  
https://arxiv.org/pdf/2304.01433

### 원문 근거

논문 Section 2.1은 다음과 같이 설명한다.

> “64 TPU v4 chips and their 16 CPU hosts comfortably fit into one rack.”

Section 2.2의 핵심 수치는 아래와 같다.

| 논문에서 언급된 항목 | Reference value | 논문 위치 / 표현 |
|---|---:|---|
| Building block | 4×4×4 = 64 TPU | Sec. 2.1 |
| CPU host 구성 | 4 TPU / host | Sec. 2.1 |
| Rack 구성 | 64 TPU + 16 CPU hosts / rack | Sec. 2.1 |
| Optical links | 16 links/face × 6 faces = 96/block | Sec. 2.2 |
| OCS 연결 수 | 48 OCS / block | Sec. 2.2 |
| Palomar OCS | 136×136 = 128 working + 8 spare | Sec. 2.2 |
| 전체 시스템 | 64 blocks × 64 TPU = 4,096 TPU | Sec. 2.2 |
| 전체 rack | 64 racks | Fig. 3 / Sec. 2.2 |

## 코드 입력과 검증값 분리

### 엔진 입력으로 사용한 Reference 값

| Input | Value | 비고 |
|---|---:|---|
| target_accelerators | 4,096 | 시스템 목표 규모 |
| chips_per_host | 4 | 논문 입력 |
| dimensions | [4,4,4] | 논문 입력 |
| faces | 6 | 3D cube |
| links_per_face | 16 | 논문 입력 |
| OCS total ports | 136 | 논문 입력 |
| OCS spare ports | 8 | 논문 입력 |

이 값들은 **정답으로 재출력됐다고 해서 정확도 점수에 포함하지 않는다.**

## 실제 scored output

| 검증 항목 | Reference | Engine | 산식 | Error |
|---|---:|---:|---|---:|
| CPU hosts | 1,024 | 1,024 | 4096 / 4 | 0.00% |
| Compute racks | 64 | 64 | 4096 / 64 | 0.00% |
| TPU / rack | 64 | 64 | 4×4×4 | 0.00% |
| Optical links / rack | 96 | 96 | 6×16 | 0.00% |
| Rack→OCS endpoints | 6,144 | 6,144 | 64×96 | 0.00% |
| OCS count | 48 | 48 | 6144 / (136-8) | 0.00% |
| Working ports / OCS | 128 | 128 | 136-8 | 0.00% |
| Working ports total | 6,144 | 6,144 | 48×128 | 0.00% |
| Spare ports total | 384 | 384 | 48×8 | 0.00% |

### Review MAPE

- Scored metrics: **9**
- MAPE: **0.00%**
- Max error: **0.00%**
- Result: **PASS (<10%)**

## 검토 결론

Case 01은 단순한 값 복사가 아니라 topology 입력에서 **rack, optical endpoint, OCS 수량을 파생**하므로 유효한 구조 검증이다. 다만 136-port OCS와 8 spare는 논문 입력이므로 해당 값 자체를 정확도 항목으로 다시 세면 안 된다.

Cable length, connector/ODF 수량, 실제 설치 route는 논문에 공개되지 않아 검증 대상에서 제외한다.

## Version 2 Independent Validation

| Item | v2 result |
|---|---:|
| Numerical result | PASS |
| MAPE | 0.0000% |
| Maximum error | 0.0000% |
| Coverage | 100.0000% |
| Validation level | **A** |
| Direct output-count inputs | 0 |

### Reference comparison

The engine calculates from `design_input.json` only, then compares the output with `reference.json`. The numeric error above is therefore the reference-versus-calculation error; unsupported or undisclosed fields remain outside the MAPE.

### Improvement from version 1

- Adds an explicit input-independence audit instead of treating a low numerical MAPE alone as A-grade evidence.
- Flags direct Leaf/Spine/ToR/Rack/Cable/Optic/OCS count-like fields when present in the design input.
- Exports `validation_v2.json` with MAPE, maximum error, coverage, and a validation level in one reproducible record.

### Interpretation

No direct output-count field found in design input.
