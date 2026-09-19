# Case 10 — Meta Grand Teton / Open Rack v3

## 검증 목적

Meta ORv3의 공개 rack power와 BBU shelf 사양에서 **필요 BBU shelf 수, pair power, backup seconds**를 계산한다.

- 검증 유형: **B — Power BOM derivation**
- Source: https://engineering.fb.com/2022/10/18/open-source/ocp-summit-2022-grand-teton/

## Reference 원문

ORv3 power section은 multiple shelves로 **30 kW racks**를 지원한다고 설명하며, BBU에 대해 다음과 같이 적는다.

> “provides 30kW when installed as a pair.”

같은 문단의 공개 사양:

| Parameter | Reference |
|---|---:|
| Rack power target | 30 kW |
| BBU capacity / shelf | 15 kW |
| Backup duration | 4 minutes |
| Pair capacity | 30 kW |

## Engine derivation

```
required BBU shelves = ceil(30 / 15) = 2
BBU pair power       = 15 × 2 = 30 kW
backup time          = 4 × 60 = 240 s
```

## Reference vs Code

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| Required BBU shelves | 2 | 2 | 0.00% |
| BBU pair capacity | 30 kW | 30 kW | 0.00% |
| Backup duration | 240 s | 240 s | 0.00% |

### Review result

- MAPE: **0.00%**
- Coverage: **3/3**
- Result: **PASS**
- Validation strength: **B**

## 검토 결론

Case 10은 network BOM 검증이 아니라 power BOM 산식 검증이다. 같은 30 kW라는 수치라도 rack power shelf와 BBU pair capacity는 의미가 다르므로 README에서 구분해 기록한다.

## Version 2 Independent Validation

| Item | v2 result |
|---|---:|
| Numerical result | PASS |
| MAPE | 0.0000% |
| Maximum error | 0.0000% |
| Coverage | 100.0000% |
| Validation level | **A-** |
| Direct output-count inputs | 0 |

### Reference comparison

The engine calculates from `design_input.json` only, then compares the output with `reference.json`. The numeric error above is therefore the reference-versus-calculation error; unsupported or undisclosed fields remain outside the MAPE.

### Improvement from version 1

- Adds an explicit input-independence audit instead of treating a low numerical MAPE alone as A-grade evidence.
- Flags direct Leaf/Spine/ToR/Rack/Cable/Optic/OCS count-like fields when present in the design input.
- Exports `validation_v2.json` with MAPE, maximum error, coverage, and a validation level in one reproducible record.

### Interpretation

No direct output-count field found in design input.
