# Case 01 — Google TPU v4 (RE-VALIDATED)

**Purpose of this case:** validate one shared, extensible BOM calculation engine that will be reused unchanged across the 30 reference cases. The engine must not contain case-specific expected answers.

**Status:** `PASS`  
**Shared engine:** `../engine/multi_arch_bom_engine.js`  
**MAPE:** **0.00%**  
**Maximum error:** **0.00%**  
**Coverage:** **100.0% (12 / 12 selected verifiable metrics)**  
**PASS threshold:** **< 10% error**

## What changed from the first attempt

The first attempt put Google TPU-v4 golden-case logic inside `DCI/index.html`. That was not the intended final methodology, because the goal is not to tune the UI/code to each paper.

That logic has been removed from `DCI/index.html`.

Case 01 is now re-run through a **shared architecture engine**. The same engine file is intended to process Google, Meta, ByteDance, Alibaba, NVIDIA, HPC and other cases by changing only `design_input.json`.

## Input supplied to the engine

```json
{
  "target_accelerators": 4096,
  "accelerator": {"per_host": 4},
  "building_block": {"dimensions": [4, 4, 4]},
  "rack": {"blocks_per_rack": 1},
  "topology": {
    "type": "optical_torus",
    "faces": 6,
    "links_per_face": 16,
    "opposing_faces_share_ocs": true,
    "ocs": {"total_ports": 136, "spare_ports": 8}
  }
}
```

No value such as `rack_count=64`, `ocs_count=48`, or `working_ports_total=6144` is supplied as an expected result to the calculation function.

## Reference vs shared-engine output

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| TPU chips | 4,096 | 4,096 | 0.00% |
| CPU hosts | 1,024 | 1,024 | 0.00% |
| Compute racks | 64 | 64 | 0.00% |
| TPU / rack | 64 | 64 | 0.00% |
| Optical links / rack | 96 | 96 | 0.00% |
| Rack→OCS link endpoints | 6,144 | 6,144 | 0.00% |
| OCS count | 48 | 48 | 0.00% |
| OCS ports / unit | 136 | 136 | 0.00% |
| Working ports / OCS | 128 | 128 | 0.00% |
| Spare ports / OCS | 8 | 8 | 0.00% |
| Working OCS ports total | 6,144 | 6,144 | 0.00% |
| Spare OCS ports total | 384 | 384 | 0.00% |

## Generic derivation

```
accelerators_per_block = product([4,4,4]) = 64
rack_count             = ceil(4096 / 64) = 64
host_count             = 4096 / 4 = 1024
optical_links_per_rack = 6 × 16 = 96
rack_ocs_endpoints      = 64 × 96 = 6144
working_ports_per_ocs  = 136 - 8 = 128
ocs_count               = 6144 / 128 = 48
```

## Scope

The 0% result applies only to deterministic quantities publicly disclosed or directly derivable from the TPU-v4 paper. It does **not** imply that undisclosed installation details such as actual cable lengths, patch panels, connector part numbers, tray routing or field spares are known.

## Rule for Cases 02–30

- Do not add expected answers to the shared engine.
- Add only a new case `design_input.json` and `reference.json`.
- Extend the engine only when a genuinely new architecture primitive is required (for example Clos, rail, dual-ToR, Dragonfly), and the extension must remain generic for later cases.
- Commit each completed case separately and update the top-level validation matrix after each case.
