# Case 01 — Google TPU v4

**Status:** `PASS` for public deterministic topology/BOM quantities  
**Tool:** `DCI/index.html` v2.6.2 multi-architecture golden validation  
**Calculator:** `computeOpticalTorusArchitecture()`  
**MAPE:** **0.00%**  
**Maximum error:** **0.00%**  
**Coverage:** **100.0% (12 / 12 selected verifiable metrics)**  
**Pass criterion:** error ≤ 10%.

## Reference architecture

Google TPU v4 uses one 4×4×4 = 64-chip building block per rack. Six faces × 16 links create 96 rack-side optical links. Sixty-four racks are connected through 48 Palomar optical circuit switches (OCSes), each with 136 ports: 128 working + 8 spare.

## Calculated vs reference

| Metric | Reference | Engine | Error | Status |
|---|---:|---:|---:|---|
| TPU chips | 4,096 | 4,096 | 0.00% | PASS |
| CPU hosts | 1,024 | 1,024 | 0.00% | PASS |
| Compute racks | 64 | 64 | 0.00% | PASS |
| TPU / rack | 64 | 64 | 0.00% | PASS |
| Optical links / rack | 96 | 96 | 0.00% | PASS |
| Rack→OCS link endpoints | 6,144 | 6,144 | 0.00% | PASS |
| OCS count | 48 | 48 | 0.00% | PASS |
| OCS total ports / unit | 136 | 136 | 0.00% | PASS |
| OCS working ports / unit | 128 | 128 | 0.00% | PASS |
| OCS spare ports / unit | 8 | 8 | 0.00% | PASS |
| Working OCS ports total | 6,144 | 6,144 | 0.00% | PASS |
| Spare OCS ports total | 384 | 384 | 0.00% | PASS |

## Why this is not hard-coded output fitting

The reference values are not assigned as calculated outputs. The engine receives architecture parameters:

```
dimensions       = [4,4,4]
blockCount       = 64
blocksPerRack    = 1
chipsPerHost     = 4
faces            = 6
linksPerFace     = 16
ocsTotalPorts    = 136
ocsSparePorts    = 8
```

and derives:

```
chipsPerBlock          = product(dimensions)
acceleratorCount       = chipsPerBlock × blockCount
hostCount              = acceleratorCount / chipsPerHost
opticalLinksPerBlock   = faces × linksPerFace
rackOcsLinkEndpoints   = blockCount × opticalLinksPerBlock
ocsWorkingPortsEach    = ocsTotalPorts - ocsSparePorts
ocsCount               = rackOcsLinkEndpoints / ocsWorkingPortsEach
```

For TPU v4 this independently yields:

```
64 × 96 = 6,144 rack-side optical endpoints
48 × 128 = 6,144 working OCS ports
```

## Scope limitation

The 0% error result is **not** a claim that every physical item in the real Google installation is known. Exact installed cable lengths, connector/patch-panel part numbers, tray-level routing details, and other undisclosed BOM items are excluded from error calculations.

The next validation case is Google TPU v5p, reusing the generic topology/reference framework rather than adding case-specific answer constants.
