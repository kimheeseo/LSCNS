# GN-model GSNR calculator — revised

This repository contains a cleaned and tested implementation of the model used in:

> *Capacity and Cost Investigations of Long-Span-Length, High-Number Spatial-Pairs Trans-Oceanic Cables*, SubOptic 2025.

The code combines:

- closed-form GN-model nonlinear interference,
- generalized droop accumulation,
- ASE and multicore crosstalk,
- submarine PFE power constraints,
- nonlinear-threshold backoff and an 18 dBm output cap,
- Shannon-like capacity optimization over repeater span length.

## Main improvements over the supplied scripts

1. Added the missing conversion helpers and removed broken imports.
2. Added explicit validation and documented units.
3. Converted crosstalk from dB/km when requested.
4. Uses an integer number of complete WDM channels.
5. Eliminated the hard-coded `1000` indexing limit and off-by-one search.
6. Matches EDFA efficiency to the actual selected launch-power point.
7. Returns GSNR at the maximum-capacity point, rather than an unrelated maximum.
8. Separates physical span/repeater counting from the paper-compatible convention.
9. Implements the paper's 2 dB nonlinear-threshold backoff and 18 dBm cap.
10. Added tests and a structured detailed result object.

## Installation and tests

```bash
python -m pip install -e ".[test]"
pytest
```

## Basic GSNR example

```python
import numpy as np
from gn_model_gsnr import SNR_GD

powers_w = np.geomspace(1e-3, 0.063, 200)
gsnr = SNR_GD(
    powers_w,
    lamb=1550,
    alpha=0.15,
    NF=4.5,
    N=75,
    gamma=0.81,
    Ltot=6000,
    B=4.3e12,
    D=-21e-12,
    Rs=100e9,
    gamma_XT=-80,
    l_FIFO=0.3,
    delta_f=112.5e9,
    gamma_xt_unit="dB_per_km",
)
```

## Capacity optimization

`optimized_capacity_extended` retains the original positional API. Set
`return_details=True` to obtain a `CapacityOptimizationResult` with span counts,
repeater counts, feasibility flags, unconstrained nonlinear thresholds, and all
per-span optimum values.

### Counting modes

- `count_mode="physical"` (default): `ceil(Ltot/Lspan)` spans and one fewer repeater.
- `count_mode="paper"`: reproduces the legacy/paper `floor(Ltot/Lspan)` convention.

The two modes are intentionally explicit because the paper and the original code
use the word *repeater* where the GSNR equation requires a span count.

## Current model limits

This remains a flat-spectrum, identical-span, closed-form GN approximation. It does
not yet include channel-by-channel ISRS/GGN, modulation-dependent EGN correction,
span-to-span parameter variation, probabilistic crosstalk, PCS/GMI, or terminal and
repair costs. Those are the recommended next research extensions.
