# GN_model — GN/EGN Paper Reproducibility Validation

This folder validates the numerical GN-model implementation in:

- `../gn_integral_general.py`
- `../gn_integral_general_modulation.py`

and compares a representative subset against:

- `../EGN_model/EGN_adaptive.py`

## References

1. P. Poggiolini et al., **A Detailed Analytical Derivation of the GN Model of Non-Linear Interference in Coherent Optical Transmission Systems**.
2. A. Carena et al., **Modeling of the Impact of Nonlinear Propagation Effects in Uncompensated Optical Coherent Transmission Links**, JLT 30(10), 2012.

## What is tested

### Paper 1 — equation-level reproducibility
`paper1_reproducibility.py` checks the implementation against the GN derivation without fitting:
- power-attenuation dB/km to field attenuation conversion,
- beta2 phase mismatch,
- one-span GN link function,
- incoherent and coherent identical-span accumulation,
- generalized non-identical-span coherent phase history,
- exact first-order NLI cubic power scaling,
- scrambled-Sobol QMC convergence.

### Paper 2 — Figure 5
`figure5_benchmark.py` reconstructs the Fig. 5 maximum-reach calculation using:
- 9 WDM channels,
- 32 GBd,
- 100-km spans,
- 5-dB EDFA NF,
- PSCF / SMF / NZDSF parameters from Table I,
- the paper's Fig. 4–7 **incoherent NLI accumulation** convention,
- a paper-like NRZ `sinc²` spectrum filtered by a fourth-order super-Gaussian with `Bopt = Δf`,
- required back-to-back OSNR digitized from Fig. 3.

The paper/reference points in `paper_fig5_digitized.csv` are digitized from the supplied PDF rather than copied from an author data file. Each row therefore carries an explicit digitization-uncertainty estimate.

### GN vs EGN comparison
The full precision path in `EGN_adaptive.py` supports rectangular spectra and coherent multi-span propagation. To compare it with Carena-2012 Fig. 5 without modifying the EGN physics implementation, the benchmark:
1. computes the **full one-span EGN NLI**,
2. converts it to an NLI coefficient,
3. applies the **paper's incoherent N-span scaling** for the reach calculation.

This is intentionally labeled a nearest-scope comparison. It is not presented as a full multi-span EGN reproduction of Fig. 5.

## Colab

Open `GN_Model_Colab.ipynb`. The committed executed copy is written to:
`results/GN_Model_Colab_executed.ipynb`.

Set `RUN_HEAVY=True` in the last cell to recompute the GN and EGN numerical integrations in Colab.

## Reproduce locally

```bash
python -m pip install -r 2026_KICS_Fall_10/GN_model/requirements.txt
python 2026_KICS_Fall_10/GN_model/paper1_reproducibility.py
python 2026_KICS_Fall_10/GN_model/figure5_benchmark.py
```

## Performance results

The CI workflow will populate this section after executing the actual repository code. Final metrics are also stored in:
- `results/paper1_reproducibility.json`
- `results/figure5_reproduction.csv`
- `results/paper_gn_egn_comparison.csv`
- `results/summary.json`

## Interpretation

The strongest evidence for this GN code is not a fitted match to one plot, but that the same engine implements the paper's link-function physics, supports a materially wider system scope (full WDM SCI/XCI/MCI, irregular channels, unequal spans, beta2+beta3, heterogeneous gain profiles), and can reproduce a published reach benchmark with no NLI scale-factor fitting.

The final judgment on general-purpose use should be based on the measured errors below and on whether the intended system remains inside GN-model assumptions: uncompensated coherent transmission, low-to-moderate nonlinearity, adequate dispersion/Gaussianization, and a transmitter/receiver spectrum that is represented accurately.
