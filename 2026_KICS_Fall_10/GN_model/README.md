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
The full precision path in `EGN_adaptive.py` uses rectangular spectra and coherent multi-span physics, while Carena-2012 Fig. 5 is a **GN** benchmark plotted with incoherent NLI accumulation. A direct EGN-vs-Fig.5 accuracy claim would therefore mix model scopes.

For the requested three-way graph, the representative **50-GHz QPSK points for PSCF, SMF and NZDSF** use a transparent paper-convention proxy: `EGN_adaptive.py` computes the full **one-span** EGN NLI, and that one-span NLI is accumulated incoherently across spans exactly like the Fig. 5 GN plotting convention. No fitted scale factor is used. The resulting EGN error is reported, but is explicitly **not** interpreted as a validation score for the native multi-span EGN algorithm.

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

GitHub Actions run #15 completed successfully after executing the repository code, the paper-1 tests, the Fig. 5 benchmark, and the executed Colab report.

### Paper 1 — equation-level reproducibility

- Assessment: **PASS**
- Maximum equation-level relative error: **1.29452 × 10⁻¹¹ %**
- Scrambled-Sobol relative standard deviation at 32,768 samples: **0.0935 %**
- Change from 8,192 to 32,768 samples: **0.1441 %**

These values show that the implementation reproduces the tested GN equations to floating-point precision and that the numerical WDM integral is stable at the selected high-resolution setting.

### Paper 2 — Carena 2012 Fig. 5

Across **63 digitized Fig. 5 points**:

- GN MAPE: **8.66 %**
- Median absolute percentage error: **6.67 %**
- Maximum point error: **50.0 %**
- Points within 5 %: **36.5 %**
- Points within 10 %: **69.8 %**
- Points within 15 %: **85.7 %**
- Points within 20 %: **92.1 %**

For points with paper reach ≥ 1,000 km (**47 points**), GN MAPE improves to **7.38 %**, with a maximum error of **19.08 %**. The largest percentage errors occur mainly at very short-reach points where a 100-km span step and PDF digitization have a large percentage effect.

Fiber-level MAPE:

| Fiber | Points | GN MAPE |
|---|---:|---:|
| PSCF | 21 | 9.58 % |
| SMF | 21 | 6.95 % |
| NZDSF | 21 | 9.43 % |

Modulation-level MAPE:

| Modulation | Points | GN MAPE |
|---|---:|---:|
| BPSK | 18 | 9.31 % |
| QPSK | 18 | 5.20 % |
| 8QAM | 15 | 10.11 % |
| 16QAM | 12 | 11.03 % |

### Paper vs GN vs EGN representative comparison

For the representative **50-GHz QPSK** PSCF / SMF / NZDSF subset:

| Fiber | Paper reach | GN reach | GN error | EGN proxy reach | EGN proxy error |
|---|---:|---:|---:|---:|---:|
| PSCF | 8,900 km | 10,200 km | 14.61 % | 14,900 km | 67.42 % |
| SMF | 4,100 km | 4,400 km | 7.32 % | 6,500 km | 58.54 % |
| NZDSF | 2,700 km | 2,800 km | 3.70 % | 4,000 km | 48.15 % |

Same-subset GN MAPE is **8.54 %**. The EGN proxy MAPE is **58.03 %**.

**Important:** the EGN number above is deliberately not treated as a native EGN validation score. Carena-2012 Fig. 5 is a GN benchmark with incoherent span accumulation, while `EGN_adaptive.py` is designed for rectangular spectra and coherent multi-span EGN physics. The three-way graph therefore uses one-span full-EGN NLI with the paper's incoherent accumulation convention only to provide the requested side-by-side comparison. A proper EGN validation must use a paper/SSFM benchmark with matching EGN assumptions.

Final machine-readable outputs are stored in:
- `results/paper1_reproducibility.json`
- `results/figure5_reproduction.csv`
- `results/paper_gn_egn_comparison.csv`
- `results/summary.json`
- `results/COLAB_EXECUTION_REPORT.md`
- `results/GN_Model_Colab_executed.ipynb`

## Interpretation and recommendation for general use

The strongest evidence for this GN code is not a fitted match to one plot. The same engine reproduces the tested GN link-function equations to floating-point precision, shows sub-0.1 % high-resolution Sobol run-to-run dispersion, and reproduces a published 63-point reach benchmark with **8.66 % overall MAPE without an NLI fitting factor**.

For engineering studies, the model is therefore well suited as a **general-purpose GN screening / design engine** when the system is within standard GN assumptions: uncompensated coherent transmission, sufficient accumulated dispersion / Gaussianization, low-to-moderate nonlinear regime, and accurately represented Tx/Rx spectra. Its support for full-WDM SCI/XCI/MCI, irregular channel plans, unequal spans, beta2+beta3, and heterogeneous span parameters makes it more broadly reusable than a Fig.-5-specific reproduction script.

It should not be presented as universally <5 % accurate. The present independent Fig. 5 validation supports roughly **7–9 % typical reach error** over the tested domain, with QPSK performing best (~5.2 % MAPE) and short-reach / high-order-modulation points showing larger percentage deviations. For publication or customer-facing high-accuracy claims, use this GN model together with matched SSFM/EGN validation for the exact fiber, spectrum, modulation and span regime.
