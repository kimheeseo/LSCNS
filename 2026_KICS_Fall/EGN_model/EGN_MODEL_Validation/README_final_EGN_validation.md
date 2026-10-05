# final_EGN.py Validation Report

## 1. Overview

This directory contains the validation results of `final_EGN.py` against the paper:

- A. Carena et al., **"EGN model of non-linear fiber propagation"**, Optics Express, Vol. 22, No. 13, 2014.
- GN baseline reference:
  - P. Poggiolini et al., **"A Detailed Analytical Derivation of the GN Model of Non-Linear Interference in Coherent Optical Transmission Systems"**, arXiv:1209.0394 v13.

The purpose of this validation is to compare the code results against the paper for:

- **Fig. 1**: SCI
- **Fig. 3**: XCI only, 3-channel system
- **Fig. 6**: XMCI = XCI + MCI, 3-channel system
- **Fig. 8**: XMCI = XCI + MCI, 9-channel system

No fitting, empirical offset, or paper-target tuning was used.

---

## 2. Model Structure

`final_EGN.py` is organized conceptually as:

```text
Poggiolini 1209.0394 GN
        ↓
GN baseline
        │
        ├── SCI EGN correction
        ├── XCI EGN correction
        └── MCI EGN correction

Carena 2014 EGN
        ↑
SCI / XCI / MCI non-Gaussian correction terms
```

The final EGN result is conceptually:

```text
EGN
= GN baseline
+ SCI correction
+ XCI correction
+ MCI correction
```

The code supports model selection such as:

```python
nli_model = "gn"
nli_model = "egn_sci"
nli_model = "egn_full"
```

---

## 3. Paper Conditions Used for Validation

### Common conditions

- Modulation: PM-QPSK
- Symbol rate: 32 GBaud
- Span length: 100 km
- Fiber loss: 0.22 dB/km
- CUT: center channel
- Channel spacing for multi-channel cases:
  - 1.05 × 32 GBd = 33.6 GHz

### Fiber parameters

| Fiber | D [ps/(nm·km)] | gamma [1/(W·km)] | alpha [dB/km] |
|---|---:|---:|---:|
| SMF | 16.7 | 1.3 | 0.22 |
| NZDSF | 3.8 | 1.5 | 0.22 |
| LS | -1.8 | 2.2 | 0.22 |

### Figure-specific setup

| Figure | Case |
|---|---|
| Fig. 1 | Single-channel SCI |
| Fig. 3 | 3 channels, XCI only, SCI removed |
| Fig. 6 | 3 channels, XMCI = XCI + MCI, SCI removed |
| Fig. 8 | 9 channels, XMCI = XCI + MCI, SCI removed |

The paper simulation used a raised-cosine power spectrum with roll-off 0.05.
The analytical EGN equations themselves are based mainly on rectangular-spectrum assumptions.

---

## 4. Validation Method

The paper curves were extracted from the PDF plot vector coordinates rather than OCR.

Comparison points:

```text
Nspan = 1, 2, 5, 10, 20, 50
```

Metrics:

- Absolute error in dB
- Mean absolute error (MAE)
- Maximum absolute error
- Relative error calculated from linear eta

---

## 5. Main Validation Results

### Overall Paper vs Code comparison

| Figure | Comparison | MAE [dB] | Max Error [dB] | Mean Relative Error |
|---|---|---:|---:|---:|
| Fig. 1 | GN vs Paper GN | 0.078 | 0.171 | 1.77% |
| Fig. 1 | EGN vs Paper EGN | 0.144 | 0.846 | 3.45% |
| Fig. 3 | GN vs Paper GN | 0.045 | 0.118 | 1.03% |
| Fig. 3 | EGN vs Paper EGN | 0.331 | 1.491 | 8.41% |
| Fig. 6 | GN vs Paper GN | 0.047 | 0.118 | 1.09% |
| Fig. 6 | EGN vs Paper EGN | 0.296 | 1.374 | 7.45% |
| Fig. 8 | GN vs Paper GN | 0.072 | 0.346 | 1.64% |
| Fig. 8 | EGN vs Paper EGN | 0.462 | 1.313 | 11.58% |

---

## 6. Interpretation

### 6.1 GN baseline

The GN baseline matches the paper very well.

Across Fig. 1, 3, 6, and 8:

```text
Typical MAE ≈ 0.05–0.08 dB
```

This indicates that the Poggiolini-based GN integral path is reproduced with high consistency.

### Assessment

```text
GN baseline: GOOD / VALIDATED
```

It is suitable as a general research baseline for NLI estimation under the GN-model assumptions.

---

### 6.2 SCI-EGN

For Fig. 1, the SCI-EGN result also shows good agreement with the paper.

Overall:

```text
MAE ≈ 0.144 dB
```

At 50 spans, the code and paper EGN values are very close for SMF, NZDSF, and LS.

The larger discrepancy seen in the first few spans is not automatically evidence of a coding failure, because the original paper itself also reports residual differences between EGN and split-step simulation at short distances.

### Assessment

```text
SCI-EGN: GOOD / RESEARCH-GRADE
```

SCI-EGN can be used for modulation-dependent SCI studies within the supported assumptions.

---

## 7. Full-EGN XCI/MCI Issue

The largest remaining issue is not the GN baseline or SCI.

The current limitation is:

```text
XCI/MCI correction integral convergence
```

The XCI/MCI correction terms depend strongly on the numerical quadrature setting.

For example, for SMF, 50 spans, the Full-EGN result changes significantly when `frequency_order` is changed.

Example trend:

| frequency_order | Example Full-EGN eta |
|---:|---:|
| 32 | ~41.76 dB |
| 48 | ~37.20 dB |
| 64 | ~41.66 dB |
| 80 | ~41.08 dB |
| 96 | ~32.92 dB |

The expected paper result is around:

```text
~40.29 dB
```

The result does not converge monotonically.

In some 9-channel cases, certain quadrature orders can even produce an excessively large negative correction, making the total XMCI result non-physical.

---

## 8. Meaning of the Error

The current status should therefore be interpreted as:

```text
GN baseline:
    Accurate

SCI-EGN:
    Accurate enough for research use

XCI-EGN / MCI-EGN:
    Formula structure implemented
    but numerical integration is not yet sufficiently stable
```

This means that the current code should **not** be described as a fully validated production-grade Full-EGN implementation.

---

## 9. Can This Model Be Used for Other Simulations?

### GN baseline

**YES**

Recommended for:

- SMF / G.652.D
- G.654.E
- generic WDM links
- launch power studies
- span-count studies
- NLI baseline comparison
- preliminary capacity studies

provided the GN-model assumptions remain valid.

---

### SCI-EGN

**YES, with conditions**

Recommended for:

- modulation-format comparison
- single-channel nonlinear correction
- SCI analysis
- PM-QPSK / PM-16QAM / PM-64QAM style studies

Recommended scope:

- coherent accumulation
- uncompensated links
- rectangular or nearly rectangular spectrum
- homogeneous EDFA links preferred

---

### Full EGN: XCI + MCI

**NOT YET recommended for quantitative design sign-off**

It can currently be used for:

- algorithm development
- sensitivity analysis
- qualitative comparison
- paper reproduction work

It should not yet be used as the final numerical reference for:

- new fiber certification
- absolute NLI prediction
- system design sign-off
- customer specification guarantee
- low-dispersion fibers without additional SSFM validation

---

## 10. Recommended Next Improvements

Priority order:

### 1. Re-check Appendix A/B/C implementation

Re-verify:

- XCI X1–X4 terms
- MCI M0–M3 terms
- normalization factors
- power convention
- conjugate terms
- integration boundaries

### 2. Improve oscillatory integral evaluation

Current fixed Gauss-Legendre quadrature is not sufficiently robust for every XCI/MCI correction term.

Recommended approaches:

- adaptive quadrature
- phase-aware interval splitting
- oscillatory quadrature
- analytical variable transformation
- convergence-controlled integration

### 3. Add convergence criteria

The final result should not be accepted until:

```text
|result(order_high) - result(order_low)| < threshold
```

for example:

```text
< 0.05 dB
```

### 4. Re-run paper validation

Repeat:

- Fig. 3
- Fig. 6
- Fig. 8

for all:

- SMF
- NZDSF
- LS

### 5. Add independent SSFM validation

After numerical convergence is achieved:

```text
Paper
vs
final_EGN.py
vs
SSFM
```

should be compared independently.

---

## 11. Current Model Status

```text
GN baseline          : VALIDATED
SCI-EGN              : VALIDATED / RESEARCH-GRADE
XCI-EGN              : IMPLEMENTED, NUMERICAL CONVERGENCE NEEDS IMPROVEMENT
MCI-EGN              : IMPLEMENTED, NUMERICAL CONVERGENCE NEEDS IMPROVEMENT
Full-EGN             : NOT YET PRODUCTION-GRADE
```

---

## 12. Related Files

Recommended directory contents:

```text
final_EGN.py

final_EGN_Carena_Fig1_3_6_8_Validation_Colab.ipynb

paper_digitized_fig1_3_6_8.csv
final_EGN_code_fig1_3_6_8.csv

final_EGN_validation_summary.csv
final_EGN_paper_vs_code_detail.csv
final_EGN_validation_overall.csv

final_EGN_quadrature_sensitivity_SMF_50span.csv

graphs/
    Fig1_SMF_paper_vs_code.png
    Fig1_NZDSF_paper_vs_code.png
    Fig1_LS_paper_vs_code.png
    Fig3_*.png
    Fig6_*.png
    Fig8_*.png
    quadrature_sensitivity_SMF_50span.png
```

---

## 13. Recommended GitHub Folder Name

### Recommended

```text
2026_KICS_Fall_10/EGN_FullModel_Validation
```

This name clearly indicates that the folder contains both:

- the Full-EGN model implementation
- its validation results

### Shorter alternative

```text
2026_KICS_Fall_10/final_EGN_validation
```

### If the folder will later include SSFM

A more future-proof name is:

```text
2026_KICS_Fall_10/EGN_Model_Validation
```

This is the preferred choice if the directory will later contain:

```text
GN
EGN
SSFM
Paper comparison
```

---

## 14. Recommended Repository Structure

```text
2026_KICS_Fall_10/
└── EGN_Model_Validation/
    ├── README.md
    ├── final_EGN.py
    ├── final_EGN_Carena_Fig1_3_6_8_Validation_Colab.ipynb
    │
    ├── data/
    │   ├── paper_digitized_fig1_3_6_8.csv
    │   ├── final_EGN_code_fig1_3_6_8.csv
    │   ├── final_EGN_validation_summary.csv
    │   ├── final_EGN_paper_vs_code_detail.csv
    │   └── final_EGN_quadrature_sensitivity_SMF_50span.csv
    │
    └── graphs/
        ├── Fig1_SMF_paper_vs_code.png
        ├── Fig1_NZDSF_paper_vs_code.png
        ├── Fig1_LS_paper_vs_code.png
        ├── Fig3_*.png
        ├── Fig6_*.png
        ├── Fig8_*.png
        └── quadrature_sensitivity_SMF_50span.png
```

---

## 15. Bottom Line

The current implementation successfully reproduces the GN baseline and SCI-EGN with high accuracy.

The main remaining technical issue is:

```text
numerical convergence of XCI/MCI correction integrals
```

Therefore the model should currently be positioned as:

```text
Research-grade GN + SCI-EGN model
with an experimental Full-EGN XCI/MCI implementation
under active numerical validation.
```

This wording is recommended for GitHub until the XCI/MCI convergence issue is resolved.
