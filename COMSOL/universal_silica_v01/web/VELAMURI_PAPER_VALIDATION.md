# Velamuri et al. (2018) — Bending-loss paper implementation review

**Date:** 2026-11-03  
**Reference:** Ashitosh Velamuri et al., *Study of Bending Losses in Optical Fibers using COMSOL*, COMSOL Conference 2018.  
**Official abstract:** https://www.comsol.com/paper/study-of-bending-losses-in-optical-fibers-using-comsol-multiphysics-61111  
**Paper PDF:** https://cn.comsol.com/paper/download/579211/velamuri_paper.pdf

## 1. Current workbench fidelity

| Paper building block | Current browser functionality | Validation / limitation |
|---|---|---|
| Bent fiber geometrical conformal mapping | **Implemented**, `n_geom(x)=n_0(x)(1+x/R)` | Profile plot; only first-order `x ≪ R` approximation |
| Stress-optic equation with P11/P12 | **Implemented**, for explicitly supplied/proxy strains | P11=0.113, P12=0.252. Current simple beam/Poisson proxy **IS NOT GEBT** |
| Geometrically Exact Beam Theory (GEBT) for actual strain field | **Not implemented** | Requires mechanical/structural model and fiber-coating properties |
| 2D full-vector complex mode FEM | **Not implemented** | Existing LP01 solver is real-valued, scalar and radial |
| PML absorbing boundary and 7λ study | **External result analyzer only** | 7λ is paper's model sensitivity finding, not computed PML here |
| Loss from imaginary effective index | **Implemented postprocessor** | Requires `Im(n_eff)` supplied by an external *complex* mode solver |
| Paper Fig. 4 bending loss versus diameter | **Not independently reproduced** | Missing paper's actual detailed RIP, mechanical field and FEM/PML |
| Comparison to actual G.654.E manufacturer data | **Not validated** | Paper's G.652/BIF example is not YOFC or PremierLink LE125 |

## 2. Units and equations

The **input is bend diameter** `d` in mm; corresponding radius is `R=d/2` in mm. This distinction is essential: a 10 mm **diameter** bend is a 5 mm **radius** bend.

For a bent mode with the attenuation-sign convention `E(z)=E0 exp(iβz)`, `β=k0 n_eff`, the loss due to positive `|Im(n_eff)|` is

```
alpha_dB/m    = (20 / ln(10)) * (2*pi / lambda_m) * abs(Im(n_eff))
L_turn_m      = pi * diameter_m
alpha_dB/turn = alpha_dB/m * L_turn_m
```

The equivalent-straight index is `n_equiv = n_stress (1+x/R)`. The stress-optic law from the paper is
`n_stress = n [1 - n²/2 (P11*epsilon1 + P12*(epsilon2+epsilon3))]`.
The plotted optional stress proxy uses `epsilon1=x/R` and `epsilon2=epsilon3=-nu*x/R`, **not GEBT**, to illustrate sensitivity, not certify accuracy.

## 3. Actual code execution: numerical postprocessing self-check

| Value | Result |
|---|---:|
| Reference bend diameter | 10 mm |
| Reference wavelength | 1550 nm |
| Reference BIF simulated loss (paper) | 0.0119 dB/turn |
| Reference measured mean (paper) | 0.012 dB/turn |
| `abs(Im(n_eff))` **inverted** from reference loss | **1.0758076087524702e-8** |
| Recomputed loss with same formula | **0.011900000000000003 dB/turn** |
| Three turns using constant imported Im(neff) | **0.03570000000000001 dB** |
| Paper's stated illustrative PML thickness | 7λ = **10.85 µm** @1550 nm |

**Interpretation:** the loss is reconstructed **because the complex index magnitude was back-calculated from the paper's loss**. This validates dimensional conversion and the code, **not** a predictive curved-fiber eigenmode solver. The output is NOT a newly reproduced COMSOL result. Published simulated-versus-measured agreement cannot be transferred to any specific fiber without calibrated inputs.

## 4. External-PML diagnostic smoke test (synthetic, not paper data)

Synthetic external `Im(n_eff)` samples (5λ: 1.20e-8; 7λ: 1.18e-8) convert into:

- 5λ: 0.01327374884117 dB/turn
- 7λ: 0.01305251969382 dB/turn
- Last-two relative difference: **1.694915254%**, which meets the tool's *user-defined* 2% demonstration threshold.

A second synthetic sample (5λ: 1.50e-8, 7λ: 1.10e-8) correctly fails. **Neither example is an actual PML study**, and no observed convergence of the PML/mesh follows from them. Do not present the 2% threshold as a published criterion.

## 5. How to use this upgrade

Open [Universal Silica Fiber Lab](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/?v=20261103-velamuri-v1) and click **02+ · 굽힘 물리식 / COMSOL 후처리**.

1. Set **diameter**, wavelength and optional stress proxy. The plots read your current Step/Graded/Trench/Multilayer core settings.
2. Obtain the complex guided-mode `n_eff` from a **real 2D mode eigensolver with PML**; enter the **magnitude of its imaginary component**.
3. Verify PML- and mesh-thickness convergence with **independent FEM runs**; paste optional PML thickness-vs-Im(n_eff) pairs into the diagnostic field.
4. Read `dB/m`, `dB/turn`, number-of-turn cumulative loss and save a provenance-tagged JSON.
5. To check the arithmetic only, select the **논문 BIF 결과 역산 예시** button.

## 6. Roadmap to actual physical bend loss simulation

Build and validate a new independent Python FEM module, not by mislabeling the current scalar browser engine:

1. Calibrated measured core/central-dip/trench index profile for a particular product.
2. Bending stress/strain profile from GEBT or a validated mechanics model, then stress-optic law and geometric mapping.
3. 2D vector complex eigenmode problem with PML radiation boundary; report complex `n_eff` and field.
4. PML thickness 1λ…7λ+ and mesh refinement sweeps, assess fitted `Im(n_eff)` stability and PML/backreflection.
5. Comparative tests against the paper's G.652/BIF cases and matched IEC 60793-1-47 bent-fiber measurements.
6. Only after validation, expose **Predictive Physical FEM** as a distinct UI mode.

**Current status: conformal-index visualization and external-complex-mode postprocessing complete; predictive 2D FEM/PML and GEBT incomplete.**
