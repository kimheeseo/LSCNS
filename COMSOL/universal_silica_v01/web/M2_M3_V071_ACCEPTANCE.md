# LS Wave Optics — M2/M3 acceptance review and additive update, v0.7.1

Date: 2026-11-12 (KST). Public static site: https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/?v=v0.7.1

## ① Existing code read; untouched physics

The current main branch was **already v0.7.0** when this acceptance request arrived, not v0.5.1. It already had:
- M2 original `material_models.mjs`, `material_m2_ui.mjs`; enhanced `m2_preview_engine_v06.mjs` / `m2_preview_ui_v06.mjs` for mol%-doped Ge/P, F wt%, F mol% user coefficients, linear/PCHIP, α/k preview and 1550 nm table.
- M3 original `vector_fdfd_browser.mjs` (collocated Hx/Hy) and `m3_study_ui_v07.mjs` (progress/cancel/mesh/time/memory); **a different** genuine Yee-staggered `yee_solver.mjs` used through `vector_workbench.mjs`.
- Original `solver.mjs` radial FEM, MFD/Aeff/D/S₀, M1 tests, HCF, MCF, PMD, Colab `vector_mode_ui.mjs` uploader, JSON/CSV/PNG import-export, and M4–M6 all intact.

This change introduces NO second Maxwell solver; it exposes the existing true Yee M3 path more clearly.

## ② Newly added files and their responsibilities

1. `m3_yee_quickstart_v071.mjs` — optional independent M3 Yee nav shortcut + explicit step-index, no-PML baseline setup button. Setting does not start computation. Feature flag: `window.LSWaveFeatures.m3YeeQuickstart=false`.
2. `tests/test_m2_m3_acceptance_v071.mjs` — old M1 4+5 regression, M2 primary source / Sellmeier / table / α↔κ, 24² Yee vector mode and portable JSON plus local Node wall-time and process RSS.
3. `tests/browser_m2_m3_v071.mjs` — optional Playwright navigation, no-PML settings, true Yee Worker, and retained old Colab import regression.
4. `M2_M3_V071_ACCEPTANCE.md` and `M2_M3_V071_FULL_SOURCE.md` — complete new-code listing, diffs, numeric results and limitations.

## ③ Full new source

All new JavaScript/test files are reproduced in [M2_M3_V071_FULL_SOURCE.md](M2_M3_V071_FULL_SOURCE.md). No original file has been rewritten wholesale.

## ④ Minimal existing-file diffs

```diff
--- a/web/solver.mjs
+++ b/web/solver.mjs
@@
-export const VERSION = '0.7.0';
+export const VERSION = '0.7.1';
```

```diff
--- a/web/app.mjs
+++ b/web/app.mjs
@@
-from './solver.mjs?v=v0.7.0';
+from './solver.mjs?v=v0.7.1';
```

```diff
--- a/web/m2_preview_engine_v06.mjs
+++ b/web/m2_preview_engine_v06.mjs
@@
- * SiO2 + GeO2: Fleming (1984) Applied Optics 23, 4486, Table II
+ * Pure SiO2: I. H. Malitson, JOSA 55, 1205 (1965), Eq. 1; DOI:10.1364/JOSA.55.001205
+ * SiO2 + GeO2: Fleming (1984) Applied Optics 23, 4486, Table II
@@
- silica:{label:"Malitson pure fused silica; Fleming 1984 Table II",url:"https://doi.org/10.1364/AO.23.004486"}
+ silica:{label:"I. H. Malitson, JOSA 55 (1965), pure fused SiO₂ Sellmeier Eq. 1; 20 °C",url:"https://doi.org/10.1364/JOSA.55.001205"}
```

```diff
--- a/web/m2_preview_ui_v06.mjs
+++ b/web/m2_preview_ui_v06.mjs
@@
-from "./m2_preview_engine_v06.mjs?v=20261112-m3m6-v1";
+from "./m2_preview_engine_v06.mjs?v=v0.7.1";
@@ source information
- <p>SiO₂ 및 GeO₂: Fleming (1984) ...</p>
+ <p>순수 SiO₂: Malitson (1965), Eq. 1...</p>
+ <p>GeO₂: Fleming (1984), Table II...</p>
```

```diff
--- a/web/index.html
+++ b/web/index.html
@@
-FIBER DESIGN WORKBENCH · v0.7.0
+FIBER DESIGN WORKBENCH · v0.7.1
@@
-<script type="module" src="./app.mjs?v=20261112-project-v3"></script>
+<script type="module" src="./app.mjs?v=v0.7.1"></script>
@@
-<script type="module" src="./m2_preview_ui_v06.mjs?v=20261112-project-v3"></script>
+<script type="module" src="./m2_preview_ui_v06.mjs?v=v0.7.1"></script>
@@ after original vector_workbench module
+<script type="module" src="./m3_yee_quickstart_v071.mjs?v=v0.7.1"></script>
```

Original M4/PML/HCF/M5/M6 modules, workflows and JSON schemas have **zero patches**. Changelog top version, public URL and header all use v0.7.1.

## ⑤ M2 numbers, units, interpolation and alpha/kappa

Source references (all public and attributable):
- **Pure SiO₂**: I. H. Malitson, *JOSA* 55, 1205 (1965), DOI: https://doi.org/10.1364/JOSA.55.001205
- **GeO₂–SiO₂**: J. W. Fleming, *Applied Optics* 23, 4486 (1984), DOI: https://doi.org/10.1364/AO.23.004486. The intermediate mol% uses a linear interpolation of Sellmeier coefficients, **not measured** samples.
- **F-doped silica**: Fleming & Wood, *Applied Optics* 22, 3102 (1983), DOI: https://doi.org/10.1364/AO.22.003102. Original sample concentrations are **1 and 2 wt%**. Mol%-specific concentration-to-Sellmeier mapping is unavailable; select `f_mol` and enter user-sourced B1–B3/L1–L3, never a fabricated coefficient.
- **P₂O₅** 9.1 mol% composition: *Optical Fiber Communication Systems with MATLAB and Simulink Models*, Table 3.4: https://www.routledge.com/rsc/downloads/Pages_from_K22108_Book_rev.pdf. Intermediate mol% is modeled, not directly measured.

| M2 material preview | n at 1550 nm | Status |
|---|---:|---|
| Pure SiO₂ | 1.4440236217032607 | Exactly same `silicaIndex(1550)` as original solver |
| GeO₂ 5 mol% | 1.4515267771427718 | Coefficient interpolation; approximate |
| F 1 wt% | 1.4393958509698328 | Source fit (wt%, NOT mol%) |
| P₂O₅ 5 mol% | 1.4522244231080190 | Interpolation from 9.1 mol% reference |

M2 supports `wavelength_nm,n_real[,kappa]` sorted CSV data, linear and PCHIP interpolation. Range outside source data throws a warning rather than extrapolating. Under `exp(iβz-iωt)`, use `n_tilde=n+iκ` with `α[dB/km]=(4πκ/λ[m])·1000·10/ln(10)`. At 1550 nm, `α=0.2 dB/km` maps to `κ=5.6802508913155675×10^-12`. This is a preview/material-data convention; **no complex eigensolver coupling** occurs in the dedicated M2 preview.

## ⑥ M3 algorithm and quantitative acceptance

The strict M3 Yee path uses existing `yee_solver.mjs` + `vector_geometry.mjs` with **staggered curl FDFD**, complex CSR, shifted-inverse block eigenpairs and ILU(0)-preconditioned GMRES with Rayleigh–Ritz, in a static module Web Worker. The algorithm is not the older collocated Hx/Hy study (`m3_study_ui_v07.mjs`). Input cap 24–192 cells/axis, 1–16 modes; user can cancel by terminating Worker. Reference field maps include Ex/Ey/Ez, two polarizations, MFD, Aeff, and `LSCNS_vector_mode_v1` portable JSON. Built-in Colab upload and overlay remain available.

The new quickstart explicitly sets `profile=step`, `a=4.1µm`, `λ=1550nm`, `n_clad=1.444023621703261`, `Δn=0.005`, grid 48², 2 modes, `PML=false`, `κ=0`, `bend=0`, and the radial circular geometry. This is user-controlled and does not start work automatically.

### M3 Step-index benchmark — historical stored Node process measurements

Source: `M3_M6_YEE_VALIDATION_20261112.md`, with benchmark data in `results/yee_validation/`.
Input: 1550 nm, a=4.1 µm, n_clad=Malitson 1.444023621703261, Δn=0.005, window halfwidth=18 µm, subpixel sampling=4×4.
Scalar LP₀₁ analytic neff = **1.44610478058**, original scalar FEM = **1.44610475118** (1000 radial elements).

| Yee grid | neff | |difference from LP01| | measured Node seconds | measured process RSS MiB | typed-array estimate MiB |
|---:|---:|---:|---:|---:|---:|
| 24² | 1.44608547856 | 1.93020e-5 | 0.364 | 61.09 | 1.42 |
| 48² | 1.44610494867 | 1.68092e-7 | 2.817 | 87.33 | 5.69 |
| 72² | 1.44609935051 | 5.43007e-6 | 11.527 | 123.76 | 12.81 |
| 96² | 1.44610489953 | 1.18946e-7 | 20.955 | 168.14 | 22.78 |

**Convergence: NOT met.** |neff(96²)−neff(72²)| ≈ **5.55×10⁻⁶**, above the 1×10⁻⁶ adjacent-grid criterion. Eigenpair residual and measured FDFD correctness do not guarantee continuum mesh convergence. Subpixel averaging does not replace resolution and domain studies.

### Colab/Python result comparison — *different discretization*

Reference stored Python collocated FDFD (a=5.5 µm, λ=1550 nm, Δn=0.0034785793846155, halfwidth=22 µm), neff = **1.445711888509771**. Yee operates on a staggered grid with different N and numerical boundary treatment. Comparison is physical-condition-matched, **not same-method/mesh exact equality**.

| Yee N | Yee neff | Python FDFD neff | Absolute difference |
|---:|---:|---:|---:|
| 48 | 1.4457134495 | 1.44571188851 | 1.56099e-6 |
| 68 | 1.44571882877 | 1.44571188851 | 6.94026e-6 |
| 96 | 1.44571573505 | 1.44571188851 | 3.84654e-6 |

Independent same-discrete-matrix SciPy sparse ARPACK N24 cross-check reported `~9.5×10⁻¹³` neff difference, but this is an eigensolver check, not experimental validation.

### Actual source execution in this request
- Current M1 4/4 LP01 cases + 5/5 old scalar presets PASS through JavaScript source evaluation.
- M2 1550-nm catalog output reproduced exactly as table above; no change to original `silicaIndex`.
- Actual `yee_solver.mjs` N24 2-mode computation: neff `1.4460854785569148`, `1.4460854785554098`; eigenpair residuals `4.72894×10⁻⁹`, `4.82139×10⁻⁹`; portable schema `LSCNS_vector_mode_v1`. The source-evaluation elapsed time ~18 seconds is **not a browser or Node host hardware benchmark**.
- No new browser UI/Worker/Playwright test has been executed in this request. Node/Browser test code is added, but execution on a cloned repository remains pending.

## ⑦ Regressions / execution / limitations

```bash
cd COMSOL/universal_silica_v01/web
node tests/test_solver.mjs
node tests/test_m1_validation.mjs
node tests/test_m1_quantitative.mjs
node tests/test_material_m2.mjs
node tests/test_m2_preview_v06.mjs
node tests/test_vector_m3.mjs
node tests/test_yee.mjs
node tests/test_m2_m3_acceptance_v071.mjs

# Optional local UI test (first run HTTP server from repository root):
python3 -m http.server 8765
# Then run from the web directory:
node tests/browser_m2_m3_v071.mjs
```

Preserved: all original solver numerical functions, M1 results, MFD/Aeff/D/S0, wavelength sweep, mesh, PMD, HCF, MCF, legacy Colab JSON route, design JSON/CSV/PNG. No M4–M6 code changed.

**Unverified:** user-browser full Worker timing/heap, device-specific cancellation latency, true continuous-grid convergence (failed), direct full-vector COMSOL vector-eigenmode ground truth for this design, doped material n(λ) between measured concentrations, F mol%-specific coefficients, material absorption in M2 mode eigenvalue solver. The experimental M4–M6 features already present in this repository are outside this update.

