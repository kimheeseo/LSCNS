# LS Wave Optics v0.7.0 — M2/M3 Update & Verification

2026-11-12 · Static GitHub Pages · M2 v0.6.0 and M3 v0.7.0 only · No M4–M6.

## ① Existing code read and minimal patches

- web/solver.mjs: original scalar P1 radial FEM, MFD, Aeff, D, S0, sweep and mesh functions **unchanged**; VERSION metadata only 0.5.1 → 0.7.0.
- web/app.mjs: only scalar solver import cache query changed; existing DOM IDs, save/load and JSON/CSV/PNG exports **unchanged**.
- web/material_models.mjs, material_m2_ui.mjs: only original F concentration provenance/labels clarified to **wt%**; zero numerical changes.
- web/vector_fdfd_browser.mjs and vector_mode_ui.mjs: full-vector operator and legacy Colab JSON viewer **unchanged**; new M3 study wraps them.
- web/index.html: 4 new CSS/JS assets, header version, original app query and the old material UI cache tag; no original UI section removed.
- M1, PMD, HCF, MCF, Geometry, PML, FEM solver, design JSON structure **untouched**.

## ② New files and responsibilities

| Milestone | File | Purpose |
|---|---|---|
| M2 | m2_preview_engine_v06.mjs | Sourced Sellmeier, F wt/mol guard, n(λ) linear/PCHIP, α↔κ, JSON v2 migration |
| M2 | m2_preview_ui_v06.mjs | Core/Cladding selectors, charts n/k, 1550 nm values, provenance, user tables |
| M2 | m2_preview_ui_v06.css | Only new M2 styles |
| M2 | tests/test_m2_preview_v06.mjs | Material/reference and original-FEM regression |
| M3 | m3_study_metrics_v07.mjs | Grid caps, convergence checks, LP01/scalar/Colab comparison |
| M3 | m3_study_worker_v07.mjs | Web Worker, progress events, measured elapsed time, multi-grid computation |
| M3 | m3_study_ui_v07.mjs | Cancel, comparison table, maps, overlay, two JSON exports |
| M3 | m3_study_ui_v07.css | Only new M3 styles |
| M3 | tests/test_m3_study_v07.mjs | Numerical/same-grid Python regression |
| M3 | tests/benchmark_m3_v07.mjs | Optional local Node RSS, heap, ArrayBuffer and elapsed measurements |
| Both | tests/browser_m2_m3_v07.mjs | Optional Playwright browser smoke and cancel test |

## ③ Complete new source

All 11 files are reproduced in full in [M2_M3_V07_FULL_SOURCE.md](M2_M3_V07_FULL_SOURCE.md).

## ④ Minimal existing-file diffs

~~~diff
--- a/web/solver.mjs
+++ b/web/solver.mjs
@@
-export const VERSION = '0.5.1';
+export const VERSION = '0.7.0';
~~~

~~~diff
--- a/web/app.mjs
+++ b/web/app.mjs
@@
-from './solver.mjs?v=v0.5.1';
+from './solver.mjs?v=v0.7.0';
~~~

~~~diff
--- a/web/index.html
+++ b/web/index.html
@@
+  <link rel="stylesheet" href="./m2_preview_ui_v06.css?v=v0.7.0" />
+  <link rel="stylesheet" href="./m3_study_ui_v07.css?v=v0.7.0" />
@@
-FIBER DESIGN WORKBENCH · v0.5.1
+FIBER DESIGN WORKBENCH · v0.7.0
@@
-  <script type="module" src="./app.mjs?v=v0.5.1"></script>
+  <script type="module" src="./app.mjs?v=v0.7.0"></script>
@@
+  <script type="module" src="./m2_preview_ui_v06.mjs?v=v0.7.0"></script>
+  <script type="module" src="./m3_study_ui_v07.mjs?v=v0.7.0"></script>
~~~

Existing material UI concentration label changed from “mol%” to “Ge/P mol%; F wt%”, plus two corresponding provenance strings in material_models.mjs. This is a **text-only** correction, leaving all coefficients and outputs identical.

Latest banner/version: **v0.7.0**, public URL query **?v=v0.7.0**, final source metadata **0.7.0**, latest CHANGELOG entry **v0.7.0**. Old unchanged assets retain historical per-module cache tags.

## ⑤ M2 physical references and numerical outputs

- **Pure SiO₂:** legacy Malitson Sellmeier, exact reused silicaIndex function.
- **GeO₂:** Fleming 1984 Applied Optics 23, 4486, Table II (pure silica/pure germania). Coefficient mixing with mole fraction is a **model assumption**, not data measured at every intermediate concentration. https://doi.org/10.1364/AO.23.004486
- **F:** Fleming and Wood 1983 Applied Optics 22, 3102, Table II: **1 wt% and 2 wt%**, NOT mol%. Values must not be labeled as 1/2 mol%. A new F mol% selection requires user-supplied 3 B coefficients, 3 resonant wavelengths L, and a cited source; no unverified wt→mol conversion. https://doi.org/10.1364/AO.22.003102
- **P₂O₅:** published 9.1 mol% reference coefficients in Optical Fiber Communication Systems with MATLAB and Simulink Models, Table 3.4; intermediate concentrations use explicitly modeled coefficient mixing. https://www.routledge.com/rsc/downloads/Pages_from_K22108_Book_rev.pdf

| Material sample | n(1550 nm) | Grounding |
|---|---:|---|
| SiO₂ pure | 1.4440236217032607 | exactly legacy Sellmeier |
| GeO₂ 5 mol% | 1.4515267771427718 | modeled, not measured at 5% |
| F 1 **wt%** | 1.4393958509698328 | measured-sample Sellmeier; **not 1 mol%** |
| P₂O₅ 5 mol% | 1.4522244231080190 | modeled from 9.1 mol% endpoint |

CSV user n(λ) input: wavelength_nm,n_real[,kappa], sorted 2 or more rows, either linear or monotonic cubic Hermite (PCHIP). **Extrapolation forbidden** and out-of-range requests warn instead of generating invented values.

Complex index convention: ñ=n+iκ, E(z,t)=E0 exp[i(2π/λ)(n+iκ)z−iωt]. Intensity decays exp(−4πκz/λ). The conversion is

α[dB/km] = 4πκ/(λ in metres) ×1000×10/ln(10)

κ = α[dB/km] × λ[m] × ln(10)/(4π×1000×10)

At 1550 nm: α=0.2 dB/km ↔ κ=5.6802508913155675×10⁻¹². **No complex mode solver/absorption coupling**, preview/data only.

Separate M2 preview JSON kind=LSWaveM2MaterialPreview, schemaVersion=2. Schema 1 migration detects ambiguous old F concentration and records a manual-review state; existing design JSON is unchanged.

## ⑥ M3 numerical method, measured values and assumptions

The existing full-vector operator, originally ported from full_vector/vector_modes.py, is a coupled **collocated** Hx/Hy 2D FDFD, with central derivatives and matrix-free shifted-inverse iteration. Inner diagonal-preconditioned BiCGSTAB relative residual=2×10⁻⁹, eigenpair relative residual threshold 1×10⁻⁷, two orthogonalized polarizations.

**Acceptance gap:** it is NOT a Yee-staggered FDFD grid and NOT vector edge-element FEM. The strict requested choice (Yee or vector FEM) is **NOT fully met**. The new M3 adds real worker-based computation, progress, cancel, numerical tables, portable JSON and comparison; it does NOT mislabel the old collocated numerical operator.

- Grid odd 29–91 (cap), 1–5 mesh choices. Web Worker ensures UI does not block; cancel uses Worker.terminate and discards incomplete data.
- Ex/Ey/Ez maps, neff, MFD, Aeff and polarization; portable JSON keeps schema=LSCNS_vector_mode_v1; mode 2 gains additive field maps.
- Existing Colab viewer and uploader retained. Overlay is shown only if geometry/material conditions AND grid match.
- New study JSON kind=LSWaveFullVectorStudy, schemaVersion=2. No changes to older JSON format.

**Step-index numerical fixture:** λ=1550 nm, a=5.5 µm, n_clad=1.444023621703261, Δn=0.0034785793846155, window halfwidth=22 µm.
- M1 scalar infinite-cladding LP₀₁: **1.445718990544404**
- Existing scalar P1 FEM: **1.4457189679091202**
- Stored Python/Colab at 69²: **1.445711888509771**
- New Worker at 69²: **1.445711888493530** (2nd polarization 1.445711888493495)
- Same-grid JS–Python neff delta: **1.62410085×10⁻¹¹**.
- Worker discrete eigenpair residual: **9.77425466×10⁻⁸**.
- Worker vector vs LP₀₁ difference: **7.10205087×10⁻⁶**; versus scalar FEM **7.07941559×10⁻⁶**.

### Grid-sweep source execution measurements (JS V8 reference environment)

| Grid N | Vector mode1 neff | |Δneff−LP₀₁| | V8 wall seconds (both polarizations) |
|---:|---:|---:|---:|
| 29 | 1.445708584358583 | 1.04062×10⁻⁵ | 1.276 |
| 41 | 1.445651096191941 | 6.78944×10⁻⁵ | 4.088 |
| 51 | 1.445709266998758 | 9.72355×10⁻⁶ | 6.611 |
| 69 | 1.445711888493530 | 7.10205×10⁻⁶ | 14.816 |
| 81 | 1.445701713505157 | 1.72770×10⁻⁵ | 22.835 |

**Non-monotonic due staircased core boundary**: the UI labels grid pairs as **미수렴** when |Δneff| > 1×10⁻⁶, and does not claim continuum-vector convergence from a small eigen residual.

**Time distinction:** V8 wall times above were actually measured in the source evaluation environment, not on the user's browser or CPU. A separately executed 69×69 Worker code path took 14.963 s in V8. Browser UI reports live performance.now() times.

**Memory distinction:** No RSS or heap measurement was available in the source evaluator, so actual memory is **미측정**. Worker UI checks performance.memory if available, otherwise explicitly prints **실측 미지원**. At N=69, a *Float64 typed-array model estimate* is ≈1.498 MiB, not measured RAM and excludes JS objects. The optional Node benchmark prints actual local process RSS/heap/arrayBuffers.

M3 does not include Yee/FEM, PML, HCF, bending/leakage loss, complex neff, material κ in eigenproblem or arbitrary 2D mask. Those remain unsupported.

## ⑦ Reproducibility and regression tests

~~~bash
cd COMSOL/universal_silica_v01/web
node tests/test_solver.mjs
node tests/test_m1_validation.mjs
node tests/test_m1_quantitative.mjs
node tests/test_material_m2.mjs
node tests/test_vector_m3.mjs
node tests/test_m2_preview_v06.mjs
node tests/test_m3_study_v07.mjs

# Actual CPU-process resource measurement on your machine
node --expose-gc tests/benchmark_m3_v07.mjs

# Optional UI browser smoke (requires Playwright):
# from LSCNS repository root: python3 -m http.server 8765
# from web directory: node tests/browser_m2_m3_v07.mjs
~~~

**Tests actually performed:** 1550 nm coefficient/reference values, alpha/kappa round trip, independent M1 cases, existing scalar presets, M3 real 29/41/51/69/81 two-mode runs and 69-grid Worker comparison with stored Python result, progress events and old-schema JSON maps. JavaScript source parsing succeeded.

**Not tested in this session:** real deployed-site browser clicks/cancel/export, or a local Node/Chromium process memory measurement. Test files exist for these but are not reported as executed. Strict Yee/edge-FEM criterion remains unmet.
