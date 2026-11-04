## 2026-11-04 · 3D 구조 / Beam Propagation (2+1D)

[**▶ 웹에서 3D 광섬유 구조·전파 실행**](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#fiber-3d) · [Colab 전파 계산](../propagation_3d/Universal_Silica_Fiber_3D_Propagation.ipynb) · [기술 설명](../propagation_3d/README.md) · [수치 검증](../propagation_3d/VALIDATION_REPORT.md)

Step·Graded·Trench·Multilayer의 고정 실리카 구조에 대해 **브라우저 실시간 2+1D Scalar BPM**으로 z 전파와 실제 강도 단면, MFD·Aeff 변화를 계산합니다. 마우스로 3D 회전, 프레임 애니메이션, CSV·JSON·PNG 저장. **HCF 2·3중 Nested Tube는 3D 단면 형상만** 표시하고 전파·Confinement Loss 계산은 지원하지 않습니다. **Full-vector 3D Maxwell FEM/FDTD 아님**, 경계는 주기적 FFT(PML 미지원).

## Full-vector Fiber Mode Lab v0.1 — 2026-11-04

[**▶ 기존 웹의 Full-vector 결과 보기**](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#full-vector-tool) · [**Google Colab 실제 벡터 계산**](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/universal_silica_v01/full_vector/LS_Full_Vector_Fiber_v0_1.ipynb) · [수치 검증](../full_vector/VALIDATION_REPORT.md) · [기본 결과 JSON](../full_vector/results/vector_mode_result.json)

원형 실리카 Step/Graded/Trench/Multilayer에 대한 **별도 2D Full-vector FDFD (Hx,Hy 결합 고유문제)** 계산 엔진입니다. Ex/Ey/Ez 장을 복원해 **MFD (전기장 적분)**와 **Aeff (전기장 분포 적분)**를 계산하며 독립 HE11 해석해와 비교했습니다. **웹 UI는 Colab에서 계산한 결과를 가져와 보기만 하며, Full-vector 고유값을 직접 계산하지 않습니다.** 기존 웹 Scalar FEM은 그대로 유지됩니다. FEM/복소 PML/HCF 누설손실 기능으로 오인하지 마세요.

## 2026-11-04 · HCF / Nested Tube Designer v0.1

[**▶ 기존 웹에서 HCF 설계 실행**](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#hcf-designer) · [모듈 사용법](hcf_designer/README.md) · [14개 테스트 검증](hcf_designer/VALIDATION_REPORT.md) · [6×2 단면 미리보기](hcf_designer/HCF_Designer_Preview.png) · [8×3 단면 미리보기](hcf_designer/HCF_3Nested_8Tube_Preview.png)

튜브 개수 4–12개, nested 1/2/3중, 코어 직경, 튜브 직경 비율/벽 두께를 입력하면 실제 입력에 대응하는 **2D 단면 SVG를 즉시 변경**하고 공진·반공진 파장, 기하학적 튜브 커버율, 이상적 J₀ 모세관의 neff·MFD·Aeff·D를 갱신합니다. A/B 비교 및 CSV·JSON·SVG 저장 지원. **손실 dB/km는 직접 계산하지 않습니다.** 사용자 제공 복소 neff의 dB/km 환산 기능은 FEM 예측이 아니라 외부값 후처리입니다. 전체 HCF Full-vector/PML 검증 및 튜브 수에 따른 물리 손실 예측은 향후 과제입니다.

# Universal Silica Fiber Workbench — Web UI v0.2.1

**Live browser app:** https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/

A standalone engineering-style interface for the existing `universal_radial.py` radial scalar FEM model. GitHub Pages serves static HTML/CSS/JavaScript and **runs actual finite-element eigenmode calculations in the browser**. It does not simply interpolate precomputed output. No external Python server, COMSOL software or GitHub credentials are needed.

## What the UI does

- Choose **Step-index**, **Graded-index**, **Trench-assisted**, or **Multilayer** circular silica profiles. Edit radius, wavelength, core index step, graded exponent, trench dimensions, concentric layer definitions, mesh resolution and nonlinear index.
- Click **모드 계산 실행** to solve the generalized finite-element eigenproblem for the radial weak-guidance LP0 fundamental mode. Show `neff`, Petermann-II and second-moment MFD, `Aeff`, `γ` and five-point **FEM-derived dispersion D**.
- View a 2D *rotational visualization of the radial mode*, the refractive-index profile, radial intensity, and a recomputed spectral dispersion plot. This is **not** a separately solved 2D Maxwell field.
- Sweep wavelength and independently recompute the FEM at every point. Compare results at multiple mesh resolutions. Export calculation JSON, CSV, PNG, a reusable input-design JSON. Import design JSON.
- **Input change invalidation:** previous result cards are cleared rather than presented as results for the changed input. Users must run the solver again.

## How to use

1. Open the [Web Workbench](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/).
2. Select a preset or set profile parameters in the left panel.
3. Click **모드 계산 실행**. Review the four headline quantities and the three primary figures.
4. Click **파장 스윕** to regenerate `D(λ)`, or **메시 수렴** for convergence checks.
5. Use **JSON 결과**, **CSV 결과**, **모드 PNG**, **설계 JSON** to download computed data. **These export files are downloaded locally, not automatically committed to GitHub.** Avoid placing a GitHub token in the webpage. To version personalized results in the repository, commit the downloaded JSON/CSV through GitHub or ask ChatGPT with the saved files.

## Exact model and limits

- **Solver:** axisymmetric scalar P1 finite elements of `∇²⊥ψ + k0² n² ψ = β² ψ`, transformed to radial form and solved as generalized tridiagonal eigenproblem via shifted inverse iteration. Natural regularity at `r=0`, Dirichlet field at outer computational radius.
- **Materials:** pure silica Malitson Sellmeier `n_cladding(λ)`, other layers are **wavelength-independent Δn offsets**. This is **not** a calibrated dopant-dependent Ge/F dispersion model.
- **Outputs:** eigenmode metrics are scalar approximations. Dispersion is a five-point finite difference of newly solved eigenvalues on identical radial geometry, not a fitted value or a manufacturer's measured dispersion. `γ` uses input `n₂` and computed `Aeff`.
- **Not supported by the LP01 physical FEM:** arbitrary 2D cross sections, microstructured PCF, ARF/NANF HCF, full-vector electromagnetic modes, polarization, birefringence, material attenuation from geometry alone, bending radiation loss from complex modes/PML, confinement loss, or model inversion of unknown manufacturer profiles. **Separate** MCF demonstration and statistical PMD modules remain visible; attenuation/bending-loss research is archived outside the live interface.
- **Domain:** model calculator radius is a numerical boundary, **not** actual 125 µm fiber cladding. Edge-truncation and mesh convergence should be checked for research use. Small solver residual does not establish physical accuracy.
- A G.654.E-like preset uses **assumed 5.5 µm core radius and fitted Δn** to target MFD ≈12.5 µm. This is not independent LE125 certification or COMSOL license-equivalent validation.

## Dispersion-slope diagnostics (2026-11-03)

The web UI now displays three **separate calculated quantities** directly beneath the four original mode metrics, after **모드 계산 실행**:

- **S₁₅₅₀**, the dispersion slope at 1550 nm: \([D(1560)-D(1540)]/20\), units ps/(nm²·km).
- **λ₀**, the zero-dispersion wavelength where \(D(λ₀)=0\), located by scanning 1100–1550 nm and a bracketed root search.
- **S₀**, the dispersion slope **at λ₀**, not at 1550 nm.

All dispersion values use five-point finite differences of freshly calculated radial-FEM `neff(λ)`. The output is recomputed after model changes and stored in exported JSON/CSV. If no zero crossing is found in the search window, λ₀ and S₀ appear as **not available** rather than guessed.

**G.654.E-like example** (`a=5.5 µm`, `Δn≈0.003478579`, 750 elements): S₁₅₅₀ ≈ **0.059346 ps/(nm²·km)**, λ₀ ≈ **1278.901 nm**, S₀ ≈ **0.093321 ps/(nm²·km)**. These are **fitted surrogate-model calculations**, not measured PremierLink LE125 specifications. Always verify whether a manufacturer document refers to `S₀` or `S₁₅₅₀` before comparing.

## Files and verification

- `index.html` (self-contained CSS), `app.mjs`: user interface.
- `solver.mjs`: independent JS implementation of Python radial FEM assembly/integration/eigenmode calculation.
- [`results/baseline_summary.csv`](results/baseline_summary.csv): five preset default calculations.
- [`results/baseline_verification.json`](results/baseline_verification.json): inputs, calculated metrics, Python references and differences; plus a mesh study and wavelength sweep.
- [`results/g654_spectrum.csv`](results/g654_spectrum.csv): wavelength-dependent computed values for illustrative G.654.E.
- [`tests/test_solver.mjs`](tests/test_solver.mjs): repeatable Node.js numeric tests (`node tests/test_solver.mjs`).
- [`VALIDATION_REPORT.md`](VALIDATION_REPORT.md): result and scope summary.

All snapshot numbers on GitHub reflect the checked-in baseline **only**. Changing interactive inputs does not change repository snapshots until a user commits a new run.

## Related

- [Original Python/Colab FEM code](../Universal_Silica_Fiber_v0_1.ipynb)
- [Original v0.1 description](../README.md)
- [Original numerical verification](../VALIDATION_REPORT.md)

## Current web tools and archived loss research (2026-11-03)

The **live GitHub Pages interface no longer displays the attenuation/link-budget or bending-loss tools**, as requested. The PMD/DGD statistical tool remains available alongside the original LP₀₁ radial FEM, dispersion slopes, and MCF demonstration.

- **Active in the web UI:** mode analysis (n_eff, MFD, Aeff, D), S₁₅₅₀ / λ₀ / S₀, wavelength and mesh sweeps, illustrative two-core MCF, statistical PMD/DGD.
- **Removed from the web UI:** attenuation estimator, measured-anchor bend loss, conformal-index/complex-n_eff bend loss postprocessor and external 2D PML/FEM results viewer. Their module scripts are no longer imported by `index.html`.
- **Code and past results are retained as research archives** in the repository, not deleted: [engineering-module documentation](ENGINEERING_MODULES.md), [Velamuri paper note](VELAMURI_PAPER_VALIDATION.md), and [standalone 2D bending FEM notebook/source](../bending_fem/). Previously published experiment reports remain accessible but should not be interpreted as live web functionality.

The PMD module is a **statistical Maxwell-DGD approximation**, not a vector-mode PMD calculation or certified cable PMD_Q value. The core LP₀₁ scalar FEM does not predict intrinsic attenuation or bending radiation loss.

## MCF · Ideal two-core Cross-talk (2026-11-03)

Select the **MCF · 2코어 XT 분석** preset/profile in the existing web UI.
Inputs: two-core pitch 12–24 µm and propagation length 0–100 mm.
Outputs: even/odd supermode Δneff, κ, complete-transfer length Lc,
ideal coherent core 1/2 fractions, and **XT = 10 log10(P2/P1)** in dB.
The cross-talk curve is drawn against propagation length and included in
downloadable CSV/JSON; graph values are visually clipped to [-60, 40] dB,
while the exact transfer null/full-transfer results display −∞/+∞.

**Interpretation:** the browser does NOT calculate a new 2D FEM solution.
It log-interpolates four previously archived legacy M8 scalar-FEM supermode
splittings (pitch 12/16/20/24 µm) at λ=1550 nm, a=4.1 µm and Δn=0.005.
The resulting oscillatory coupling is a lossless, identical-core coherent
model. It cannot predict real **MCF cable XT (dB/km)**, random long-distance
coupling, bend/twist, wavelength-dependent XT, polarization, or measured
product performance. The improved M8 conforming-FEM validation is distinct
from these archived legacy pitch references.

Numerical source: [mcf_xt.mjs](mcf_xt.mjs). Tests:
`node COMSOL/universal_silica_v01/web/tests/test_solver.mjs`.
