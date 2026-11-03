# Universal Silica Fiber Lab — Engineering Loss & PMD Modules

**Added 2026-11-03 · Phases 1→2→3 completed in order.**

These modules provide **parameter/measurement-anchored engineering estimation**, not full Maxwell or dopant-specific physics. Existing scalar LP01 FEM mode and G.654.E validation functionality remain separate and unchanged.

## 1. Attenuation / link budget

- **Engine:** [attenuation.mjs](./attenuation.mjs)
- **UI:** [attenuation_ui.mjs](./attenuation_ui.mjs)
- **Test:** [tests/test_attenuation.mjs](./tests/test_attenuation.mjs)
- **Inputs:** α1550, α1625 (dB/km), query λ ∈ [1550,1625] nm, length (km), Pin (dBm), connector and splice counts/losses.
- **Method:** linear interpolation of measured/specification anchor attenuation; no extrapolation beyond the anchors.
- **Outputs:** α(λ), fiber loss, joint losses, total link loss and Pout, curve of loss vs length, JSON download.
- **Validation example (illustrative, NOT a manufacturer measurement):** α1550=0.160 dB/km, α1625=0.175, λ=1550, length=80 km, 2 connectors ×0.2 dB, 1 splice ×0.05 dB, Pin=0 dBm ⇒ 12.8 dB fiber +0.4 connectors +0.05 splice ⇒ **13.25 dB total / −13.25 dBm output.**
- **Boundary:** no Rayleigh/OH physics identification, temperature effects, bending/PMD, or spurious extrapolation beyond measured wavelengths.

## 2. Macrobending loss (measured-anchor model)

- **Engine:** [bending_loss.mjs](./bending_loss.mjs)
- **UI:** [bending_ui.mjs](./bending_ui.mjs)
- **Test:** [tests/test_bending_loss.mjs](./tests/test_bending_loss.mjs)
- **Inputs:** two measured bending radii (mm) and four **positive measured loss-per-turn anchors** at 1550 and 1625 nm (2 radii × 2 wavelengths), query radius and wavelength, number of turns, straight-link loss.
- **Method:** log(loss/turn) interpolation in reciprocal radius 1/R, then log interpolation across λ. Strict **no extrapolation**. Additional bend loss is per-turn × turns, assuming additive turns.
- **Outputs:** dB/turn, total bend loss, bend length, straight-link + bend total, interpolated curve and JSON.
- **Validation example (illustrative, NOT a YOFC/LE125 measurement):** anchors at radii 10/20 mm; 1550 nm losses 0.15/0.005 dB/turn, 1625 nm losses 0.32/0.015 dB/turn; at 15 mm, 1550 nm, 10 turns ⇒ **0.015536 dB/turn**, **0.155362 dB additional bend loss**. If 13.25 dB baseline, combined **13.405362 dB**.
- **Boundary:** This is **not** scalar 2D/3D curved-waveguide FEM with PML or radiated-loss physics; do not call computed values product-level bending predictions. No zero-loss anchors in logarithmic interpolation; record the instrument detection limit and use justified positive floor if needed.

## 3. PMD / statistical DGD

- **Engine:** [pmd.mjs](./pmd.mjs)
- **UI:** [pmd_ui.mjs](./pmd_ui.mjs)
- **Test:** [tests/test_pmd.mjs](./tests/test_pmd.mjs)
- **Inputs:** PMD coefficient (ps/√km), length (km), coefficient **meaning**: mean-DGD/√km or RMS-DGD/√km.
- **Assumption:** strongly coupled, statistically stationary, isotropic random birefringent sections, resulting in an approximate **Maxwell DGD** distribution. DGD scales as √L.
- **Outputs:** mean, RMS, median, P95 and P99 DGD, Maxwell probability-density curve and JSON.
- **Validation example (illustrative):** mean DGD coefficient=0.05 ps/√km, 80 km ⇒ mean DGD=**0.447214 ps**, RMS=**0.485406 ps**, P95=**0.783433 ps**, P99=**0.943942 ps**.
- **Boundary:** A single-link DGD distribution quantile **is NOT** the manufacturer cable **PMD_Q** population specification. Real polarization-dependent n_eff and group delay, stress, twist, frequency-dependent Jones matrix and time evolution are not calculated in the present scalar LP01 FEM.

## Development/testing

Run in repository root (Node.js 18+):

```sh
node COMSOL/universal_silica_v01/web/tests/test_solver.mjs
node COMSOL/universal_silica_v01/web/tests/test_attenuation.mjs
node COMSOL/universal_silica_v01/web/tests/test_bending_loss.mjs
node COMSOL/universal_silica_v01/web/tests/test_pmd.mjs
```

The initial UI numeric values in all 3 sections are **examples only**, not measured values of a specific commercial fiber. For a manufacturer comparison, first supply production/measurement data with measurement conventions and specimen conditions.

## Future fidelity upgrades (not yet implemented)

1. **Attenuation:** full spectral measured data and a calibrated Rayleigh + IR + OH model with identifiable coefficients, uncertainty bars and reference checks.
2. **Bending loss:** 2D full-vector or scalar curved-index eigenmode with PML, complex beta, curvature and PML convergence, then IEC 60793-1-47 mandrel-winding reference comparisons.
3. **PMD:** ellipticity/stress anisotropy vector eigenmodes and wavelength-dependent group birefringence, random Jones-section coupling, measured IEC 60793-1-48 / ITU-T G.650.2 reference comparisons.

## Test standards / context

- [ITU-T G.650.3](https://www.itu.int/epublications/publication/itu-t-g-650-3-2017-08-test-methods-for-installed-single-mode-optical-fibre-cable-links): installed link attenuation (includes splice/connector loss); PMD and spectral attenuation measurement context.
- [ITU-T G Suppl. 40 (2024)](https://www.itu.int/rec/dologin_pub.asp?id=T-REC-G.Sup40-202407-I%21%21PDF-E&lang=f&type=items): lists IEC 60793-1-47 macrobend-loss and IEC 60793-1-48 PMD measurement methods.
- [ITU-T G.657](https://www.itu.int/rec/dologin_pub.asp?id=T-REC-G.657-202408-I%21%21PDF-E&lang=e): practical macrobend measurement and number-of-turns conventions.
