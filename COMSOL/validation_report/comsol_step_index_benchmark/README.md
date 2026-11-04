# COMSOL public Step-Index Fiber vs LSCNS radial FEM — reproducible benchmark

This benchmark uses COMSOL's **public generic step-index fiber example**, **NOT** a commercial or standards-unique G.654.E fiber. See [COMSOL official documented inputs](https://doc.comsol.com/6.4/doc/com.comsol.help.models.rf.step_index_fiber/step_index_fiber.html).

**Matching physical conditions:** 1550 nm, core radius 8 µm, cladding radius 40 µm, n(core) 1.4457, n(cladding) 1.4378. COMSOL publishes the fundamental neff rounded as **1.4444** (V≈4.895).

The current production JavaScript radial FEM assumes silica Sellmeier as cladding index, so its standard **G.654.E concept preset cannot be directly compared** to this COMSOL example without changing the material baseline.

`run.mjs` therefore **imports the exact production solver source**, substitutes *only* the two cladding baseline expressions with COMSOL's constant index 1.4378, and solves the same generalized axisymmetric scalar FEM matrix, quadrature and PII/Aeff postprocessors. It performs an independent mesh sweep and compares the neff with COMSOL's published 4-decimal value and separately with the exact analytical **scalar LP01** Bessel solution (1.4444018444092133). The production web model is not mutated or upgraded by this test.

Run: `node COMSOL/validation_report/comsol_step_index_benchmark/run.mjs`.

## Methodological limitations
- COMSOL: full-vector mode analysis. LSCNS radial FEM: scalar weak-guidance approximation. They will differ, especially for high index contrast.
- The four-decimal COMSOL value carries rounding uncertainty of roughly ±0.00005; claiming relative agreement more precise than this is not justified.
- Exact scalar LP01 Bessel equation is a **separate mathematical oracle**, not the COMSOL reported result.
- Same G.654.E standard classification or same MFD specification is not enough to verify the same material RIP. Actual LE125/YOFC profiles need their own measured RIP input for fair COMSOL-vs-LSCNS performance comparisons.
