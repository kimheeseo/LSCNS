# COMSOL-style UI / LS Wave Optics Workbench

**Visual / navigation layer only.** No changes to the numerical FEM / FDFD / PML / HCF / MCF / PMD / BPM engines or scientific assumptions.

## COMSOL-inspired UI

- Home / Definitions / Geometry / Sketch / Materials / Physics / Mesh / Study / Results / Developer **functional ribbon tabs** with direct commands routed to the existing controls.
- **Model Builder** tree with real clickable example presets, global definitions, component geometry, physics, mesh, study and results. Tree nodes use the authoritative `#preset` selection and its existing change handlers.
- **Settings** center pane: retains the original controls, formulas, numerical parameters and IDs, without moving or cloning the functional elements.
- **Graphics** results viewport: the existing mode plots, neff/MFD/Aeff, dispersion, HCF/NANF, MCF, Full-vector, 3D propagation and PMD panels are kept; only light-mode appearance and visual order are changed.
- **Add Study** sidebar: mode analysis, wavelength sweep, mesh convergence, Full-vector Colab, HCF FEM Colab, PMD and BPM are linked to existing functional modules.
- **Messages / Progress / Log**: reflects the existing actual FEM engine/notice text, with collapsible dock. Does not fabricate computed values.
- Supports desktop 4-column layout and mobile responsive fallback. Existing reference Step-index neff benchmark remains available.

## Use

Open [the workbench](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/), select a design from **Model Builder** or the original **설계 예제** dropdown, set numerical inputs in **Settings**, and run **Compute**. Inspect **Graphics** and message log. HCF / Full-vector research presets still use the existing JSON / Colab workflows.

The UI resembles COMSOL's desktop ribbon/pane structure by user request. **LS Wave Optics is independent software; the visual style does not add COMSOL FEM capabilities or imply affiliation with COMSOL.**

## Tests

- Browser regression: `COMSOL/universal_silica_v01/web/tests/browser_comsol_workspace.mjs`.
- Includes actual Chromium screenshot and computes the public COMSOL step-index benchmark from the original solver after selecting the model through the tree. The previous independent physics tests and repository workflows remain unchanged.
