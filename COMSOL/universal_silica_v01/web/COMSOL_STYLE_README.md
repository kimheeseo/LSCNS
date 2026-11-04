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

## Global Definitions → Parameters / Component → Geometry (2026-11-04)

- Click **Model Builder → Global Definitions → Parameters 1** for an editable COMSOL-style table with Name, Expression, Value and Description. Supports user parameters, dependencies, arithmetic, `sqrt()` and `[m]/[mm]/[um]/[nm]` units. Invalid values and cyclic references are marked, and the table and geometry definitions can be exported as JSON.
- Click **Component 1 → Geometry 1 → Circle 1 / Circle 2** to edit Object Type, Radius, Sector Angle, Position x/y, Rotation Angle and optional Layer 1 thickness. The settings show a live 2D SVG geometry preview. Circle 1 is the outer calculation domain; Circle 2 is the core for the standard concentric radial model.
- Select **Apply concentric circles to Radial FEM** to copy valid radius / wavelength / core-cladding index contrast into the **existing** authoritatively-bound numerical controls. Run Compute after applying. No original numerical solver is replaced.
- **Parameter naming:** `d` is Circle 2's diameter for the radial SMF example; `Delta` is a reference hole pitch to match the supplied COMSOL screenshot; `dn_core` (not `Delta`) is the actual refractive-index difference applied to the radial FEM. `rb` is the outer domain radius, `wl` is the wavelength, and `nbg` is an illustrative background index unless its checkbox explicitly fixes the cladding index. `neffMin`, `neffMax`, `wlr`, `nair`, and hole pitch `Delta` are **reference expressions**, not eigenvalue solver bounds or HCF geometry.
- **Physics limitations:** Only two concentric 360° Solid circles with no additional layers can be passed into the current 1D Radial FEM. Nonzero x/y, partial sectors, Curve objects and geometric layers can be edited/previewed, but must not silently alter numerical results. HCF and Full-vector designs keep their existing specialized workflows and Colab solvers.
- **Restore original sidebar:** click **Fiber settings** or press Escape. The original geometry charts and numerical controls are retained. This is a COMSOL-inspired editor, not an embedded COMSOL application.
- Files: `comsol_parameter_geometry.mjs`, `comsol_parameter_geometry.css`. Existing `comsol_workspace.mjs` routes actual tree actions into these views.
