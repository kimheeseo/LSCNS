# Optical Fiber FEM Tool — Circle and Layer UI update (2026-11-09)

## Scope
This update reproduces the *user-facing* COMSOL-style 2D Geometry operations shown in the supplied screenshots without claiming that the numerical FEM algorithm is COMSOL's proprietary solver.

## User workflow

1. Launch the web tool, choose **2D → Optics → Wave Optics → Electromagnetic Waves, Frequency Domain (ewfd)**.
2. Press **Add**, select the entry in **Added physics interfaces**, press **Study**, select **Mode Analysis**, and press **Done**.
3. In **Global Definitions → Parameters 1**, use **Load from File** to import TXT, CSV, TSV or JSON. Edit Name, Expression and Description cells. Value is recalculated.
4. Right-click **Geometry 1 → Circle**, enter **Radius = rb+wlr**, then **Build Selected**.
5. A gray disk is shown in **Graphics**. Hover over it to highlight it in red. Parameter-driven geometry is recalculated when the expression or Parameter value changes. The Graphics ruler indicates the physical scale.
6. Expand **Layers**, enter **Layer 1 thickness = wlr** and optionally add more rows such as **Layer 2 = wlr/2**. Press **Build Selected** to refresh the inner boundaries.
7. Right-click **Circle 1 (c1) → Duplicate**; a new **Circle 2 (c2)** carries the radius, layer thicknesses and other geometry settings.
8. Right-click a Circle → **Delete** to open the confirmation dialog. **No** cancels; **Yes** removes the selected feature. **Undo** can restore it.

## Geometry equations

For a circle centered at `(x, y)` with outer radius `R = rb + wlr`, inward layers with thicknesses `t_1, t_2, ...` produce boundaries

`R_k = R - sum(t_i for i <= k)`.

Example: `t_1=wlr` → `R_1=rb`; `t_2=wlr/2` → `R_2=rb-wlr/2`.

Negative, undefined, or too-large layer thicknesses are rejected. Live SVG preview recalculates expression values through `LSWaveParameterService.calculate(...)`. Four radial divider lines at the cardinal axes are drawn per positive layer for a COMSOL-like boundary visualization.

## Implemented vs unsupported

- **Implemented:** dynamic parametric circular geometry, gray SVG geometry display, red hover feedback, shape select/duplicate, yes/no delete confirmation, layer table with up to sixteen thickness rows, inward radial boundaries, updated Home ribbon and original Parameters editor.
- **Geometry-only:** Layers/extra circles and off-center/rotated shapes generate a 2D SVG preview. These added features are *not* passed to an arbitrary full-vector 2D Maxwell FEM or Boolean geometry engine.
- **Existing linked solver:** Two concentric 360-degree solid circles *without nonzero extra layers* remain the supported basis of the browser's radial scalar LP01 FEM model.
- Full-vector EWFD, new domains created by layer thickness, arbitrary 2D meshes, PEC/PML physics, and material assignment to each layer require separate numerical work. The UI shall not imply validation of those physics.

## Automation

The CI workflow `.github/workflows/comsol-style-workbench.yml` runs the existing interface and solver regressions as well as `tests/browser_geometry_layers_hover.mjs` to check formula, multi-layer boundaries, hover, duplication, deletion and edited Parameters. Chromium tests run on Ubuntu 24.04.
