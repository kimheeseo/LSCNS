# COMSOL Materials — Relative permeability / Relative permittivity (2026-11-12)

## Screenshot-aligned workflow
1. Open **Materials → Blank Material → Material Properties → Basic Properties**.
2. Right-click **Relative permeability** or **Relative permittivity** and select **Add to Material**.
3. Each property is added as an independent **Material Contents** row with columns **Status**, **Property**, **Variable**, **Value / Expression**, **Unit**, **Property group**.
4. **Relative permeability** uses variable `mur_iso` (unit 1), **Relative permittivity** uses `epsilonr_iso` (unit 1).
5. The added value is initially empty and has an invalid/missing status. A finite real scalar or valid Parameters 1 expression (e.g. `1`, `2.25`) changes its individual status to **✓**. Invalid or empty expressions display **!**, without changing the other row's status.
6. Right-clicking **Add to Material** again does not duplicate the property. Duplicated materials copy both Basic Properties' values independently. Existing Refractive index `n_iso` / `ki_iso` behavior remains.

## Implementation & checks
- Files: `comsol_materials_editor.mjs`, `comsol_materials_editor.css`, `index.html`.
- Browser workflow test: `tests/browser_material_relative_permeability_permittivity.mjs`; it requires Playwright and a running local app server.
- JavaScript syntax, source-field bindings, uniqueness of Material UI IDs, and page asset versions were checked. Browser execution is not yet verified.
- `Relative permeability` and `Relative permittivity` are currently **UI/editor properties of the Geometry material preview**. They are not passed to a complete Maxwell/FEM solver or used for mesh/material physics calculations. Only finite real-valued scalar expressions are evaluated; complex tensors are not supported.
