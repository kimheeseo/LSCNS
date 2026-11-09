# COMSOL-like ewfd Model Wizard & Geometry update (2026-11-09)

## Target interaction flow

1. Launch `web/index.html` → **Select Space Dimension**.
2. Select **2D** → **Select Physics** with only **Optics → Ray Optics / Wave Optics**.
3. Expand Wave Optics → **Electromagnetic Waves, Frequency Domain (ewfd)** → **Add**.
4. Click the newly added **ewfd** in **Added physics interfaces**. Only then is **Study** available.
5. Select Study → under **Preset Studies for Selected Physics Interfaces**, select **Mode Analysis**. The selection is added to **Added study**.
6. Click **Done** → existing Model Builder plus **Geometry 1** Settings. A newly finished wizard starts with blank geometry and blank Global Parameters, similar to the reference COMSOL screenshots.
7. Click Model Builder **Global Definitions → Parameters 1**, and use **Load from File** (.txt/.csv/.tsv/.json) to populate the Name / Expression / Value / Description table. Uploaded definitions are parsed and evaluated, not executed as code. Upload does **not** run the FEM or apply geometry automatically.
8. Click Model Builder **Component 1 → Geometry 1** → Geometry Settings (Units, Constraints and Dimensions, Advanced). Right-click **Geometry 1** → **Circle** to add **Circle 1 (c1)** and open Object Type / Size and Shape / Position / Rotation Angle / Layers / Resulting Entities controls.

## Import test

Upload [COMSOL_microstructured_parameters.txt](./examples/COMSOL_microstructured_parameters.txt) through **Global Definitions → Parameters 1 → Load from File**. Adjust parameter expressions and click **Apply** in the supported geometry workflow before starting mode calculations. The sample contains `dn_core` for the browser solver in addition to reference expressions.

## Supported vs. UI-only behavior

- **Implemented:** selection wizard, ewfd/Mode Analysis navigation, Model Builder tree, parameter parsing and evaluation, Circle feature editing, 2D preview, existing scalar Radial FEM pathway and existing specialized research links.
- **Not implemented as COMSOL-equivalent solvers:** full-vector 2D EWFD, Ray Optics, frequency-domain Maxwell FEM, advanced boundary conditions/PML, non-circular 2D meshing, Boolean geometry, and disabled study types.
- The UI's **Geometry 1 → Constraints and Dimensions** and **Advanced repair/build settings** are present for workflow compatibility; they do not add a CAD constraint solver.
- The **new wizard** starts with no circles; the legacy design presets still restore their default two-circle model when explicitly selected.
- For the linked scalar Radial FEM, two centered solid 360° circles (`c1` outer domain, `c2` inner core) are necessary; extra shapes and rotations are preview-only.

## Files modified

- `comsol_startup_wizard.mjs`, `comsol_startup_wizard.css` — wizard and Study navigation.
- `comsol_parameter_geometry.mjs` — Parameters import.
- `comsol_geometry_cad.mjs` — Geometry Settings, context menu and blank model.
- `comsol_workspace.mjs` — Component Definitions model tree entry.
- `index.html` — versioned script/styles loading.

## Smoke test scenarios

- Study button is disabled until an **added** ewfd is selected.
- Done is disabled until **Mode Analysis** is selected.
- Parameter sample imports ten rows and displays evaluated values; invalid formulas are retained with an error marker.
- Geometry 1 right-click → Circle adds `c1` to an empty new model and opens its settings; modifying radius changes preview.
- Existing direct solver mode remains available through the original preset interface, while complex/vector FEM is clearly distinguished from the browser scalar engine.
