# COMSOL-like ewfd Model Wizard & Geometry update (2026-11-09)

## Target interaction flow

1. Launch `web/index.html` → **Select Space Dimension**.
2. Select **2D** → **Select Physics** with only **Optics → Ray Optics / Wave Optics**.
3. Expand Wave Optics → **Electromagnetic Waves, Frequency Domain (ewfd)** → **Add**.
4. Click the newly added **ewfd** in **Added physics interfaces**. Only then is **Study** available.
5. Select Study → under **Preset Studies for Selected Physics Interfaces**, select **Mode Analysis**. The selection is added to **Added study**.
6. Click **Done** → existing Model Builder plus **Geometry 1** Settings. A newly finished wizard starts with blank geometry and **nine prefilled, editable Global Definitions → Parameters 1 rows** matching the supplied COMSOL screenshot: d=5[um], Delta=6.75[um], rb=10[um], nbg=1.45, nair=1, wl=1.45[um], neffMax=1.4454, neffMin=1.429, wlr=wl/sqrt(nbg^2-neffMax^2).
7. Click Model Builder **Global Definitions → Parameters 1**, to edit Name, Expression and Description or use **Load from File** (.txt/.csv/.tsv/.json) to add/override values in the Name / Expression / Value / Description table. **Value** is auto-evaluated and displayed in COMSOL-style SI notation. Uploaded definitions are parsed and evaluated, not executed as code. Upload does **not** run the FEM or apply geometry automatically.
8. Click Model Builder **Component 1 → Geometry 1** → Geometry Settings (Units, Constraints and Dimensions, Advanced). Right-click **Geometry 1** → **Circle** to add **Circle 1 (c1)** and open Object Type / Size and Shape / Position / Rotation Angle / Layers / Resulting Entities controls.

## Import test

Upload [COMSOL_microstructured_parameters.txt](./examples/COMSOL_microstructured_parameters.txt) through **Global Definitions → Parameters 1 → Load from File**. Adjust parameter expressions and click **Apply** in the supported geometry workflow before starting mode calculations. The optional sample contains `dn_core` for the browser solver in addition to reference expressions. The nine-row reference default deliberately omits `dn_core`; when absent, the original radial solver uses the existing `delta_n_core` input. The screenshot's `rb=10[um]` is **below** the current radial solver's 15–150 µm outer-radius range, so increase it before attempting that numerical workflow.

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
- After a new ewfd model, nine reference rows are shown and editable; Restore defaults resets the table to those nine rows. Optional sample imports ten rows including `dn_core`; evaluated values and invalid-formula warnings are displayed.
- Geometry 1 right-click → Circle adds `c1` to an empty new model and opens its settings; modifying radius changes preview.
- Existing direct solver mode remains available through the original preset interface, while complex/vector FEM is clearly distinguished from the browser scalar engine.
