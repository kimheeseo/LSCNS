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

## Dynamic Geometry 1 · Circle and Rotate (2026-11-04)

- **Geometry 1** contains ordered, dynamically added `Circle N (cN)` and `Rotate N (rotN)` features. Select a feature from the Model Builder tree to see its specific **Settings** panel. Existing design-example presets remain available below the sequence.
- **Circle Settings:** Object Type (`Solid`/`Curve`), Size and Shape (Radius, Sector angle), Position (Center, x/y in metres, including parameter expressions), Rotation Angle, Layers, and **Selections of Resulting Entities** (selection checkbox, Show in physics: Domain selection, color None/Blue/Orange/Green). The selection metadata and color affect the UI preview; the arbitrary-domain physics selection feature is not implemented in Radial FEM.
- **Rotate Settings:** reference one or more **preceding Circle nodes**, optionally keep the input objects, enter a single angle or `range(start,step,end)` (e.g. `range(0,60,300)`), and choose the center of rotation x/y. Up to 72 angles and 450 total preview shapes are allowed to bound browser work.
- **Right-click Geometry 1** to display a COMSOL-style context menu: Build All, Circle, and Transforms → Rotate. Other visible COMSOL commands that are not implemented (Ellipse, Fillet, Boolean, etc.) are **disabled**, not decorative false functions.
- **Right-click Circle or Rotate** for Build Selected/Preceding, Add Before/After (Circle/Rotate), Move Up/Down, Copy, Duplicate, Delete, Disable/Enable, Rename, Settings, Properties and Help. Copy as Code copies that feature's JSON if clipboard access is available.
- **Graphics** switches to an independent 2D drawing with grid and zoom controls while CAD is active. The geometry SVG is distinct from FEM mode plots. Click drawn features to select their editor; Build Selected and Build Preceding restrict the displayed sequence.
- **Existing solver compatibility:** only the supported two concentric full Solid circles (Circle 1 as external computation domain, Circle 2 as core) can be applied to the 1D Radial FEM; enabled Rotate nodes or extra circles must be disabled before applying. Off-center circles, sectors, curves and added layers remain geometry previews, **not 2D full-vector FEM calculations**.
- New files: `comsol_geometry_cad.mjs` and `comsol_geometry_cad.css`. The existing parameter resolver remains the single authority for named expressions; the original FEM solver is unchanged.
- Browser coverage: `tests/browser_comsol_parameters_geometry.mjs` tests the parameter link, extra Circle addition, `range(0,60,300)` Rotate copies, context-menu operations and Radial FEM safety guards.

## Wave Optics Mode · COMSOL 5.5 UI reference (2026-11-04)

**Source design references**: user-provided *Introduction to COMSOL Multiphysics* (COMSOL 5.5, Model Builder/Settings/Geometry/Materials/Mesh/Study/Results) and COMSOL's Wave Optics Module guide (Electromagnetic Waves, Frequency Domain, `ewfd`, with default Wave Equation, Electric / Perfect Electric Conductor / Initial Values nodes).

- The workbench initially opens **Settings → Electromagnetic Waves, Frequency Domain (ewfd)** in the model tree and retains the existing numerical Graphics and result cards.
- **Model Builder**: Component 1 → Geometry 1 (Circle/Rotate CAD); Materials → Core, Cladding; Electromagnetic Waves, Frequency Domain (ewfd) → Wave Equation, Electric 1; Perfect Electric Conductor 1; Initial Values 1; Mesh 1 → Size 1; Study 1 → Mode Analysis; Results → Mode Intensity, Refractive Index, Derived Values.
- Selecting a node changes its **context-specific Settings** without cloning or replacing the original numerical solver. A **Fiber settings** back button restores the prior design inputs.
- **Connected, functional numerical settings**: core-cladding index difference `delta_n_core`, cladding Sellmeier/fixed index model `clad_material`/`clad_index_override`, analysis wavelength `wavelength_nm`, and the radial FEM mesh size `elements`. Apply and Compute call the original authoritative fields and buttons; Mode Analysis uses the actual existing scalar LP₀₁ solver, and results panels read actual computed values.
- **Physics GUI reference only**: COMSOL's `ewfd` and its Wave Equation / Perfect Electric Conductor / Initial Values nodes are shown for a familiar workflow, but the web app does **not** claim a general 2D full-vector COMSOL-equivalent Maxwell FEM, PEC/PML boundary conditions, or arbitrary-domain material assignments. Those would require separate physics/mesh implementations. HCF and Full-vector Colab research tools remain available as before.
- **Mesh GUI reference**: the active mesh is 1D radial P1, not COMSOL's 2D triangle/mapped or 3D tetrahedral mesh. Existing mesh convergence tests remain functional.
- **Results**: computed `n_eff`, MFD, `A_eff`, D and the existing mode/index plots are available; real results are not fabricated.
- **Smoke test**: `tests/browser_comsol_wave_optics.mjs` verifies ewfd tree, Materials/Physics/Size/Study/Results settings, parameter applications, actual solver execution, and switching between original Geometry/Parameters/HCF modules. The preexisting COMSOL and CAD tests continue to run.

See `comsol_wave_optics_settings.mjs` and `comsol_wave_optics_settings.css`. This is a COMSOL-inspired independent application, not affiliated with COMSOL AB.

## COMSOL UI / Operation Phase 1–2 (2026-11-04)

The design and numerical backends are deliberately separate. **No eigenmode, HCF, MCF, PMD, FDFD or FEM solver source is altered in this phase.**

### Phase 1 — UI fidelity
- A COMSOL-style **Model Builder → Settings → Graphics → Add Study** desktop workspace, ribbon actions, and Messages/Progress/Log dock are preserved.
- **Geometry 1** is now its own Model Builder node, separate from **Circle / Rotate** features: clicking the Geometry label opens root Settings (length-unit display selector, ordered feature sequence, and finalization explanation); clicking the triangle only expands/collapses the tree.
- Wave Optics **ewfd**, Materials, Mesh, Study, Results remain in the Model Builder; each opens a context-specific Settings panel. The quick-access strip exposes Geometry Undo, Redo, Build All.

### Phase 2 — Operator behavior
- **Geometry 1 right click**: existing Circle, Transforms → Rotate, Build All menu.
- **Feature right click**: Build Selected, Build Preceding, insert before/after, Copy, Paste, Duplicate, Move Up/Down, Disable, Delete, Rename. Only commands with actual implementations are enabled; unsupported 2D Boolean operations remain disabled.
- Geometry edits are undoable and redoable using the Settings toolbar or **Ctrl+Z / Ctrl+Shift+Z** (when focus is outside text controls). Copy/Paste creates distinct feature IDs. History resets when the design preset is changed.
- Graphics supports click-to-select, object selection dropdown, mouse wheel zoom, pan by dragging the blank canvas, grid toggle and Fit. The 2D view is a **geometry preview, not a solved electromagnetic field**.
- **Build Selected/All** regenerates the preview of the configured feature sequence. The **Apply supported geometry to Radial FEM** operation remains separate: only the supported concentric two-circle model affects authoritative numerical inputs. No general 2D Maxwell-equivalent FEM is implied.
- Tests: `tests/browser_comsol_ui_operations.mjs` plus existing `browser_comsol_workspace.mjs`, `browser_comsol_parameters_geometry.mjs`, `browser_comsol_wave_optics.mjs`. CI workflow: `.github/workflows/comsol-style-workbench.yml`.

## Phase 3 · Connected COMSOL-style workflow (2026-11-04)

The Model Builder's **Geometry → Materials → Mesh → Study → Results** sequence is wired to the **existing scalar radial LP₀₁ FEM**, while leaving the numerical solver, HCF/MCF tools and Colab Full-vector research implementation unchanged.

- **Geometry 1 / Circle 1 / Circle 2**: geometry sequences continue to support arbitrary 2D preview features, but the linked solver accepts **only the original two enabled, concentric, 360° Solid circles**. Their geometrical regions are represented as **Domain 1 = Core (Circle 2)** and **Domain 2 = Cladding (inside Circle 1, outside Circle 2)**. The public `window.LSWaveCADService.status()` exposes solver eligibility, applied vs pending geometry state, and evaluated radii; `apply()` writes the supported parameters into original FEM form controls.
- Geometry edits and changes to **Global Parameters** (e.g. `d` in `d/2`) invalidate the applied-geometry status. The linked workflow refuses to Compute until the updated geometry is supported and has been **Applied**. Arbitrary Rotate or extra Circle operations are preview-only and block the linked calculation while enabled. This guard applies to the **linked Study workflow**; the legacy standalone Compute input panel remains available for independent model operation.
- **Materials**: Core and Cladding are clickable, numbered radial regions in the Material Settings. Selecting either opens the real editable underlying FEM inputs: `delta_n_core` for Core (`n₁=n₂+Δn`), and `clad_material` / `clad_index_override` for Cladding. **This is a real radial material assignment, not general 2D geometry-domain selection.**
- **Mesh**: Mesh 1 → Size 1 edits the original scalar FEM element count (`elements`), not a separate decorative representation or a general 2D triangular FEM mesh.
- **Study**: `Compute linked mode` verifies scalar-mode eligibility, supported/applied geometry, valid refractive-index and mesh inputs, and invokes the **original real `calculateBtn`**. The original asynchronous solver's completion message is observed to update study completion.
- **Results**: The workflow tracks the actual authoritative model-input signature. Editing wavelength, core/cladding material properties, radii, or mesh marks previously computed results **Stale** until the same model is computed again. Actual `n_eff`, MFD, `A_eff`, and dispersion results are read from the existing solver output.
- The right **Model Workflow** sidebar provides 5 step-navigation buttons, Geometry compatibility and Apply controls, Compute, and clear Ready/Blocked/Pending/Current/Stale indications.
- Automated Playwright regression: `tests/browser_comsol_linked_workflow.mjs` plus the prior four UI, numerical and geometry browser tests, managed by `.github/workflows/comsol-style-workbench.yml`.

**Limitations:** This release does not implement true COMSOL 2D/3D `ewfd` solving, PEC/PML boundaries, arbitrary domain topology, generic finite-element mesh generation, or Full-vector Maxwell eigensolvers. It does not validate manufactured fiber profiles absent measured refractive-index data. Future steps can extend the computational geometry engine separately.
