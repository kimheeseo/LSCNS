# Definitions → Adjacent: named boundary selections (2026-11-10)

## Implemented COMSOL-style workflow

1. In a new 2D / Wave Optics / EWFD / Mode Analysis model, create Geometry 1.
2. Build a circle with `Radius = rb+wlr` and `Layer 1 = wlr` if using the screenshot reference.
3. Add a small Circle 2, position it off axis, and use Rotate 1 with `range(0,60,300)` to form six holes.
4. Optionally create **Materials → Blank Material → Material 1**; this feature and the existing Material domain selection remain available.
5. Select **Definitions** in the upper ribbon, then **Selections → Adjacent**.
6. The model tree receives **Component 1 → Definitions → Selections → Adjacent 1**.
7. Change **Geometric entity level** from Domain to **Boundary**.
8. Click any circular air-hole edge in Graphics to add it to the highlighted blue set, or click again to remove it.
9. Change **Label** to a descriptive name such as `Air Hole Boundaries`. The model tree is updated immediately.
10. Click other Geometry or Material nodes, then return to the named Adjacent selection. Its level, label, and selected edge keys are preserved for the current model session.

## Geometry picking and numerical scope

- The Geometry preview publishes independently selectable circular outlines, layer interfaces, four radial boundary segments per layer, domain regions, and center points.
- Individual air-hole boundaries are labeled by stable Geometry feature keys such as `rot1/c2/0/outer-boundary`.
- A physical mouse click is mapped to the nearest actual circle/line using the Graphics coordinate transform; a repeat click toggles that edge.
- Input Entity levels **Domain / Boundary / Point** are supported for selection. The panel includes an **Output Entities: Adjacent boundaries** section and an exterior/interior boundary control for visual parity with COMSOL.
- **Limitation:** This is currently a named **2D Geometry selection and labeling tool**, not a full COMSOL topological adjacency operator. It does not yet derive adjacent boundaries from mesh connectivity, perform general Boolean Union, or assign these selections to a full-vector EWFD FEM solver. Checkbox settings under Output Entities are stored as preview settings but do not generate a topological adjacency result.

## Regression testing

Run `node COMSOL/universal_silica_v01/web/tests/browser_adjacent_boundary_labels.mjs` in a prepared Playwright environment and/or inspect the **COMSOL-style Workbench UI Regression** GitHub Actions workflow. The test checks:

- Blank Material and existing Geometry compatibility
- Definitions ribbon Adjacent command activation
- Automatic Model Builder selection-node creation
- Domain → Boundary switch and stable independent edge IDs
- Actual physical mouse click on a circular air-hole boundary
- Two independently highlighted air-hole boundaries
- Label update, tree reopening, and persistence after switching panes

Screenshots uploaded by the workflow:

- `/tmp/comsol_adjacent_material1_before.png`
- `/tmp/comsol_adjacent_created.png`
- `/tmp/comsol_adjacent_two_boundaries_selected.png`

## Graphics correction: reference Domain selection view

- Adjacent opens with uniform gray built domains, thin solid outlines for original and Rotate objects, no Geometry feature-name labels, and the grid off by default. The Grid button toggles an independent Adjacent grid preference.
- Geometry, cumulative-selection and material highlights do not carry into Adjacent. Returning to Geometry restores its editing display.
- With Input Entities set to Domain, physical mouse coordinates identify exactly one preview region. Hover outlines that region in red; clicking outlines it in blue and adds it to Input selections. Clicking again removes only that region.
- Boundary and Point modes retain their individual input selection. Pausing Activate Selection preserves the displayed selection while blocking further picking and accidental Geometry editing.
- The browser regression now checks real Domain hover/click/toggle, paused selection, empty margin clicks, the neutral Graphics appearance, and the existing Boundary workflow. This remains a preview selection operator; the numerical scope described above is unchanged.
