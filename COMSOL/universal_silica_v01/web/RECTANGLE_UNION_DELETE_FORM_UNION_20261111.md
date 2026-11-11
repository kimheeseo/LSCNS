# COMSOL-style Geometry: Rectangle, Union, Delete Entities, Form Union (2026-11-11)

## User-requested flows

1. **Geometry 1 → Rectangle** or **+ Rectangle** creates an independent feature with Object Type, Width, Height, Base (Corner/Center), Position (x/y), Rotation, and Selections of Resulting Entities. The rectangle is drawn in the 2D Graphics pane, including rotation.
2. **Geometry 1 → Booleans and Partitions → Union** or **+ Union** creates an editable Union feature. In Input objects select at least two preceding solid circles/rectangles or individual Array copies. **Build Selected** generates an outer joined silhouette, hides the consumed inputs by default, and supports Keep input objects and Keep interior boundaries. The union is a sampled preview contour, not a B-rep or FEM-domain Boolean.
3. **Geometry 1 → Delete Entities** or **+ Delete Entities** creates a feature with Geometric entity level = Domain and input object selection. Build Selected filters chosen full 2D preview objects from downstream output; later Geometry features see the reduced set. Boundary/Point topology deletion remains disabled.
4. **Form Union (fin)** appears in the Geometry tree. It exposes Form Union enabled, Keep interior boundaries, Repair tolerance (display setting), and Build Finalization. It can preview a combined silhouette for supported solid circles and rectangles. This does not modify the 1D radial FEM solver or produce a 2D FEM mesh.

## Important behavior / limitations

- User-supplied screenshots were used as a UI and workflow reference; exact COMSOL numerical CAD topology is not claimed.
- 2D preview union uses bounded sampling (340 cells on the longer bounding-box side), so curved edges are approximated and output should **not** be treated as an engineering-quality CAD boundary or used for mesh generation.
- Delete Entities removes **whole preview objects/domains**, not a subset of a boundary, edge, or point.
- Rectangle and Union are omitted from the existing circular/radial Materials and Adjacent Domain model rather than falsely displaying a bounding circle as an exact rectangle domain.
- Existing Circle, Array, Rotate, Mirror, input checkbox selection, Build Selected / Build All, Zoom Extents, undo/redo, and JSON export remain in the same editor.
- The module is loaded from `geometry_boolean.mjs` in the same directory.

## Checks included

- `tests/test_geometry_boolean.mjs`: rotated Rectangle coordinates, full-circle and rectangle containment, connected and disconnected Union regions, micron-scale numerical stability.
- `tests/browser_geometry_rectangle_union_delete.mjs`: UI workflow smoke test for Rectangle → Union → Delete Entities → Form Union.
- Updated `tests/browser_geometry_array.mjs` menu expectation: Union is now enabled; other currently unsupported Boolean entries remain disabled.
- These browser scripts were committed but not executed against the deployed GitHub Pages site in this session.
