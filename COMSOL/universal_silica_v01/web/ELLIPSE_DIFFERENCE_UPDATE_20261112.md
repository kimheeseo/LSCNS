# COMSOL-style Geometry: Ellipse, Difference and individual Delete Entities — 2026-11-12

## Source screenshot mapping

1. **Union result** — existing `Geometry 1 → Booleans and Partitions → Union` keeps an outer merged silhouette by default (`Keep interior boundaries` is off). Overlapping circles and rectangles can form a cross-like outline as in the reference. The contour comes from sampling, not COMSOL's exact B-rep.
2. **Ellipse Settings** — `Geometry 1 → Ellipse` and `+ Ellipse` add a feature with Type (Solid/Curve), a- and b-semiaxes, Sector angle, Base=Center, x/y Position, Rotation, a collapsed Layers section, and resulting-entity selection. Symbolic parameter names such as `a_c`, `b_c` are evaluated using the existing Parameter service.
3. **Difference Settings** — `Geometry 1 → Booleans and Partitions → Difference` adds two independent multi-selection lists: `Objects to add` and `Objects to subtract`. Each has its own ON/OFF Graphics picking button. Selecting an object for one list removes it from the other list. The preview supports `Keep input objects` and `Keep interior boundaries`; Repair tolerance is shown as Automatic only.
4. **Delete Entities** — `Geometric entity level: Object` displays individual preceding object outputs (`rot1(3)`, `arr1(2)`, etc.) in a checkbox list and allows picking from Graphics. Build Selected filters selected whole objects out of downstream Geometry; the source feature remains in the sequence.

## Verification and limitations

- **Difference** computes the union of add objects minus the union of subtract objects, using a bounded occupancy grid; it returns multiple loops for holes and disconnected pieces.
- Supports full solid circles, full solid ellipses, solid rectangles, and nested supported preview Boolean results. An ellipse's sector-angle preview supports partial geometry, but Boolean operations require 360-degree solid shapes.
- A standalone `geometry_boolean.mjs` contains `boundsOf`, `containsPoint`, `unionContours`, `differenceContours`. Rotated ellipse bounds and containment were checked, as were overlapping Union and Difference with an interior hole.
- Browser regression script: `tests/browser_geometry_ellipse_difference.mjs`. Unit regression: `tests/test_geometry_boolean.mjs`. These were added/extended, but full Playwright execution against the published GitHub Pages deployment has **not** been completed.
- These are **2D interactive Graphics previews**, not exact COMSOL CAD Boolean topology, physics domains, meshing, or FEM analysis. Ellipse Layers, precise CAD repair tolerance, and deleting individual boundaries/points are not implemented.
- Existing Circle/Rectangle, Array, Mirror, Rotate, Form Union/Assembly, Zoom Extents, and the original radial FEM remain untouched except for labels and source handling required for these additions.

## Quick manual check

1. In Mode Analysis Geometry 1, add Ellipse 1 with a=`120[um]`, b=`75[um]`, rotation=`20`. Verify the ellipse is displayed.
2. Add Circle 1 with radius=`20[um]` at the center.
3. Under Booleans and Partitions, add Difference. Check `ell1` under Objects to add and `c1` under Objects to subtract.
4. Click Build Selected. The resulting object should have an outer elliptical contour and inner circular opening.
5. Activate `Keep input objects` to preserve ellipse and circle separately while displaying the Difference result.
6. Add Delete Entities, choose Object, select `dif1`, then Build Selected. Its result should disappear while the retained source objects stay visible.
