# Conversions and Zoom Extents — 2026-11-11

This update builds on the working Array implementation and preserves the current UI and solver behavior.

## Conversions

Geometry 1 → right-click → Conversions now opens the screenshot's submenu:

- Convert to Solid
- Convert to Curve
- Convert to Point
- Split

These four operations are displayed but disabled because their settings and implementations have not been specified yet. This preserves the earlier requirement that unfinished features must not be enabled as inert buttons. Conversions itself opens the submenu; it is positioned immediately below Transforms.

## Zoom Extents

The Graphics fit button is now labeled **Zoom Extents** in its tooltip and accessible name, with an extents icon. It fits all currently displayed Geometry objects, including full array outputs.

Previously the view used the greatest distance from the origin and kept the origin centered. A long one-sided array therefore wasted space on the empty side and did not use the available Graphics area well. The new camera uses the union of all circle outer bounds:

- xmin = min(x − radius); xmax = max(x + radius)
- ymin = min(y − radius); ymax = max(y + radius)
- camera center = the midpoint of these bounds
- uniform scale = the smaller horizontal/vertical scale that fits the bounds with an 8% margin on each side

Zoom Extents resets manual pan and zoom while preserving object coordinates, dimensions, selections and material/PML assignments. Zooming after a fit remains centered on the fitted world position. The miniature overview fits the full geometry independently of the main camera. When Geometry units are µm, a long array keeps µm coordinate labels.

The same world-to-screen transform is used for rendering and inverse coordinate picking in Material, Adjacent and PML. Nonzero camera centering therefore does not change which domain or boundary a physical mouse click selects. The existing radial FEM is unchanged.

## Verification

- `tests/test_geometry_view.mjs`: exact six-disk bounds; horizontal, vertical and translated diagonal fits; differing viewport aspect ratios; pan/zoom and inverse coordinate mapping; the existing concentric-circle scale; empty view.
- `tests/browser_zoom_extents_conversions.mjs`: menu labels/state; zoom/pan a six-disk array until cropped, then fit and verify every disk is inside Graphics; stable repeated fitting; physical world-coordinate labels; zoom center; overview behavior; vertical and translated layouts; actual Material/Adjacent/Boundary/PML clicks after fitting; Geometry reset.
- Existing Array, Circle/Rotate/Mirror, Materials, Adjacent, PML, panel resizers and linked solver regressions remain enabled.

Arbitrary 2D shapes remain a Geometry preview; this update does not create general-purpose CAD conversion or a full 2D FEM mesh.
