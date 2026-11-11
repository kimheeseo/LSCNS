# COMSOL PML Settings / Definitions Context — screenshot alignment (2026-11-12)

## Requested behavior
- The first supplied image shows **Form Union/Assembly** with Label=Form Union, Action=Form a union, Repair tolerance=Automatic. That panel remains available through **Form Union (fin)**; no new algorithm was introduced in this update.
- Right-click **Component 1 → Definitions → Perfectly Matched Layer** to create **Perfectly Matched Layer 1 (pml1)** under **Artificial Domains**. The Definitions menu now also displays screenshot context items such as Variable Utilities, Component Couplings, Absorbing Layer and Infinite Element Domain (unsupported actions remain disabled).
- The PML **Settings** editor includes Label, Name, Domain Selection with Manual and Graphics picking, Override (collapsed), Geometry and Scaling.
- **Geometry Type** menu: Cartesian / Cylindrical / User defined.
- When **Cylindrical** is chosen, a two-column **Center coordinate** table displays **Xm (m)** and **Ym (m)** with default values 0, 0. Switching between geometry types retains user-entered coordinates.
- **Scaling**: Coordinate stretching type=Polynomial, Typical wavelength from=Physics interface, Physics=Electromagnetic Waves, Frequency Domain (ewfd), PML scaling factor=1, PML scaling curvature parameter=1. The wavelength-source dropdown also supports a user-defined text expression.
- The selected PML node stores these settings independently. Reopening the same node restores its values; newly created nodes start with the defaults above.
- The existing colored Graphics overlay and manual domain picker are unchanged.

## Implementation files
- `comsol_pml_editor.mjs` and `comsol_pml_editor.css`
- `comsol_definitions_context.mjs`
- `index.html` (asset cache versions bumped)
- `tests/browser_pml_artificial_domains.mjs` expanded for these fields and state restoration
- `tests/browser_definitions_right_click.mjs` updated for screenshot menu items

## Scope and validation
- Source-level JavaScript syntax and field wiring have been checked. The Playwright browser regression scripts have been extended, **not executed against the published site**.
- All PML settings remain a **COMSOL-inspired geometry/editor preview**, not exact COMSOL semantics. In particular, User defined geometry, center-coordinate transforms, coordinate-stretching coefficients, physics-coupled FEM matrices, mesh and confinement loss are *not* computed. Changing PML parameters does not change optical simulation results.
