# COMSOL guidebook alignment — current-version update (2026-11-10)

Reference: *Introduction to COMSOL Multiphysics*, version 5.5 (214-page supplied introduction). This is a general Model Builder introduction, not a Wave Optics theory manual. Page references below are the printed page numbers, which match the supplied PDF pages.

The update starts from `8ecbf98`, preserving the existing Material Contents validation, Definitions context menu, PML Artificial Domains editor, per-domain Graphics picking, Adjacent display, and panel resizers. The recent screenshot-driven UI changes take priority over generic examples in the guide.

## Changes

| Guidebook section | Current-version behavior |
| --- | --- |
| Model Builder, pp. 17–18; Appendix B, pp. 176–178 | Up/Down selects the previous/next visible tree node. Left/Right collapses/expands branches. Switching Settings preserves keyboard focus despite dynamic tree rebuilding. |
| Previous/next selection, p. 177 | Alt+Left/Right returns to existing node selections without losing Geometry, Material, Adjacent or PML settings. |
| Disable/Enable, p. 176 | F3 disables the current Geometry feature; F4 enables it. Repeated presses are idempotent. The context menu shows the corresponding shortcut. |
| Move a node, p. 178 | Ctrl+Up/Down moves the selected Geometry feature by one step using the existing ordering operation. |
| Selection and list editing, pp. 25, 75, 177 | Material Manual, Adjacent inputs and PML Manual provide All/Clear controls; Ctrl+A/Ctrl+D acts on the current entity level when focus is in Graphics/selection Settings. Delete on a focused selection checkbox removes that item only. Text-field editing is left intact. All/named modes are not silently converted to Manual. |
| Dynamic help, pp. 10, 18, 176 | F1 or Developer → Context Help opens local guidance for the current editor. It updates while selecting other nodes; Escape closes Help without closing Settings. |

PML uses the same common selection and keyboard conventions. The introduction is not a source for its Wave Optics coordinate-stretching equations.

## Scope retained

- Circle/Rotate/Mirror/layer geometry, named selections, Material domains and PML settings remain a preview workflow for arbitrary 2D cross-sections.
- Custom Material/PML/Adjacent settings are not newly connected to the numerical solver by this update.
- The existing linked two-domain scalar radial FEM is unchanged. No claim is made that the guidebook's structural, heat-transfer, multiphysics, or general full-vector FEM algorithms have been implemented.
- The existing Open/Save Design JSON controls retain their existing scope. This update does not claim a complete saved-model round trip for every editor.
- This is a focused common-UI alignment pass, not an assertion of complete feature parity with all 214 pages or all COMSOL products.

## Verification

`tests/browser_guidebook_interactions.mjs` uses real keyboard and checkbox input to check tree/history navigation, repeated F3/F4, Geometry reordering, scoped bulk selection and Delete, native text selection, per-level state, PML settings preservation, and F1/escape/context updates. The existing workbench GUI, Material/PML, physics and preset regressions remain enabled.
