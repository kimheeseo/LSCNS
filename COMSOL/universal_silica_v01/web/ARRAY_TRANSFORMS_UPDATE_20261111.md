# Geometry Array and context menus — 2026-11-11

Implemented against the current guidebook-aligned workbench, preserving existing Circle, Rotate, Mirror, Materials, Adjacent and PML editors.

## Lecture screenshot workflow

1. Right-click Geometry 1. **Booleans and Partitions** appears immediately above **Transforms**.
2. Transforms lists **Array, Copy, Mirror, Move, Rotate, Scale**, in that order. Array, Mirror and Rotate are operational. Copy/Move/Scale remain disabled until their settings/behavior are specified and implemented, preserving the user's earlier requirement.
3. Booleans and Partitions expands to **Union, Intersection, Difference, Compose, Partition Objects, Partition Domains, Partition Edges**. These are explicitly disabled menu entries, not implemented Boolean operations.
4. Array creates an **Array 1 (arr1)** node and opens Settings with Input, Size, Displacement and Selections of Resulting Entities. The input list starts empty and Graphics input picking is active.
5. Rectangular defaults to x size = 1, y size = 1; Linear instead shows a single Size = 1. Switching types preserves each type's counts. Both use the shared x/y displacement fields, initially zero.
6. Pick preceding objects individually in Graphics or from the input checkboxes. Build Selected / Build All shows the generated array. Reactivate input picking to return to source objects, including objects that the built Array has consumed.
7. Contribute to / New uses the existing named cumulative selections. Materials can select individual array preview domains or reference all outputs through the cumulative selection.

## Array semantics and scope

Counts include the original position. For an input object centered at (x0, y0):

- Rectangular: (x0 + i·dx, y0 + j·dy), 0 ≤ i < nx, 0 ≤ j < ny.
- Linear: (x0 + i·dx, y0 + i·dy), 0 ≤ i < Size.

The Array consumes selected inputs and replaces them with the specified copies, including the original-position copy. Unselected objects stay in the sequence. Multiple input objects and individual outputs of preceding transforms are supported by stable entity keys. Circular radii, sectors and inward layers are retained.

Count fields accept parameter expressions resolving to positive integers. Displacement accepts the current Geometry length unit for bare numbers or explicit units such as `8[um]` and `1[mm]`. Negative and zero displacement are allowed; coincident copies are not silently discarded. Generated object counts are limited to 450 across the preview.

Array output participates in Geometry preview selection, Materials, cumulative selections, undo/redo and Geometry JSON export. It does not perform Boolean Union, produce a general 2D FEM mesh, or alter the existing radial FEM. Applying a model with an enabled Array to the radial FEM is blocked explicitly.

## Validation

- `tests/test_geometry_array.mjs`: exact rectangular and linear offsets, inclusive origin, signed/zero displacement, integer validation and size limits.
- `tests/browser_geometry_array.mjs`: menu order/state, default fields and type switching, real Graphics input toggles, generated coordinates and layers, mixed units, independent multiple/nested inputs, validation, cumulative/Material integration, persistence, undo/redo, export and FEM guard.
- Existing workbench, Rotate/Mirror, Material, Adjacent, PML and guidebook interaction regressions remain enabled.
