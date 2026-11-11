# Array output selection and Conversions menu — 2026-11-11

## User reference and implemented behavior

- In the Graphics pane, each built Array output has a stable `data-cad-entity` key and a visible input-list label such as `arr1(1)`, `arr1(2)`, etc.
- Plain click on an Array output selects exactly that object and highlights only that shape in rose, while keeping the parent Array Settings visible.
- Ctrl+click, Cmd+click, or Shift+click toggles additional outputs. The Graphics selection strip reports the selection.
- Selected objects retain their identity when the SVG redraws for geometry edits; outputs removed by a changed Build are pruned. Undo/redo snapshots include the selected output keys.
- The downstream Array, Rotate and Mirror Input pickers continue to operate on the individual stable keys; activating input picking has its pre-existing priority over ordinary output selection.
- Geometry 1 right-click menu exposes **Conversions** with **Convert to Solid**, **Convert to Curve**, **Convert to Point**, **Split**.
- Conversion commands are displayed but disabled: no conversion algorithms were specified. Existing Transforms and Booleans and Partitions are unchanged.
- This changes the 2D Geometry editor and presentation only. It does not claim additional Boolean, CAD or FEM capabilities.

## Verification checklist

1. Create Circle; add Array 1 with Linear size 5, x displacement `205*r_c`, Build Selected.
2. Click the third output: only the third circle gets the persistent rose selection (mouseover red remains temporary).
3. Ctrl+click the fifth output: third and fifth remain highlighted; Ctrl+click third again deselects only third.
4. Add Array 2, activate Input picking, and select `arr1(1)`, `arr1(3)` independently.
5. Right-click Geometry 1 → Conversions; verify the four entries and their disabled state.
6. Check Array type switch, Build All, Rotate/Mirror, and Geometry Settings.
7. Undo/redo and reload to check scene behavior. A full browser regression was not run in this change.

