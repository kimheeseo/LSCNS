# LS Datacenter Campus — N-XI concept digital twin v4.8.0

This upgrade changes the 3D application only. The 2D Gold Pixel Tour and DataCenter BOM design application remain unchanged.

## Implementation stages

| Stage | Implementation |
|---|---|
| 1 Rendering/modeling | Vendored Three.js 0.160.0; PBR metal/glass, physical lighting, soft shadow, procedural reflection environment; independently modeled building walls, doors, floor, ceiling lights and trays; distinctive transformer fins/bushings, generator exhaust/grilles, tower fans, chiller vessels, pumps, cabinets and NOC monitors. |
| 2 People/security | Floor-grid A* with corner prevention and reachable start points; acceleration, bounded heading changes, personal clearance and local detours; walk → face equipment → inspect → tablet → next task. Electrical, cooling, data hall, network and NOC engineers plus two guards. Visitor reception, carried baggage/loading, conveyor, metal-detector passage, card verification and hinged entry door. |
| 3 Rack service | Separate 42U/48U concept frame, front/rear hinged doors, four individually selected rail-mounted chassis, upper ToR, storage sled, patch panel, movable cassette/lid, connectors, fibers, rear NIC ports, A/B PDU, cable manager and overhead tray. Real scene transforms drive controls. Flexible cables update when a chassis is pulled. Front/rear camera views and actual port raycast selection. |
| 4 Connections/MMR/colo | Dedicated Network/MMR building; carrier/ODF/DWDM-style transport front panels distinct from compute equipment. Three physically fenced customer bays with separate hinged access gates, tenant racks and power/fiber routes. Network, power and hydronic paths use separate colors, curved tube geometry, metadata selection and dimming of other links. |
| 5 Integration | Desktop/mobile Playwright renders real WebGL geometry and reads actual rotations/translations. Unit tests exercise navigation and model structure. Existing facility modes, export/import, micro-scale legacy inspection, fault scenario data, training overlays and BOM engine remain. Legacy response vehicles/smoke are bridged to pooled Three geometry. |

## Files

- `LS_Datacenter_3D.html`: state/input bridge, scene delegates and safe walking.
- `LS_Datacenter_3D_NXI.js` / `.css`: renderer, interactions, staff/security and workspace layout.
- `LS_Datacenter_3D_Extensions_NXI.js`: isolated copy of existing extensions with GPU rack placement inside hall corridors; original extension file is retained.
- `assets/datacenter-3d/models.js`: mesh factories and independent model components.
- `assets/datacenter-3d/navigation.js`: floor-grid A* and staff state machine.
- `assets/datacenter-3d/response.js`: pooled compatibility adapter for existing fault visuals.
- `assets/datacenter-3d/model-manifest.json`: model provenance, dimensions and assumptions.
- `vendor/three.module.js`, package metadata and MIT license: fixed, locally served renderer; no CDN runtime dependency.
- `tests/nxi-models.test.mjs`, `tests/nxi-browser.cjs`, `.github/workflows/nxi-browser.yml`: reproducible checks and screenshots.
- `LS_Datacenter_3D_NXI_Preview.html`: isolated preview used before replacing the main entry point.

## Model/animation list

Models include rack assemblies, transformer, switchgear/UPS/CDU/PDU cabinets, generator, cooling tower, chiller, pump manifold, water/fuel tanks, battery bank, busway, fire cylinders, NOC/security desks, reception, X-ray conveyor with suitcase, metal detector, card-reader hinged gate, CCTV, waiting chairs, A/B/C cages and clothed people with helmet/ID/tools/tablet. All are mesh geometry; no flat photo or SVG billboard substitutes for these objects. Procedural model modules are the model assets; GLB files are not required.

Animations include walking limbs, gradual facing, inspection/tablet postures, cooling fans, CCTV sweep, automatic building doors, rack front/rear doors, selected chassis rails, cassette and lid, flexible cable repositioning, customer access doors, carried/loaded/conveyed baggage and approved security door passage.

## Use

Start in the large 3D workspace. Click equipment to show details; choose **선택 설비 검사** for rack controls. Select a chassis before pulling it. Use **후면 보기** and **후면 도어 열기/닫기** for NIC/PDU/cable inspection. **방문자 보안 시연** runs the entry process. **패널 복귀** returns retained operating controls. Roof, cutaway, orbit/pan/zoom, walking and staff follow controls remain available. The layer menu separates the IT layer from overhead services.

## Verification and limits

The browser report and screenshots attached to the release are authoritative for the final test run. Failed intermediate runs were used to fix elapsed-time animation, A* edge starts, guard/visitor queue conflicts and wide-mode control access. They are not presented as passing runs.

These are low-poly procedural engineering **concept models**, not certified BIM or manufacturer CAD. Rack heights (42U approx. 2.18 m; 48U approx. 2.38 m), widths/depths, people, security equipment and building dimensions are explicit modeling assumptions. The four chassis demonstrate mechanics; they do not represent a complete vendor-specific RU population or bill of materials.

400G/800G paths are educational examples. 1.6T and MPO/MTP/SN/MDC/MMC variants are conceptual/RFQ metadata, not verified reach, connector compatibility, SKU or optical budget. Power/cooling values remain simulation assumptions; this application does not validate actual facility performance, security detection, access credentials or thermal/fluid behavior.

The scene represents one operating floor with IT/service layer visibility, **not a validated multi-storey BIM**. People use grid A*, not a CAD-derived NavMesh. Dynamic queues use local yielding/replanning; arbitrary user-created layouts need separate route acceptance testing. Security authorization is a scripted training sequence, not production access control. No airport photo was present in this conversation; the equipment was modeled procedurally from the requested process.

WebGL is required. Software rendering in CI is substantially slower than a dedicated GPU and is not evidence of production-device FPS. Instanced repeated parts, rack LOD, native Three frustum culling, pooled response effects and reduced shadow update rate limit work; high-detail closeups still cost more than the legacy renderer. Hardware/browser-specific acceptance remains necessary. At micro/nano scales the retained legacy renderer is used for the original educational inspection features.
