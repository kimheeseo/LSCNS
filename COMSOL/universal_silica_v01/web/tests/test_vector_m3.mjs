/**
 * M3 browser vector Maxwell operator regression against the *existing* Python
 * full_vector/results/vector_mode_result.json baseline (same 69x69 mesh).
 * Run with Node JS, independent of Colab and network.
 */
import assert from "node:assert/strict";
import {silicaIndex} from "../solver.mjs";
import {solveVectorFDFD,makeVectorGrid,vectorResultJSON} from "../vector_fdfd_browser.mjs";
const ncl=silicaIndex(1550),dn=.0034785793846155;
const input={profile:"step",wavelength_nm:1550,core_radius_um:5.5,n_clad:ncl,n_core:ncl+dn,
 window_halfwidth_um:22,grid_n:69};
assert.equal(makeVectorGrid(input).index.length,69*69);
const result=solveVectorFDFD(input,{modes:2});
assert.equal(result.modes.length,2);
const PYTHON_NEFF=1.445711888509771; // Existing Python 2D H-field FDFD result, same grid/window/profile.
for(const m of result.modes){
 assert.ok(m.neff>ncl&&m.neff<ncl+dn);
 assert.ok(Math.abs(m.neff-PYTHON_NEFF)<2e-6,"FDFD agreement with existing Python same-grid eigs");
 assert.ok(m.residual<2.5e-5,"Eigenpair relative residual");
 assert.ok(m.Ez_intensity_fraction>0&&m.Ez_intensity_fraction<.05);
 assert.ok(Math.abs(m.Ex_intensity_fraction+m.Ey_intensity_fraction+m.Ez_intensity_fraction-1)<1e-10);
 assert.ok(m.Aeff_um2>0&&m.MFD_2nd_um>0);
 assert.ok(m.edge_power_fraction<.01);
}
assert.ok(result.modes[0].polarization!==result.modes[1].polarization,"Polarization partners distinct");
assert.ok(Math.abs(result.modes[0].neff-result.modes[1].neff)<5e-5,"Circular fiber pair near-degenerate");
const portable=vectorResultJSON(result);
assert.equal(portable.schema,"LSCNS_vector_mode_v1");
assert.equal(portable.schemaVersion,1);
assert.ok(Array.isArray(portable.modes[0].field_map.intensity_normalized_to_max_total.Ex));
assert.ok(portable.modes[0].field_map.intensity_normalized_to_max_total.Ez.length===69);
assert.throws(()=>makeVectorGrid({...input,grid_n:30}),/odd/);
assert.throws(()=>makeVectorGrid({...input,n_core:ncl-.001}),/guided/);
console.log("PASS M3: real 2D coupled Maxwell Hx-Hy FDFD, Ex/Ey/Ez, two polarizations and Python same-grid neff");
