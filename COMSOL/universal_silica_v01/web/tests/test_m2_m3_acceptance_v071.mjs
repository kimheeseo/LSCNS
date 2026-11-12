/**
 * Acceptance test for existing M2 preview and STRICT M3 Yee solver.
 * Run from web/: node tests/test_m2_m3_acceptance_v071.mjs
 * Only verifies M2/M3; intentionally does not exercise M4 PML/HCF/M5/M6.
 */
import assert from "node:assert/strict";
import {performance} from "node:perf_hooks";
import {PRESETS,solve,silicaIndex} from "../solver.mjs";
import {runM1Quantitative} from "../m1q_engine.mjs";
import {M2_CATALOG,M2_SOURCES,evaluateM2Material,interpolateMeasured,
 kFromAlpha,alphaFromK} from "../m2_preview_engine_v06.mjs";
import {parseIndexTable} from "../material_models.mjs";
import {solveYee,portableYee} from "../yee_solver.mjs";
const m1=runM1Quantitative();
assert.equal(m1.passCount,4,"M1 4 analytic cases must remain PASS");
assert.equal(m1.regressionPassCount,5,"M1 original five FEM presets must remain unchanged");
const pure=evaluateM2Material({kind:"silica"},1550);
assert.equal(pure.n,silicaIndex(1550),"SiO2 must exactly reuse original Sellmeier");
assert.ok(M2_SOURCES.silica.url.includes("JOSA.55.001205"),"Malitson primary reference required");
const values={};
for(const entry of M2_CATALOG){
 const n=evaluateM2Material(entry.model,1550).n;
 assert.ok(n>1&&n<3);
 values[entry.id]=n;
}
assert.ok(values.geO2>values.silica);
assert.ok(values.f_wt<values.silica);
assert.ok(values.p2o5>values.silica);
assert.throws(()=>evaluateM2Material({kind:"f_mol",molPercent:1},1550),/requires user Sellmeier coefficients/);
const rows=parseIndexTable("1400,1.45,0\n1500,1.448,0.00000001\n1600,1.446,0");
assert.equal(interpolateMeasured(rows,1550,"linear").n,1.447);
assert.ok(Math.abs(interpolateMeasured(rows,1550,"pchip").n-1.447)<1e-12);
assert.throws(()=>interpolateMeasured(rows,1650,"linear"),/Out of table range/);
const k=kFromAlpha(0.2,1550);
assert.ok(Math.abs(k-5.6802508913155675e-12)<1e-20);
assert.ok(Math.abs(alphaFromK(k,1550)-.2)<1e-12);

const ncl=silicaIndex(1550),cfg={
 wavelength_nm:1550,core_radius_um:4.1,n_core:ncl+.005,n_clad:ncl,
 window_halfwidth_um:18,grid_n:24,modes:2,geometry:{kind:"radial"},
 pml:{enabled:false,thickness_um:4,strength:3,order:3},
 bend_radius_um:0,k_core:0,k_clad:0,subpixel:true
};
const scalar=solve({...PRESETS.g652,clad_index_override:ncl});
const before=process.memoryUsage();
const t0=performance.now();
const result=solveYee(cfg);
const wallMs=performance.now()-t0,after=process.memoryUsage();
const portable=portableYee(result);
assert.equal(portable.schema,"LSCNS_vector_mode_v1");
assert.equal(result.modes.length,2);
assert.ok(result.modes.every(m=>Number.isFinite(m.neff)&&m.residual<1e-8));
assert.ok(Math.abs(result.modes[0].neff-scalar.n_eff)<5e-4,
 "At coarse N=24, the genuine Yee vector neff should be near scalar LP01, not equal.");
assert.ok(portable.modes.every(m=>m.field_map?.intensity_normalized_to_max_total?.Ez));
assert.ok(result.modes.every(m=>m.Aeff_um2>0&&m.MFD_2nd_um>0));
assert.equal(portable.input.pml.enabled,false);
console.log(JSON.stringify({
 type:"M2M3 v0.7.1 acceptance",
 m1:"4/4 cases + 5/5 original presets PASS",
 m2_1550:values,alpha_0p2_kappa:k,
 yee_n:24,yee_neff:result.modes.map(m=>m.neff),
 scalar_neff:scalar.n_eff,yee_residuals:result.modes.map(m=>m.residual),
 measured_wall_ms:wallMs,process_rss_before_mib:before.rss/1048576,
 process_rss_after_mib:after.rss/1048576,
 note:"Process RSS is measured local Node total, not a browser Worker memory measurement.",
 grid_convergence:"NOT certified by a single N24 run; see existing N24/48/72/96 table"
},null,2));
