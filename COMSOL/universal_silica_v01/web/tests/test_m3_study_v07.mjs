/**
 * M3 v0.7 regression, same-grid independent Python reference.
 * Run: node tests/test_m3_study_v07.mjs
 * n=69 benchmark is computation-heavy. Console time is measured on local host.
 */
import assert from "node:assert/strict";
import {readFileSync} from "node:fs";
import {performance} from "node:perf_hooks";
import {PRESETS,solve,silicaIndex} from "../solver.mjs";
import {lp01Characteristic} from "../validation_lp01.mjs";
import {solveVectorFDFD,vectorResultJSON} from "../vector_fdfd_browser.mjs";
import {M3_PYTHON_REF,validateM3Study,estimatedTypedArrayMiB,
 meshComparisonRows,normalizeColabJSON,referenceRunMatch,studyDocument} from "../m3_study_metrics_v07.mjs";
const raw=JSON.parse(readFileSync(new URL("../../full_vector/results/vector_mode_result.json",import.meta.url),"utf8"));
const colab=normalizeColabJSON({...raw,input:{...raw.input,n_clad:M3_PYTHON_REF.n_clad}});
const ncl=silicaIndex(1550),config={profile:"step",wavelength_nm:1550,core_radius_um:5.5,
 window_halfwidth_um:22,n_core:ncl+.0034785793846155,n_clad:ncl};
validateM3Study(config,[29,41,69]);
assert.throws(()=>validateM3Study(config,[30]),/odd/);
assert.throws(()=>validateM3Study(config,[91,29]),/increase/);
assert.ok(estimatedTypedArrayMiB(69)>0);
assert.ok(referenceRunMatch(config,69,colab.input));
assert.equal(referenceRunMatch(config,51,colab.input),false);
assert.equal(referenceRunMatch({...config,n_clad:ncl+.001,n_core:config.n_core+.001},69,colab.input),false);
const sc=solve({...PRESETS.g654,clad_index_override:ncl});
const lp=lp01Characteristic({wavelength_nm:1550,core_radius_um:config.core_radius_um,
 n_core:config.n_core,n_clad:ncl});
const runs=[],start=performance.now();
for(const grid_n of [29,41,69]){
 const t=performance.now();
 const v=solveVectorFDFD({...config,grid_n},{modes:2});
 const elapsed=performance.now()-t;
 runs.push({grid_n,config:{...config,grid_n},modes:v.modes,wallMs:elapsed});
 assert.ok(v.modes.every(m=>m.residual<1e-7));
 const portable=vectorResultJSON(v);
 if(grid_n===69){
  assert.equal(portable.schema,"LSCNS_vector_mode_v1");
  assert.ok(Math.abs(v.modes[0].neff-colab.modes[0].neff)<1e-8,
   "same-grid Python/JS mode match");
  assert.ok(portable.modes[0].field_map.intensity_normalized_to_max_total.Ex.length===69);
 }
 console.log("grid",grid_n,"ms",elapsed.toFixed(1),"neff",v.modes[0].neff);
}
const rows=meshComparisonRows(runs,sc.n_eff,lp.n_eff,colab);
assert.equal(rows[0].colabComparisonStatus,"미검증 (입력/격자 상이)");
assert.equal(rows[2].colabComparisonStatus,"동일 격자·입력");
assert.ok(rows[2].deltaToColab<1e-8);
assert.equal(rows[1].convergenceStatus,"미수렴","Do not claim monotone convergence for staircase index");
const report=studyDocument({config,grids:[29,41,69],runs,scalarNeff:sc.n_eff,
 analyticNeff:lp.n_eff,colab,vectorJSON:vectorResultJSON(solveVectorFDFD({...config,grid_n:29},{modes:2}))});
assert.equal(report.schemaVersion,2);
assert.equal(report.version,"0.7.0");
assert.equal(report.storedColabReference.grid_n,69);
assert.equal(report.finalGridConvergenceConfirmed,false);
console.log("PASS M3 v0.7: LP01/scalar/Colab neff, mesh convergence correctly not claimed, cost",
 (performance.now()-start).toFixed(1),"ms");
