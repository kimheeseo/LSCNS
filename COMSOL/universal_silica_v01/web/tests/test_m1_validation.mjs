/**
 * M1 independent oracle + legacy numerical regression.
 * Run: node tests/test_m1_validation.mjs
 * Also run existing: node tests/test_solver.mjs
 */
import assert from "node:assert/strict";
import {PRESETS,solve} from "../solver.mjs";
import {lp01Characteristic} from "../validation_lp01.mjs";
import {COMSOL_PUBLIC,LP01_CASES,TOL} from "../validation_cases.mjs";
import {fixtureModel,validateMesh,runM1Validation} from "../validation_engine.mjs";

assert.equal(COMSOL_PUBLIC.reference_neff_rounded,1.4444);
assert.equal(COMSOL_PUBLIC.core_radius_um,8);
assert.equal(COMSOL_PUBLIC.outer_radius_um,40);
assert.equal(COMSOL_PUBLIC.n_core,1.4457);
assert.equal(COMSOL_PUBLIC.n_clad,1.4378);
for(const c of LP01_CASES){
 const analytic=lp01Characteristic(c);
 assert.ok(Math.abs(analytic.n_eff-c.expectedAnalytic)<=TOL.analyticVsSciPy,c.id+" analytic vs independently computed SciPy");
 assert.ok(Math.abs(analytic.V-c.expectedV)<1e-11,c.id+" V");
 assert.ok(Math.abs(analytic.residual)<1e-9,c.id+" characteristic-equation residual");
 assert.ok(analytic.n_eff>c.n_clad&&analytic.n_eff<c.n_core,c.id+" guided interval");
 const m=fixtureModel(c);
 const fem=solve(m);
 assert.ok(Math.abs(fem.n_eff-analytic.n_eff)<=TOL.lp01Neff,c.id+" FEM vs LP01");
 assert.ok(fem.residual<=TOL.residual,c.id+" FEM relative residual");
 assert.equal(PRESETS.comsol_step.elements,750,"Presets must not be mutated");
}
assert.throws(()=>lp01Characteristic({...LP01_CASES[0],n_core:1.43}),/Core index/);
assert.throws(()=>lp01Characteristic({...LP01_CASES[0],core_radius_um:0}),/positive/);
const mesh=validateMesh();
assert.ok(mesh.rows[0].passed);
assert.equal(mesh.meshes.length,4);
for(let i=1;i<mesh.meshes.length;i++)
 assert.ok(mesh.meshes[i].analyticAbsError<mesh.meshes[i-1].analyticAbsError,"Mesh accuracy should improve");
const result=await runM1Validation();
assert.equal(result.schemaVersion,1);
assert.equal(result.kind,"M1-validation-report");
assert.equal(result.totalCount,28);
assert.equal(result.passedCount,28);
assert.equal(result.passed,true);
assert.ok(result.rows.some(row=>row.group==="COMSOL 공개값"&&row.passed));
assert.ok(result.rows.some(row=>row.group==="FEM vs LP01"&&row.passed));
assert.equal(result.rows.filter(row=>row.group==="기존 회귀").length,20);
const comsolFem=solve({...PRESETS.comsol_step});
assert.ok(Math.abs(comsolFem.n_eff-COMSOL_PUBLIC.reference_neff_rounded)<=TOL.comsolRoundedNeff);
assert.equal(result.details[0].femNeff,comsolFem.n_eff);
console.log("PASS M1: 28/28 numeric comparisons, independent LP01 oracle, COMSOL rounding, legacy baselines and mesh trend.");
