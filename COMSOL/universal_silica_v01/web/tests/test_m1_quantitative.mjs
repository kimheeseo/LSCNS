/**
 * M1 v0.5.1 independent LP01 quantitative verification and legacy regression.
 * Run: node tests/test_m1_quantitative.mjs
 */
import assert from "node:assert/strict";
import {PRESETS,solve} from "../solver.mjs";
import {M1Q_CASES,M1Q_PREUPDATE,M1Q_TOLERANCES} from "../m1q_cases.mjs";
import {m1qFixtureModel,runM1Quantitative} from "../m1q_engine.mjs";
import {m1qReportJSON,m1qReportCSV} from "../m1q_export.mjs";
const before={...PRESETS.g652};
const report=runM1Quantitative();
assert.equal(report.version,"0.5.1");
assert.equal(report.schemaVersion,1);
assert.equal(report.kind,"LSWaveM1QuantitativeVerification");
assert.equal(report.caseCount,4);
assert.equal(report.regressionCount,5);
assert.equal(report.passCount,4);
assert.equal(report.regressionPassCount,5);
assert.equal(report.passed,true);
for(const [i,c] of report.cases.entries()){
 const source=M1Q_CASES[i];
 assert.equal(c.id,source.id);
 assert.equal(c.status,"PASS");
 assert.ok(c.absError<=M1Q_TOLERANCES.absNeff,c.id+" LP01 neff tolerance");
 assert.ok(c.sciPyAbsError<=M1Q_TOLERANCES.analyticSciPy,c.id+" SciPy independent benchmark");
 assert.ok(c.femEigenResidual<=M1Q_TOLERANCES.eigenResidual,c.id+" FEM algebraic residual");
 assert.ok(c.relativePercent>=0&&c.relativePercent<.0001,c.id+" relative error");
 assert.equal(c.inputs.n_clad,source.n_clad);
 assert.equal(c.inputs.n_core,source.n_clad+source.delta_n_core);
 assert.equal(m1qFixtureModel(source).profile,"step");
 if(c.id!=="official_comsol_145")assert.equal(c.comsolStatus,"미검증");
 else assert.equal(c.comsolStatus,"반올림 범위 내");
}
assert.ok(report.cases.find(c=>c.id==="low_v").V<2.4048255577);
assert.equal(report.cases.find(c=>c.id==="different_wavelength").inputs.wavelength_nm,1310);
assert.equal(report.cases.find(c=>c.id==="requested_41").inputs.core_radius_um,4.1);
for(const item of report.regression){
 const ref=M1Q_PREUPDATE.find(x=>x.id===item.id);
 assert.equal(item.status,"PASS");
 assert.equal(item.before.neff,ref.neff);
 for(const key of ["neff","mfd","aeff","D"]){
  assert.ok(item.absError[key]<=item.tolerances[key],item.id+" "+key+" changed");
 }
}
assert.deepEqual(PRESETS.g652,before,"Running verification must not modify original presets");
const old=solve({...PRESETS.g652});
assert.ok(Math.abs(old.n_eff-M1Q_PREUPDATE[1].neff)<=M1Q_TOLERANCES.regressionNeff);
const json=JSON.parse(m1qReportJSON(report));
assert.equal(json.caseCount,4);
const csv=m1qReportCSV(report);
assert.ok(csv.startsWith("\uFEFF"));
assert.equal(csv.trim().split("\r\n").length,25,"1 header + 4 cases + 20 regression metric rows");
assert.ok(csv.includes("relative_error_percent"));
assert.ok(csv.includes("미검증"));
console.log("PASS M1 v0.5.1: 4 independent LP01 cases, 5 legacy presets / 20 metric regressions, JSON+CSV exports");
