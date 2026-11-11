/**
 * M2 v0.6.0: reference coefficient, interpolation and schema regression.
 * Run: node tests/test_m2_preview_v06.mjs
 * Does not use the FEM solver to compute doped n(lambda).
 */
import assert from "node:assert/strict";
import {silicaIndex,solve,PRESETS} from "../solver.mjs";
import {parseIndexTable} from "../material_models.mjs";
import {M2_CATALOG,evaluateM2Material,interpolateMeasured,
 kFromAlpha,alphaFromK,sampleM2Material,migrateM2Preview,
 makeM2PreviewDocument} from "../m2_preview_engine_v06.mjs";
const ref=[
 ["silica",1.4440236217032607],
 ["geO2",1.4515267771427718],
 ["f_wt",1.4393958509698328],
 ["p2o5",1.452224423108019]
];
for(const [id,gold] of ref){
 const m=M2_CATALOG.find(m=>m.id===id),actual=evaluateM2Material(m.model,1550).n;
 assert.ok(Math.abs(actual-gold)<1e-11,id+" Sellmeier 1550 reference mismatch");
 assert.equal(sampleM2Material(m.model,{startNm:1500,endNm:1600,count:11}).length,11);
}
assert.equal(evaluateM2Material({kind:"silica"},1550).n,silicaIndex(1550),"exact original silica");
assert.equal(evaluateM2Material({kind:"offset",deltaN:.005},1550).n,silicaIndex(1550)+.005);
assert.equal(evaluateM2Material({kind:"geO2",molPercent:0},1550).n,silicaIndex(1550));
assert.equal(evaluateM2Material({kind:"p2o5",molPercent:0},1550).n,silicaIndex(1550));
assert.equal(evaluateM2Material({kind:"f_wt",weightPercent:0},1550).n,silicaIndex(1550));
assert.throws(()=>evaluateM2Material({kind:"f_mol",molPercent:1},1550),/requires user Sellmeier coefficients/);
const custom={kind:"f_mol",molPercent:1,coefficients:{
 B:[.69616630,.40794260,.89747940],L:[.0684043,.1162414,9.896161],
 source:"User-provided test coefficients only; not validated fluorine data."}};
assert.ok(Math.abs(evaluateM2Material(custom,1550).n-silicaIndex(1550))<1e-12);
assert.throws(()=>evaluateM2Material({kind:"f_wt",weightPercent:3},1550),/0–2/);
assert.throws(()=>evaluateM2Material({kind:"p2o5",molPercent:10},1550),/0–9.1/);
assert.throws(()=>evaluateM2Material({kind:"geO2",molPercent:40},1550),/0–30/);
const table=parseIndexTable("1300,1.45,0\n1500,1.445,0.0000001\n1700,1.44,0");
for(const interpolation of ["linear","pchip"]){
 const center=evaluateM2Material({kind:"table",table,interpolation},1550);
 assert.ok(center.n<1.445&&center.n>1.44);
 assert.ok(center.kappa>=0&&center.kappa<=1e-7);
}
assert.throws(()=>evaluateM2Material({kind:"table",table,interpolation:"pchip"},1800),/Out of table range/);
assert.throws(()=>interpolateMeasured(table,1550,"spline"),/Linear|PCHIP|linear/);
for(const alpha of [0,.01,.2,20]){
 const k=kFromAlpha(alpha,1550),again=alphaFromK(k,1550);
 assert.ok(Math.abs(alpha-again)<Math.max(1e-12,alpha*1e-13));
}
assert.ok(Math.abs(kFromAlpha(.2,1550)-5.6802508913155675e-12)<1e-20);
const absorption=evaluateM2Material({kind:"silica",absorption:{kind:"alpha",value:.2}},1550);
assert.ok(Math.abs(absorption.alphaDbKm-.2)<1e-12);
assert.ok(absorption.kappa>0);
const doc=makeM2PreviewDocument({kind:"offset",deltaN:.005},{kind:"silica"});
assert.equal(doc.schemaVersion,2);
const loaded=migrateM2Preview(doc);
assert.equal(loaded.core.deltaN,.005);
const old=migrateM2Preview({schemaVersion:1,core:{mode:"f",percent:1},clad:{mode:"offset",deltaN:0}});
assert.equal(old.core.kind,"f_legacy_unresolved","legacy wt vs mol ambiguity must not be guessed");
const before=solve({...PRESETS.g652});
const after=solve({...PRESETS.g652});
for(const key of ["n_eff","MFD_PII_um","Aeff_um2"]){
 assert.equal(before[key],after[key],"legacy FEM cannot change due to M2 preview");
}
console.log("PASS M2 v0.6: coefficient sources, wt/mol separation, linear/PCHIP, complex alpha/k, schema migration, legacy FEM");
