import assert from 'node:assert/strict';
import {PRESETS} from '../solver.mjs';
import {
 VELAMURI_REFERENCE,stressOpticIndex,equivalentBendIndex,bentIndexCut,
 bendLossFromComplexNeff,imagNeffFromLoss,assessPMLConvergence
} from '../bend_physics_bridge.mjs';

let checks=0;
const close=(a,b,tol=1e-11)=>{assert.ok(Math.abs(a-b)<=tol,`got ${a} vs ${b}`);checks++;};
const paper={wavelengthNm:1550,bendDiameterMm:10,lossDbPerTurn:0.0119};
const nImag=imagNeffFromLoss(paper);
close(nImag,1.0758076087524702e-8,1e-13);
const output=bendLossFromComplexNeff({
 wavelengthNm:1550,bendDiameterMm:10,neffImagMagnitude:nImag,turns:3});
close(output.lossDbPerTurn,0.0119);
close(output.totalBendLossDb,0.0357);
close(output.bendRadiusMm,5);
const index=equivalentBendIndex({straightIndex:1.45,xUm:0,bendDiameterMm:10});
close(index.nEquivalent,1.45);
const fiber=bentIndexCut(PRESETS.g654,{wavelengthNm:1550,bendDiameterMm:10,
 claddingRadiusUm:62.5,stressMode:'none',samples:301});
assert.equal(fiber.rows.length,301);checks++;
close(fiber.rows[150].nStraight,fiber.rows[150].nEquivalent);
assert.ok(fiber.rows[0].nGeometryOnly<fiber.rows[0].nStraight);checks++;
assert.ok(fiber.rows.at(-1).nGeometryOnly>fiber.rows.at(-1).nStraight);checks++;
const elastic=bentIndexCut(PRESETS.g654,{wavelengthNm:1550,bendDiameterMm:10,
 claddingRadiusUm:62.5,stressMode:'beam_proxy',samples:301});
assert.ok(Math.abs(elastic.rows.at(-1).nEquivalent-fiber.rows.at(-1).nEquivalent)>1e-7);checks++;
assert.ok(stressOpticIndex(1.45,.001,-.00017,-.00017)>0);checks++;
const passing=assessPMLConvergence({wavelengthNm:1550,bendDiameterMm:10,
 rows:[{thicknessLambda:5,neffImagMagnitude:1.20e-8},
       {thicknessLambda:7,neffImagMagnitude:1.18e-8}],tolerancePercent:2});
assert.equal(passing.passed,true);checks++;
assert.ok(passing.relativeChangePercent<2);checks++;
const failing=assessPMLConvergence({wavelengthNm:1550,bendDiameterMm:10,
 rows:[{thicknessLambda:5,neffImagMagnitude:1.50e-8},
       {thicknessLambda:7,neffImagMagnitude:1.10e-8}],tolerancePercent:2});
assert.equal(failing.passed,false);checks++;
assert.throws(()=>bendLossFromComplexNeff({wavelengthNm:1550,bendDiameterMm:0,neffImagMagnitude:1e-8}));checks++;
assert.throws(()=>bendLossFromComplexNeff({wavelengthNm:1550,bendDiameterMm:10,neffImagMagnitude:-1e-8}));checks++;
assert.throws(()=>bentIndexCut(PRESETS.g654,{wavelengthNm:1550,bendDiameterMm:10,claddingRadiusUm:5001}));checks++;
assert.throws(()=>assessPMLConvergence({wavelengthNm:1550,bendDiameterMm:10,rows:[{thicknessLambda:7,neffImagMagnitude:1e-8}]}));checks++;
assert.equal(VELAMURI_REFERENCE.year,2018);checks++;
console.log(`PASS ${checks} Velamuri bending bridge, index profile, complex neff and external PML tests.`);
