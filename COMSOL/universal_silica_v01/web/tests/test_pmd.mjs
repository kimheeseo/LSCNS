import assert from 'node:assert/strict';
import {estimatePMD,maxwellCDF,maxwellPDF,maxwellQuantile} from '../pmd.mjs';
const base={pmdCoefficientPsSqrtKm:0.05,lengthKm:80,coefficientDefinition:'mean'};
const x=estimatePMD(base);
assert.ok(Math.abs(x.meanDgdPs-0.05*Math.sqrt(80))<1e-12);
assert.ok(Math.abs(maxwellCDF(x.p95DgdPs,x.maxwellScalePs)-0.95)<2e-6);
assert.ok(Math.abs(maxwellCDF(x.p99DgdPs,x.maxwellScalePs)-0.99)<2e-6);
assert.ok(x.medianDgdPs<x.p95DgdPs&&x.p95DgdPs<x.p99DgdPs);
assert.ok(maxwellPDF(x.meanDgdPs,x.maxwellScalePs)>0);
const rms=estimatePMD({...base,coefficientDefinition:'rms'});
assert.ok(Math.abs(rms.rmsDgdPs-base.pmdCoefficientPsSqrtKm*Math.sqrt(80))<1e-12);
assert.equal(estimatePMD({...base,lengthKm:0}).p95DgdPs,0);
assert.equal(maxwellQuantile(0,x.maxwellScalePs),0);
for(const bad of [{pmdCoefficientPsSqrtKm:-0.1},{lengthKm:-1},{coefficientDefinition:'unknown'}])
 assert.throws(()=>estimatePMD({...base,...bad}));
console.log('PASS: 11 PMD Maxwell-statistics, convention and bounds tests');
