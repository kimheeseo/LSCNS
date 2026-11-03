import assert from 'node:assert/strict';
import {estimateAttenuation} from '../attenuation.mjs';
const base={alpha1550:0.160,alpha1625:0.175,wavelengthNm:1550,lengthKm:80,
 launchDbm:0,connectorCount:2,connectorDb:0.2,spliceCount:1,spliceDb:0.05};
const x=estimateAttenuation(base);
assert.ok(Math.abs(x.alphaDbKm-0.160)<1e-12);
assert.ok(Math.abs(x.fiberLossDb-12.8)<1e-9);
assert.ok(Math.abs(x.totalLossDb-13.25)<1e-9);
assert.ok(Math.abs(x.outputDbm+13.25)<1e-9);
const midpoint=estimateAttenuation({...base,wavelengthNm:1587.5,lengthKm:100});
assert.ok(Math.abs(midpoint.alphaDbKm-0.1675)<1e-12);
assert.equal(estimateAttenuation({...base,lengthKm:0,connectorCount:0,spliceCount:0}).totalLossDb,0);
for(const invalid of [{wavelengthNm:1640},{alpha1550:-1},{connectorCount:1.5},{lengthKm:-1}])
 assert.throws(()=>estimateAttenuation({...base,...invalid}));
console.log('PASS: 10 attenuation estimation / bounds checks');
