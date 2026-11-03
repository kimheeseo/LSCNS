import assert from 'node:assert/strict';
import {estimateBending} from '../bending_loss.mjs';
const data={radiusSmallMm:10,radiusLargeMm:20,
 lossSmall1550DbTurn:0.15,lossLarge1550DbTurn:0.005,
 lossSmall1625DbTurn:0.32,lossLarge1625DbTurn:0.015,
 radiusMm:15,wavelengthNm:1550,turns:10,straightLinkLossDb:13.25};
const result=estimateBending(data);
assert.ok(Math.abs(estimateBending({...data,radiusMm:10}).lossPerTurnDb-0.15)<1e-12);
assert.ok(Math.abs(estimateBending({...data,radiusMm:20}).lossPerTurnDb-0.005)<1e-12);
assert.ok(Math.abs(result.lossPerTurnDb-0.015536162529769288)<1e-10);
assert.ok(Math.abs(result.bendLossDb-10*result.lossPerTurnDb)<1e-10);
assert.ok(Math.abs(result.totalLinkLossDb-result.bendLossDb-13.25)<1e-10);
assert.ok(Math.abs(result.bendLengthM-2*Math.PI*0.015*10)<1e-10);
assert.ok(estimateBending({...data,wavelengthNm:1625}).lossPerTurnDb>result.lossPerTurnDb);
for(const bad of [{radiusMm:9},{radiusMm:21},{wavelengthNm:1650},
 {lossSmall1550DbTurn:0},{turns:0.5},{radiusLargeMm:8}])
 assert.throws(()=>estimateBending({...data,...bad}));
console.log('PASS: 13 macrobend measured-anchor and bounds tests');
