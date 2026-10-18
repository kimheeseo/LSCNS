const test=require('node:test');
const assert=require('node:assert/strict');
const {calculate}=require('../engineering-v82.js');
const {summarize}=require('../ls-cable-perspective.js');

test('LS perspective separates installed trunks, spare-inclusive purchases and patch leads',()=>{
  const design=calculate({targetGPU:128});
  const summary=summarize(design);
  assert.ok(summary);
  const rows=Object.values(design.optical).filter(o=>o.links>0);
  assert.equal(summary.totalInstalledTrunks,rows.reduce((n,o)=>n+o.trunk.installedCableCount,0));
  assert.equal(summary.totalPurchasedTrunks,rows.reduce((n,o)=>n+o.trunk.cableCount,0));
  assert.ok(summary.totalPurchasedTrunks>summary.totalInstalledTrunks);
  assert.equal(summary.totalInstalledPatch,design.bom.filter(x=>x.item.includes('optical patch leads')).reduce((n,x)=>n+x.installedQty,0));
  assert.ok(summary.issues.some(x=>x.includes('SKU 미선정')));
});

test('named module specifications and point-to-point routes are shown without inventing trunks',()=>{
  const design=calculate({targetGPU:128,serverOpticSku:'nvidia-mma4z00-ns400',serverDistanceM:30,
    serverCabling:'p2p',mpoEndGender:'female',marginDb:0.5});
  const summary=summarize(design);
  const server=summary.rows.find(x=>x.segment==='server');
  assert.equal(server.opticSku,'NVIDIA MMA4Z00-NS400');
  assert.equal(server.connector,'MPO-12');
  assert.equal(server.polish,'APC');
  assert.equal(server.cableMode,'p2p');
  assert.equal(server.installedTrunks,0);
  assert.ok(!summary.issues.some(x=>x.startsWith('서버–Leaf: 광모듈 SKU 미선정')));
  assert.equal(summarize({usable:false}),null);
});
