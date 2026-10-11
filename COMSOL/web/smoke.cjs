/* Static Wave Optics Lab smoke + independent numerical oracle assertions */
'use strict';
const assert=require('node:assert/strict');
const fs=require('node:fs'),path=require('node:path');
const api=require('./app.js');
const near=(a,b,tol,label)=>{assert.ok(Number.isFinite(a),label+' finite');assert.ok(Math.abs(a-b)<=tol,label+': '+a+' vs '+b+' (tolerance '+tol+')')};
const standard=api.fiber({wavelength_um:1.55,radius_um:4.1,delta_n:.005});
near(standard.neff,1.4461047805814204,2e-7,'LP01 analytic neff compared to scipy');
near(standard.Aeff,79.29135654686326,1,'LP01 analytic Aeff');
near(standard.MFD,10.593669506635894,.12,'LP01 analytic MFD');
assert.ok(standard.singleMode,'single mode');
near(api.silica(1.55),1.444023621703261,2e-5,'Malitson fused silica');
const two=api.mcf(16);near(two.split,.00007427797622971966,1e-11,'MCF 16um FEM reference');near(two.Lmm,10.434,0.02,'coupling length');
assert.ok(api.mcf(12).split>api.mcf(20).split,'MCF coupling decreases');
const cap=api.capillary(1.55,15,1.00027);near(cap.neff,.9994878124536852,2e-9,'capillary n_eff vs Python');near(cap.D,3.373167521614031,.005,'capillary D vs Python');
for(const x of [1.3,1.55,1.65])assert.ok(api.capillary(x,15,1.00027).neff<1.00027);
assert.throws(()=>api.fiber({wavelength_um:3.5,radius_um:4.1,delta_n:.005}));
assert.throws(()=>api.mcf(25));
const html=fs.readFileSync(path.join(__dirname,'../index.html'),'utf8');
assert.match(html,/web\/app\.js/);assert.match(html,/web\/style\.css/);
for(const id of ['wl','radius','dn','pitch','hwl','hradius','air','metrics','mode-map','chart-a','chart-b','calculate','export-csv']){
 assert.ok(html.includes('id="'+id+'"'),'missing web control: '+id);
}
console.log(JSON.stringify({result:'PASS',checks:17,LP01:{neff:standard.neff,Aeff:standard.Aeff,MFD:standard.MFD,V:standard.V},MCF:two,HCF:cap},null,2));
