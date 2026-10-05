const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path');
const M=require('../bom-catalog.js');
const read=folder=>M.normalize(JSON.parse(fs.readFileSync(path.join(__dirname,'../product_catalog',folder,'catalog.json'),'utf8')),folder+'/catalog.json');
const optical=read('Coherent/Optical Transceivers'),aec=read('Amphenol/Active Electrical Cables'),dac=read('Amphenol/Copper DAC');
test('Coherent Korean fields are read; 400G FR4 is not a DR4 candidate; unknown IB is not qualified',()=>{
 const r={item:'플러그형 트랜시버',transceiver:true,requirement:{speed:400,lengthM:100,standard:'DR4',connector:'MPO-12/APC',package:'QSFP-DD',protocol:'Ethernet'}};
 const found=M.match(r,optical);assert.equal(found.length,1);assert.match(found[0].name,/FTCD4533/);assert.match(found[0].status,/관련 제품/);assert.match(found[0].matchBasis,/프로토콜/);
 assert.equal(M.match({...r,requirement:{...r.requirement,speed:800}},optical).length,0);
 assert.equal(M.match({...r,requirement:{...r.requirement,lengthM:3000}},optical).length,0);
});
test('AEC exact length and PCIe protocol exclude incompatible copper cables and optical modules',()=>{
 const row={item:'연결 케이블',media:'AEC',cable:true,lengthM:3,requirement:{speed:800,package:'OSFP',protocol:'Ethernet'}};
 const found=M.match(row,[...aec,...dac,...optical]);assert.deepEqual(found.map(x=>x.id),['NJMMNK-0303']);assert.match(found[0].status,/검증 필요/);
 const pcie=M.match({item:'AEC',media:'AEC',lengthM:5,requirement:{protocol:'PCIe Gen5',package:'OSFP-XD'}},aec);assert.deepEqual(pcie.map(x=>x.id),['NEUUSH-0105']);
 const before=JSON.stringify(row);M.match(row,aec);assert.equal(JSON.stringify(row),before,'matching cannot mutate installation/purchase quantities');
});
test('UPS and rack reference rows link registered products without treating racks as servers',()=>{
 const products=[...read('Eaton/UPS'),...read('Schneider Electric/Racks'),...optical];
 assert.equal(M.match({item:'UPS'},products).length,3);assert.equal(M.match({item:'IT 랙'},products).length,2);
 assert.ok(M.match({item:'UPS'},products).every(x=>x.status==='관련 제품 · 검증 필요'));
 assert.equal(M.match({item:'서버'},products).length,0);
});
test('full catalog loader discovers new paths at each revision, reports errors and retries',async()=>{
 const original=global.fetch;let revision='one',fail=false;
 global.fetch=async url=>{let v;if(url.includes('/git/ref/'))v={object:{sha:revision}};else if(url.includes('/contents/'))v=[{name:'product_catalog',sha:revision}];else if(url.includes('/git/trees/'))v={truncated:false,tree:[{type:'blob',path:'new-vendor/new-category/catalog.json'}]};else{if(fail)return{ok:false,status:503};v={company:'Added vendor',category:'UPS',products:{new:{name:'UPS '+revision,officialUrl:'https://example.com/ups',specs:{}}}};}return{ok:true,json:async()=>v};};
 try{let first=await M.load(true);assert.equal(first.products[0].name,'UPS one');revision='two';let second=await M.load(true);assert.equal(second.products[0].name,'UPS two');assert.equal(second.revision,'two');fail=true;let failed=await M.load(true);assert.equal(failed.errors.length,1);assert.equal(failed.products.length,0);fail=false;assert.equal((await M.load(true)).errors.length,0);}finally{global.fetch=original;}
});
test('product-name data rate excludes mismatched AOCs and battery racks are not IT racks',()=>{
 const products=M.normalize({company:'Example',category:'All IT Datacom Products',products:{low:{name:'100G QSFP28 Active Optical Cable (AOC)',businessUrl:'https://example.com/aoc',officialUrl:'https://example.com/list',specs:{}},battery:{name:'Indoor Battery Racks - Bay Rack',officialUrl:'https://example.com/battery'},panel:{name:'Rack Termination Box',officialUrl:'https://example.com/panel'}}},'Example/All IT Datacom Products/catalog.json');
 assert.equal(M.match({item:'연결 케이블',media:'AOC',requirement:{speed:400}},products).length,0);
 assert.equal(M.match({item:'IT 랙'},products).length,0);assert.equal(products[0].source,'https://example.com/aoc');
 const r={item:'플러그형 트랜시버',requirement:{speed:400,lengthM:100,standard:'400GBASE-DR4',connector:'MPO-12/APC',package:'QSFP-DD',protocol:'Ethernet'}};
 assert.equal(M.match(r,optical).length,1);
 const sas=M.normalize({company:'Example',products:{sas:{name:'Mini-SAS HD Active Optical Cable (AOC)',officialUrl:'https://example.com/sas'},sub:{name:'Sub-rack SC-APC Complete',officialUrl:'https://example.com/sub'}}},'Example/catalog.json');
 assert.equal(M.match({item:'연결 케이블',media:'AOC',requirement:{speed:400,protocol:'InfiniBand'}},sas).length,0);assert.equal(M.match({item:'IT 랙'},sas).length,0);
});
