const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path');
const M=require('../bom-catalog.js'),E=require('../engineering-v82.js');
const base=path.join(__dirname,'../product_catalog');
const manifest=JSON.parse(fs.readFileSync(path.join(base,'catalog-manifest.json'),'utf8'));
const catalogs=manifest.paths.map(p=>[p,JSON.parse(fs.readFileSync(path.join(base,p),'utf8'))]);
const products=catalogs.flatMap(([p,c])=>M.normalize(c,p));
test('catalog records preserve old IDs, explicit nulls and field evidence',()=>{
 const ids=new Set();
 for(const [p,c] of catalogs)for(const [id,x] of Object.entries(c.products||{})){
  assert.ok(!ids.has(p+'#'+id));ids.add(p+'#'+id);
  for(const k of ['source_url','verified_date','spec_basis','status','deprecated'])assert.ok(Object.hasOwn(x,k),p+'#'+id+' '+k);
  assert.ok(['active','announced','EOL','RFQ'].includes(x.status));
  if(x.status==='active'||x.status==='announced'){
   assert.ok(x.source_url&&x.verified_date);
   for(const [k,v] of Object.entries(x.specs||{}))if(v!=null)assert.ok(x.spec_sources[k],p+'#'+id+' '+k);
  }
 }
});
test('RFQ never recovers numeric specifications from model names or archived data',()=>{
 const c={company:'Test',products:{x:{name:'800G OSFP DR8 transceiver',officialUrl:'https://example.org/x',spec_basis:null,status:'RFQ',verified_date:null,source_url:null,specs:{Speed:null},legacy_specs:{Speed:'800G'}}}};
 const [p]=M.normalize(c,'Test/Transceiver/catalog.json');
 assert.equal(p.speed,'');assert.equal(p.package,'');assert.equal(p.standard,'');
 const found=M.match({item:'transceiver',requirement:{speed:800}},[p]);
 assert.ok(found.length);assert.match(found[0].status,/RFQ/);assert.equal(found[0].evidence,'품목 분류 일치');
});
test('EOL and announced products are not BOM purchasing candidates',()=>{
 for(const p of products.filter(p=>['EOL','announced'].includes(p.catalog_status)))assert.equal(M.match({item:p.name},[p]).length,0);
 const eol=catalogs.find(([p])=>p.includes('Discontinued Products'))[1];
 assert.ok(Object.values(eol.products).every(p=>p.deprecated&&p.status==='EOL'));
});
test('default B200 1024 GPU receipt has unchanged quantities and real verified candidates',()=>{
 const d=E.calculate({systemId:'b200',targetGPU:1024}),before=JSON.stringify(d);
 assert.equal(d.summary.systemUnits,128);assert.equal(d.summary.targetGPU,1024);assert.equal(d.bom.length,32);assert.equal(d.errors.length,0);
 const rows=d.productRequirements.map(r=>M.match(r,products));
 assert.ok(rows.some(r=>r.length>0));assert.ok(rows.some(r=>r.some(p=>p.catalog_status==='active')));
 assert.equal(rows[0][0].name,'NVIDIA DGX B200');
 assert.equal(JSON.stringify(d),before);
});
test('known protocol, connector and reach conflicts stay excluded',()=>{
 const p=products.find(p=>p.id==='MMA4Z00-NS400'&&p.catalog_status==='active');
 assert.equal(M.match({item:'transceiver',requirement:{speed:800}},[p]).length,0);
 assert.equal(M.match({item:'transceiver',requirement:{connector:'MPO-12/UPC'}},[p]).length,0);
 assert.equal(M.match({item:'transceiver',lengthM:60},[p]).length,0);
});
