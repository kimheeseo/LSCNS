const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const Core=require('../visual-hq-core.js');
const root=path.resolve(__dirname,'..');
const image=(name,width,height)=>({src:'media/'+name+'.webp',width,height});
const context={scope:'tour',profile:'aidc',composition:'hq1-reference',zoneId:'hall',assetId:'tor-a',partId:null};
const ready={...context,status:'ready',view:'front',width:3840,height:2160,alpha:false,
  thumbnail:image('rack-480',480,270),preview:image('rack-1600',1600,900),full:image('rack-3840',3840,2160)};
const normalize=asset=>Core.normalize({schemaVersion:1,assets:[asset]});

test('an absent, pending or malformed asset never produces an image request',()=>{
  assert.deepEqual(Core.normalize(null),[]);
  assert.deepEqual(normalize({...ready,status:'awaiting-source'}),[]);
  assert.equal(Core.resolve([],context),null);
  assert.deepEqual(normalize({...ready,full:null}),[]);
});
test('identity matching excludes other profiles, zones, assets, parts and compositions',()=>{
  const entries=normalize(ready);assert.equal(Core.resolve(entries,context).assetId,'tor-a');
  for(const [key,value] of [['profile','onprem'],['zoneId','network'],['assetId','tor-b'],['partId','tor-a:cassette-1'],['composition','changed']]){
    assert.equal(Core.resolve(entries,{...context,[key]:value}),null,key);
  }
  assert.equal(Core.resolve(entries,{...context,assetId:null}),null,'no parent substitution');
});
test('part-only and zone-only assets cannot masquerade as rack images',()=>{
  const zone=normalize({...ready,assetId:null,view:'interior'});
  assert.equal(Core.resolve(zone,context),null);
  assert.ok(Core.resolve(zone,{...context,assetId:null}));
  const part=normalize({...ready,partId:'tor-a:cassette-1',view:'detail'});
  assert.equal(Core.resolve(part,context),null);
  assert.ok(Core.resolve(part,{...context,partId:'tor-a:cassette-1'}));
});
test('asset paths are local PNG/WebP files below the media directory',()=>{
  for(const src of ['https://example.com/x.webp','../x.webp','media/../x.webp','media/x.svg','media/x.webp?x=1','data:image/png;base64,test']){
    assert.equal(Core.rendition({src,width:100,height:100}),null,src);
  }
  assert.ok(Core.rendition(image('one',256,256)));
});
test('renditions require explicit dimensions and a consistent aspect ratio',()=>{
  assert.deepEqual(normalize({...ready,width:100}),[]);
  assert.deepEqual(normalize({...ready,preview:image('bad',1600,1600)}),[]);
  assert.deepEqual(normalize({...ready,alpha:'true'}),[]);
});
test('BOM imagery is restricted to exact transparent icon slots',()=>{
  const bom={...ready,scope:'bom',profile:'b200',zoneId:null,assetId:null,partId:null,view:'icon',alpha:true,
    diagram:{svgId:'rackSvg',nodeKey:'B200 8-GPU planning node|1-10'}};
  assert.equal(normalize(bom).length,1);
  assert.deepEqual(normalize({...bom,alpha:false}),[]);
  assert.deepEqual(normalize({...bom,diagram:{svgId:'topologySvg',nodeKey:'x'}}),[]);
  assert.equal(Core.resolve(normalize(bom),{...bom,diagram:{...bom.diagram,nodeKey:'other'}}),null);
});
test('signatures are stable across object-key order but change with actual geometry/input',()=>{
  assert.equal(Core.signature({a:1,b:2}),Core.signature({b:2,a:1}));
  const T={facilityMode:'aidc',visibleZones:()=>[{id:'hall',pos:[0,0],size:[30,40],color:'#123456'}],
    visibleAssets:()=>[{id:'tor-a',zone:'hall',x:1,z:2,w:3,d:4,h:5,type:'rack'}]};
  const before=Core.composition(T);T.facilityMode='onprem';assert.notEqual(before,Core.composition(T));
  T.facilityMode='aidc';T.visibleAssets=()=>[{id:'tor-a',zone:'hall',x:99,z:2,w:3,d:4,h:5,type:'rack'}];
  assert.notEqual(before,Core.composition(T),'same-profile geometry changes invalidate imagery');
  assert.notEqual(Core.bomComposition({input:{targetGPU:256}}),Core.bomComposition({input:{targetGPU:512}}));
});
test('pending mapping IDs belong to the host model and never request nonexistent media',()=>{
  const html=fs.readFileSync(path.join(root,'LS_Datacenter_3D.html'),'utf8');
  const source=html.slice(html.indexOf('const Z=['),html.indexOf('let rackIsolateActive=false;'));
  const model=vm.runInNewContext(source+';({Z,assets})');
  const pending=JSON.parse(fs.readFileSync(path.join(root,'assets/visual-hq/pending-targets.json')));
  for(const target of pending.targets){
    assert.ok(model.Z.some(z=>z.id===target.zoneId));
    if(target.assetId)assert.ok(model.assets.some(a=>a.id===target.assetId&&a.zone===target.zoneId));
    assert.equal(target.status,'awaiting-source');assert.equal(target.preview,null);
  }
  const manifest=JSON.parse(fs.readFileSync(path.join(root,'assets/visual-hq/manifest.json')));
  assert.equal(Core.normalize(manifest).length,manifest.assets.filter(a=>a.status==='ready').length);
});
