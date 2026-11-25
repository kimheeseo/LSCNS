/* Validate supplied assets without making network requests or changing model data. */
const fs=require('node:fs'),path=require('node:path'),assert=require('node:assert/strict');
const root=path.resolve(__dirname,'..'),Core=require('../visual-hq-core.js');
const directory=path.join(root,'assets/visual-hq');
const manifest=JSON.parse(fs.readFileSync(path.join(directory,'manifest.json'),'utf8'));
assert.equal(manifest.schemaVersion,1);assert.ok(Array.isArray(manifest.assets));
const entries=Core.normalize(manifest),ready=manifest.assets.filter(a=>a.status==='ready');
assert.equal(entries.length,ready.length,'Invalid/duplicate ready asset: check identity, dimensions, paths, views and alpha');
for(const entry of entries){
  assert.match(entry.composition,/^hq1-[0-9a-f]{8}$/,'Capture the actual composition from the tool');
  for(const key of ['thumbnail','preview','full']){
    const file=path.join(directory,entry[key].src);assert.ok(fs.existsSync(file),'Missing '+entry[key].src);
    assert.ok(fs.statSync(file).size>0,'Empty '+entry[key].src);
    const bytes=fs.readFileSync(file);
    const png=bytes.subarray(0,8).equals(Buffer.from([137,80,78,71,13,10,26,10]));
    const webp=bytes.toString('ascii',0,4)==='RIFF'&&bytes.toString('ascii',8,12)==='WEBP';
    assert.ok(png||webp,'Not a PNG/WebP image: '+entry[key].src);
    if(png){assert.equal(bytes.readUInt32BE(16),entry[key].width);assert.equal(bytes.readUInt32BE(20),entry[key].height);}
    assert.ok(entry[key].width<=entry.width,'Rendition exceeds master width');
  }
}
console.log(JSON.stringify({ready:entries.length,pending:JSON.parse(fs.readFileSync(path.join(directory,'pending-targets.json'))).targets.length,
  missingImages:0,mode:entries.length?'exact-match imagery':'existing SVG fallback'},null,2));
