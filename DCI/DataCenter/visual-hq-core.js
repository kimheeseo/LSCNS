/* Presentation-only asset matching. No calculation or model mutations. */
(function(root,factory){
  const api=factory();
  if(typeof module==='object'&&module.exports)module.exports=api;
  else root.DCVisualHQCore=api;
})(typeof window==='undefined'?globalThis:window,function(){
  'use strict';
  const REVISION='20261126-visual-hq-v1';
  const views=new Set(['overview','cutaway','interior','front','rear','icon','detail']);
  function canonical(value){
    if(Array.isArray(value))return value.map(canonical);
    if(value&&typeof value==='object')return Object.fromEntries(Object.keys(value).sort().map(k=>[k,canonical(value[k])]));
    return value;
  }
  function signature(value){
    const text=JSON.stringify(canonical(value));let hash=2166136261;
    for(let i=0;i<text.length;i++)hash=Math.imul(hash^text.charCodeAt(i),16777619);
    return 'hq1-'+(hash>>>0).toString(16).padStart(8,'0');
  }
  function composition(T){
    return signature({revision:REVISION,profile:T.facilityMode,
      zones:T.visibleZones().map(z=>[z.id,z.pos,z.size,z.color]),
      assets:T.visibleAssets().map(a=>[a.id,a.zone,a.type,a.x,a.z,a.w,a.d,a.h])});
  }
  function bomComposition(result){
    return signature({revision:REVISION,input:result.input,rackLayout:result.rackLayout,
      cooling:result.cooling?.mode,system:result.systemProfile});
  }
  function rendition(value){
    if(!value||typeof value.src!=='string'||!/^media\/(?:[a-zA-Z0-9_-]+\/)*[a-zA-Z0-9_.-]+\.(?:png|webp)$/.test(value.src))return null;
    if(!Number.isInteger(value.width)||value.width<1||!Number.isInteger(value.height)||value.height<1)return null;
    return {src:value.src,width:value.width,height:value.height};
  }
  function normalize(manifest){
    if(manifest?.schemaVersion!==1||!Array.isArray(manifest.assets))return [];
    const normalized=[],keys=new Set();
    for(const entry of manifest.assets){
      if(entry.status!=='ready'||!['tour','bom'].includes(entry.scope)||!views.has(entry.view))continue;
      if(typeof entry.profile!=='string'||!entry.profile||typeof entry.composition!=='string')continue;
      if(typeof entry.alpha!=='boolean'||!Number.isInteger(entry.width)||!Number.isInteger(entry.height))continue;
      const full=rendition(entry.full),preview=rendition(entry.preview),thumbnail=rendition(entry.thumbnail);
      if(!full||!preview||!thumbnail||entry.width!==full.width||entry.height!==full.height)continue;
      if([preview,thumbnail].some(r=>Math.abs(r.width/r.height-full.width/full.height)>.01))continue;
      if(entry.scope==='tour'&&(!entry.zoneId||!['aidc','onprem','server-room'].includes(entry.profile)))continue;
      if(entry.partId&&!entry.assetId&&!entry.zoneId)continue;
      if(entry.scope==='bom'&&(!entry.diagram?.svgId||!entry.diagram?.nodeKey||entry.view!=='icon'||!entry.alpha))continue;
      if(entry.scope==='bom'&&!['rackSvg','coolingSvg','powerArchitectureSvg'].includes(entry.diagram.svgId))continue;
      const key=JSON.stringify([entry.scope,entry.profile,entry.composition,entry.zoneId,entry.assetId,entry.partId,entry.view,entry.diagram]);
      if(keys.has(key))continue;keys.add(key);
      normalized.push({...entry,full,preview,thumbnail});
    }
    return normalized;
  }
  function resolve(entries,context){
    // Exact identity at every level: never borrow a parent/other-profile picture.
    const candidates=entries.filter(e=>e.scope===context.scope&&e.profile===context.profile&&e.composition===context.composition&&
      (e.zoneId||null)===(context.zoneId||null)&&(e.assetId||null)===(context.assetId||null)&&
      (e.partId||null)===(context.partId||null)&&(!context.view||e.view===context.view)&&
      (e.diagram?.svgId||null)===(context.diagram?.svgId||null)&&
      (e.diagram?.nodeKey||null)===(context.diagram?.nodeKey||null));
    const rank={overview:0,cutaway:1,interior:2,front:3,detail:4,icon:5,rear:6};
    return candidates.sort((a,b)=>rank[a.view]-rank[b.view])[0]||null;
  }
  return {REVISION,signature,composition,bomComposition,normalize,resolve,rendition};
});
