'use strict';
(()=>{
  const CACHE=new Map();
  const P='./product_catalog/';
  const MAP={
    trunk:[
      P+'Corning/Trunk/EDGE™ Armored Trunk/catalog.json',
      P+'YOFC/Trunk/MPO-MTP Pre-terminated Trunk Cable/catalog.json',
      P+'Sumitomo Electric/SWK™ Series/SWK Cable Assemblies/catalog.json',
      P+'Sumitomo Electric/Cable Assemblies/Cable Assemblies/catalog.json'
    ],
    patch:[
      P+'Corning/Housing/EDGE™ Housing, FX/catalog.json',
      P+'Corning/Housing/Pretium® Connector Housing (PCH)/catalog.json',
      P+'YOFC/Panel/G4 Fixed Type Fibre Optic Patch Panel/catalog.json',
      P+'ZTT/Data Center Infrastructure/High-density ODF/catalog.json',
      P+'Furukawa Electric/Optical Connectivity/Patch Panel & ODF/catalog.json',
      P+'Sumitomo Electric/Wall Mount Enclosures/FTWM-01L Wall Mount Enclosure/catalog.json',
      P+'Sumitomo Electric/Fiber Panels & Shelves/PrecisionFlex Pre-Stubbed Patch Panels/catalog.json',
      P+'Sumitomo Electric/Cassettes & Interconnect Panels/Interconnect Panels/catalog.json',
      P+'Sumitomo Electric/Wall Mount Enclosures/FSPWM-12T Wall Mount Enclosures/catalog.json',
      P+'Sumitomo Electric/Wall Mount Enclosures/FTWM-04L Wall Mount Enclosure/catalog.json',
      P+'Sumitomo Electric/Fiber Panels & Shelves/PrecisionFlex Pre-Terminated Patch Panels/catalog.json',
      P+'Sumitomo Electric/Fiber Panels & Shelves/Flex Patch Panels/catalog.json',
      P+'Sumitomo Electric/Fiber Panels & Shelves/PrecisionFlex Empty Patch Panels/catalog.json',
      P+'Sumitomo Electric/Fiber Panels & Shelves/1RU LGX Compact Patch Panel/catalog.json',
      P+'Sumitomo Electric/Fiber Panels & Shelves/2RU Flush Mount Panels/catalog.json',
      P+'Sumitomo Electric/Fiber Panels & Shelves/2RU High Density Panels and Interconnect Panels/catalog.json',
      P+'Sumitomo Electric/Fiber Panels & Shelves/PrecisionFlex High Density MPO-LC Cassettes/catalog.json',
      P+'Sumitomo Electric/Cassettes & Interconnect Panels/PrecisionFlex FOX Splice Cassettes/catalog.json',
      P+'Sumitomo Electric/Cassettes & Interconnect Panels/PrecisionFlex LGX MPO Cassettes/catalog.json',
      P+'Sumitomo Electric/Wall Mount Enclosures/FTWM-02L-2D Wall Mount Enclosure/catalog.json',
      P+'Sumitomo Electric/Wall Mount Enclosures/FTWM-04L-2D Wall Mount Enclosure/catalog.json',
    ],
    connector:[
      P+'USConnec/MTP Connectors/MTP Universal/catalog.json',
      P+'USConnec/MTP Connectors/MTP PRO/catalog.json',
      P+'USConnec/MTP Connectors/MTP PRO X/catalog.json',
      P+'USConnec/MTP Connectors/MTP-16/catalog.json',
      P+'USConnec/MTP Connectors/MTP 900µ & Fanout/catalog.json',
      P+'USConnec/MTP Connectors/Fast-Track & Pin Clamp/catalog.json'
    ],
    adapter:[P+'USConnec/MTP Connectors/Adapters/catalog.json'],
    rack:[P+'ZTT/Data Center Infrastructure/IT Cabinet/catalog.json']
  };
  const OPTIC={
    400:P+'YOFC/Transceiver/400G/catalog.json',
    200:P+'YOFC/Transceiver/200G/catalog.json',
    100:P+'YOFC/Transceiver/100G/catalog.json',
    50:P+'YOFC/Transceiver/50G/catalog.json',
    40:P+'YOFC/Transceiver/40G/catalog.json',
    25:P+'YOFC/Transceiver/25G/catalog.json',
    10:P+'YOFC/Transceiver/10G SFP+/catalog.json',
    2.5:P+'YOFC/Transceiver/2.5G/catalog.json',
    1.25:P+'YOFC/Transceiver/1.25G/catalog.json'
  };
  async function load(path){
    if(CACHE.has(path))return CACHE.get(path);
    const promise=fetch(new URL(path,location.href),{cache:'no-store'}).then(r=>{if(!r.ok)throw new Error(path+' HTTP '+r.status);return r.json()});
    CACHE.set(path,promise);return promise;
  }
  function merged(manifest,meta){return {...(manifest.defaultSpecs||{}),...(meta.specs||{})}}
  function flatten(manifest,path){return Object.entries(manifest.products||{}).map(([key,meta])=>({path,manifest,key,meta,specs:merged(manifest,meta)}))}
  async function families(paths){const out=[];for(const path of paths){try{out.push(...flatten(await load(path),path))}catch(e){}}return out}
  function val(c,...keys){for(const k of keys)if(c.specs[k]!=null&&c.specs[k]!=='—')return String(c.specs[k]);return''}
  function num(v){const m=String(v||'').replace(/,/g,'').match(/\d+(?:\.\d+)?/);return m?Number(m[0]):NaN}
  function reachM(v){const s=String(v||'').toLowerCase();const all=[...s.matchAll(/(\d+(?:\.\d+)?)\s*(km|m)/g)].map(m=>Number(m[1])*(m[2]==='km'?1000:1));return all.length?Math.max(...all):NaN}
  function fc(c){return num(val(c,'Fiber Count','Fiber count','Fiber Count / Capacity'))}
  function model(c){return c.meta.name||c.key}
  function url(c){return c.meta.officialUrl||c.manifest.officialUrl||''}
  function alt(c){return c.manifest.company+' '+model(c)}
  function toRow(row,c,alts,fit,evidence,qty=row.qty){
    return {category:row.category,vendor:c.manifest.company,product:c.manifest.category,model:model(c),qty,fit,evidence,source:url(c),alternatives:alts.map(alt),catalogBased:true,catalogPath:c.path};
  }
  function fallback(row,reason,alts=[]){
    return {...row,evidence:[row.evidence,reason].filter(Boolean).join(' · '),alternatives:alts.slice(0,2).map(alt),catalogBased:false};
  }
  function token(media){const m=String(media||'').match(/\b(SR\d*|DR\d*|FR\d*|LR\d*|ER\d*|ZR\d*|PSM\d*|CWDM\d*)\b/i);return m?m[1].toUpperCase():''}
  function connectorScore(expected,cand){
    const e=String(expected||'').toUpperCase(),c=String(cand||'').toUpperCase();
    if(!e||!c)return 0;
    if(/MPO/.test(e))return /MPO/.test(c)?(/MPO-16/.test(e)&&/MPO-16/.test(c)?35:/MPO-12/.test(e)&&/MPO-12/.test(c)?35:20):-80;
    if(/LC/.test(e))return /LC/.test(c)?30:-80;
    return 0;
  }
  async function matchOptic(row,r){
    const speed=Number(r.systemProfile?.linkSpeed);
    const path=OPTIC[speed];if(!path)return fallback(row,'product_catalog: '+speed+'G optical transceiver family not registered');
    const candidates=await families([path]);
    const server=/Server-facing/i.test(row.category);
    const o=server?r.optical?.server:r.optical?.leafSpine;
    const distance=Number(server?r.input?.serverDistanceM:r.input?.leafSpineDistanceM)||0;
    const expected=o?.profile?.connector||'',media=o?.profile?.media||'',t=token(media);
    const scored=candidates.map(c=>{
      const remark=val(c,'Remark'),reach=reachM(val(c,'Reach')),rate=num(val(c,'Data Rate')),conn=val(c,'Connector');
      if(Number.isFinite(reach)&&reach<distance)return null;
      let score=0;if(rate===speed)score+=30;if(t&&remark.toUpperCase().includes(t))score+=100;score+=connectorScore(expected,conn);
      if(Number.isFinite(reach))score-=Math.min(40,Math.max(0,(reach-distance)/1000));
      return {c,score,reach};
    }).filter(Boolean).sort((a,b)=>b.score-a.score||a.reach-b.reach);
    if(!scored.length)return fallback(row,'product_catalog: no transceiver meets distance/connector constraints');
    const primary=scored[0];
    if(primary.score<20)return fallback(row,'product_catalog: no sufficiently compatible transceiver',scored.slice(0,2).map(x=>x.c));
    return toRow(row,primary.c,scored.slice(1,3).map(x=>x.c),'Catalog match',`product_catalog · ${media} · ${distance} m`);
  }
  async function matchTrunk(row,r){
    const candidates=await families(MAP.trunk),need=Number(r.input?.trunkFiberCount)||0;
    const exact=candidates.filter(c=>fc(c)===need);
    const expected=String(r.optical?.server?.profile?.fiberType||r.optical?.leafSpine?.profile?.fiberType||'OS2').toUpperCase();
    const rank=c=>{const f=val(c,'Fiber Category','Fiber category').toUpperCase();let s=0;if(expected.includes('OS2')&&(f.includes('OS2')||f.includes('G.657')||f.includes('SINGLE')))s+=30;return s};
    exact.sort((a,b)=>rank(b)-rank(a));
    if(exact.length){const configurable=/yes/i.test(val(exact[0],'Configurable'));return toRow(row,exact[0],exact.slice(1,3),configurable?'Configurable catalog family match · RFQ':'Exact fiber-count catalog match',`product_catalog · ${need}F · quantity retained from segment calculation${configurable?' · exact termination/polarity/length code requires RFQ':''}`);}
    const near=candidates.filter(c=>Number.isFinite(fc(c))&&fc(c)>=need).sort((a,b)=>fc(a)-fc(b));
    return fallback(row,`product_catalog: no exact ${need}F structured-trunk SKU; engine quantity retained pending SKU review`,near.slice(0,2));
  }
  async function matchPatch(row){
    const candidates=await families(MAP.patch);
    const preferred=['EDGE-01U-EMOD','FT01L03-LSA','FT02FMFP-6LGX-INTERCONNECT','FT02SEL12-S','PFCST-1U-F12-BK','PCH-01U','iCONEC-DFGDSL01','GPX02-600Y1'];
    candidates.sort((a,b)=>preferred.indexOf(a.key)<0?1:preferred.indexOf(b.key)<0?-1:preferred.indexOf(a.key)-preferred.indexOf(b.key));
    if(!candidates.length)return fallback(row,'product_catalog: patch-panel/housing family unavailable');
    return toRow(row,candidates[0],candidates.slice(1,3),'Planning catalog candidate','product_catalog · exact panel capacity/topology requires project review');
  }
  function connectorBase(profile){
    const c=String(profile?.connector||'').toUpperCase();
    if(/MPO-16/.test(c))return 16;
    if(/MPO/.test(c))return 12;
    return 0;
  }
  async function matchConnector(row,r){
    const candidates=await families(MAP.connector);
    if(!candidates.length)return fallback(row,'product_catalog: MTP connector families unavailable');
    const profile=r.optical?.[row.segment]?.profile||r.optical?.server?.profile||r.optical?.leafSpine?.profile||{};
    const base=connectorBase(profile),wantGender=String(r.input?.mpoGender||'review').toLowerCase();
    const scored=candidates.map(c=>{
      const style=val(c,'Connector Style','Style').toUpperCase(),fibers=val(c,'Fiber Count'),gender=val(c,'Gender').toLowerCase();
      let score=0;
      if(base===16)score+=/MPO\s*16|MPO-16/.test(style)?100:-40;
      else if(base&&/MPO/.test(style)&&!/16/.test(style))score+=55;
      if(base===8&&/(1x8|4\+4)/i.test(fibers))score+=30;
      if(base===12&&/(1x12|2x12)/i.test(fibers))score+=30;
      if(base===16&&/(1x16|2x16)/i.test(fibers))score+=30;
      if(wantGender!=='review'&&wantGender!=='n/a'&&gender.includes(wantGender))score+=15;
      return {c,score};
    }).sort((a,b)=>b.score-a.score);
    if(!scored.length)return fallback(row,'product_catalog: no MTP connector candidate');
    return toRow(row,scored[0].c,scored.slice(1,3).map(x=>x.c),'Catalog connector candidate · RFQ',`product_catalog · ${profile.connector||'MPO/MTP'} · final gender/pinning/polarity must be confirmed`);
  }
  async function matchAdapter(row,r){
    const candidates=await families(MAP.adapter);
    if(!candidates.length)return fallback(row,'product_catalog: MTP adapter family unavailable');
    candidates.sort((a,b)=>(/Standard Footprint/.test(val(a,'Adapter Style'))?-1:0)-(/Standard Footprint/.test(val(b,'Adapter Style'))?-1:0));
    return toRow(row,candidates[0],candidates.slice(1,3),'Catalog adapter candidate · RFQ','product_catalog · key orientation / mounting / color must be confirmed');
  }
  async function matchRack(row){
    const c=(await families(MAP.rack))[0];
    if(!c)return fallback(row,'product_catalog: rack family unavailable');
    return toRow(row,c,[],'Catalog family match','product_catalog · exact cabinet width/height and accessory SKU RFQ');
  }
  async function one(row,r){
    const label=[row.category,row.product,row.model,row.evidence].filter(Boolean).join(' ');
    if(/adapter/i.test(label))return matchAdapter(row,r);
    if(/connector/i.test(label))return matchConnector(row,r);
    if(/Server-facing optic|Leaf↔Spine optic/i.test(row.category))return matchOptic(row,r);
    if(/Structured cabling/i.test(row.category))return matchTrunk(row,r);
    if(/Patch panel\s*\/\s*housing/i.test(row.category))return matchPatch(row,r);
    if(/^Rack$/i.test(row.category))return matchRack(row,r);
    return fallback(row,'No mapped product_catalog family; retain engine/RFQ result');
  }
  window.DCCatalogMatch=async function(r){
    if(!r||!r.usable||!Array.isArray(r.products))return r;
    const engineProducts=r.products.map(x=>({...x,alternatives:[...(x.alternatives||[])]}));
    const out=[];for(const row of engineProducts){try{out.push(await one(row,r))}catch(e){out.push(fallback(row,'product_catalog match error: '+e.message))}}
    r.engineProducts=engineProducts;r.products=out;
    const matched=out.filter(x=>x.catalogBased).length,rfq=out.length-matched;
    r.catalogMatchSummary={matched,rfq,total:out.length,source:'product_catalog',mode:'selected families only; full DB is not loaded into the calculation view'};
    return r;
  };
})();