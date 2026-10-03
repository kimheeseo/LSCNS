/* Product-family matching is an RFQ shortlist, never a platform qualification. */
(() => {
 'use strict';
 const baseKey={server:'serverBase',leafSpine:'uplinkBase',core:'coreBase'};
 const keys=['serverBase','uplinkBase','coreBase','fiberPolarity','mpoGender','cableJacket','fiberPatchM'];
 const catalog=[
  {vendor:'YOFC',name:'MPO/MTP pre-terminated trunk · G.657.A2 SMF',kind:'fiber',fiber:'OS2',base:12,cores:[12,24,48,72,96,144],source:'https://en.yofc.com/view/3030.html'},
  {vendor:'Corning',name:'Professional LC UPC duplex OS2 patch family · reference SKU 1 m / project length RFQ',kind:'lc',fiber:'OS2',base:2,source:'https://ecatalog.corning.com/optical-communications/AU/en/Fiber-Optic-Cable-Assemblies/Indoor-Cable-Assemblies/Two-Fiber-Indoor-Cable-Assemblies/Professional-2-0-mm-SM-2-Fiber-Patch-Cord/p/040402G5Z20001M'},
  {vendor:'US Conec',name:'MTP / MTP-16 connector component family',kind:'connector',bases:[8,12,16],source:'https://www.usconec.com/connectors/mtp-connectors'},
  {vendor:'YOFC',name:'Cat.6 Unshielded RJ45 patch cord',kind:'copper',source:'https://en.yofc.com/view/3010.html'},
  {vendor:'Corning',name:'EDGE8 MTP trunk · Base-8',kind:'fiber',base:8,fiber:'OS2',cores:[8,16,48,96,144],source:'https://ecatalog.corning.com/optical-communications/emea/en/Fiber-Optic-Cable-Assemblies/Indoor-Cable-Assemblies/Multifiber-Indoor-Cable-Assemblies/EDGE8%C2%AE-MTP%C2%AE-Trunk/p/edge8-mtp-trunk-cable'},
  {vendor:'Corning',name:'EDGE MTP trunk · Base-12',kind:'fiber',base:12,fiber:'OS2',cores:[12],source:'https://ecatalog.corning.com/optical-communications/CALA/en/Fiber-Optic-Cable-Assemblies/Indoor-Cable-Assemblies/Multifiber-Indoor-Cable-Assemblies/EDGE%E2%84%A2-MTP%C2%AE-Trunk/p/edge-trunk-cable'},
  {vendor:'SENKO',name:'MPO PLUS connector / assembly component',kind:'connector',bases:[8,12,16],source:'https://www.senko.com/product/mpo-plus-standard-connector/'},
  {vendor:'SENKO',name:'MPO PLUS dust-shutter adapter · center/offset key',kind:'adapter',bases:[8,12,16],source:'https://www.senko.com/product/mpo-plus-dust-shutter-adapter/'},
  {vendor:'YOFC',name:'MPO/MTP pre-terminated cable · OM4 family',kind:'fiber',fiber:'OM4',base:12,source:'https://en.yofc.com/view/3030.html'},
  {vendor:'YOFC',name:'MPO/MTP patch cable · OM4 family',kind:'patch',fiber:'OM4',base:12,source:'https://en.yofc.com/view/3040.html'},
  {vendor:'YOFC',name:'UDF high-density fiber panel',kind:'panel',source:'https://en.yofc.com/view/3029.html'},
  {vendor:'Sumitomo Electric Lightwave',name:'Indoor ribbon cable · bulk / field termination',kind:'bulk',fiber:'OS2',source:'https://sumitomoelectriclightwave.com/product/indoor-rohs-riser-ribbon-cable/'},
  {vendor:'Fujikura',name:'WTC / SWR fiber cable · bulk family',kind:'bulk',source:'https://www.optic-product.fujikura.com/fiber-optic-cable/'}
 ];
 function media(p,segment,x){
  if(!p.optical)return {...p,base:0,connectorGroups:0};
  const selected=x[baseKey[segment]]||'auto',mpo=/MPO/.test(p.connector),speed=/800G/.test(p.media),dr=/DR/.test(p.media),sr=/SR/.test(p.media);
  let base=mpo?(/MPO-16/.test(p.connector)?16:12):2,groups=mpo?(speed?2:1):(speed?2:1);
  if(selected!=='auto'&&baseKey[segment]){
   const want=Number(selected);if(![8,12,16].includes(want))throw Error('Base 선택은 auto/8/12/16입니다.');
   if(!mpo)throw Error(segment+': LC duplex 광모듈에는 MPO Base를 적용할 수 없습니다. 자동을 선택하세요.');
   if(sr&&(!speed&&want!==16||speed&&want===16))throw Error(segment+': 선택한 SR 광인터페이스와 Base가 맞지 않습니다. 400G SR8=Base-16, 800G 2×SR4=Base-8/12.');
   if(dr&&!speed&&want===16)throw Error(segment+': 400G DR4는 MPO-12의 8 활성 심수입니다. Base-8/12를 선택하세요.');
   base=want;if(dr&&speed&&want===16){groups=1;p={...p,media:'800G DR8P · single MPO-16',connector:'MPO-16/APC',reference:'https://www.cisco.com/c/en/us/products/collateral/interfaces-modules/transceiver-modules/osfp-800g-transceiver-modules-ds.html',evidence:'DR8P 인터페이스 제품군 · 장비별 SKU 호환성 별도 확인'};}
  }
  const installed=base*groups,requested=Number(x.trunkFiberCount||16),actual=Math.ceil(requested/base)*base;
  if(actual>144)throw Error(segment+': Base 배수로 올림한 트렁크 심수가 144F를 초과합니다.');
  return {...p,base,connectorGroups:groups,installedChannelFibers:installed,trunkFiberCount:actual,requestedTrunkFiberCount:requested,polarity:x.fiberPolarity||'review',gender:x.mpoGender||'review',baseNote:base===8?'MPO-12 ferrule · 8 populated (4+4)':base===16?'MPO-16 offset key · MPO-12와 직접 결합 불가':base===12?'MPO-12 center key · 12 populated':'LC duplex · 2 fibers per group'};
 }
 function match(q){return catalog.filter(c=>{
  const o=q.profile;if(!o)return false;
  if(c.kind==='lc')return o.optical&&o.fiberType===c.fiber&&/LC/.test(o.connector)&&o.installedChannelFibers===2&&q.unit==='channel'&&!/adapter/i.test(q.item);
  if(c.kind==='copper')return /RJ45/.test(o.connector)&&q.unit==='assembly';
  if(c.kind==='connector')return /MPO/.test(o.connector)&&!['module','assembly','reel'].includes(q.unit)&&c.bases.includes(o.base);
  if(c.kind==='adapter')return /adapter/i.test(q.item)&&/MPO/.test(o.connector)&&c.bases.includes(o.base);
  if(c.kind==='panel')return /panel/i.test(q.item);
  if(c.kind==='bulk')return q.unit==='reel'&&(!c.fiber||c.fiber===o.fiberType);
  if(!o.optical||q.unit==='module'||/adapter/i.test(q.item)||q.unit==='reel')return false;
  if(c.fiber!==o.fiberType||c.base!==o.base||c.cores&&!c.cores.includes(q.fibers))return false;
  return c.kind==='patch'?/patch|harness/i.test(q.item):q.unit==='trunk';
 });}
 function enrich(r,x){
  for(const k of keys)r.input[k]=x[k]??(k==='fiberPatchM'?2:'review');
  const patchM=Number(x.fiberPatchM??2);if(!Number.isFinite(patchM)||patchM<0||patchM>100)throw Error('패치코드 길이는 0–100 m입니다.');
  r.productRequirements=r.bom.map(b=>{const o=r.optical[b.segment],p=o?.profile;return {category:b.category,item:b.item,segment:b.segment||'',installed:b.installedQty??b.qty,purchase:b.qty,unit:b.unit,lengthM:/patch|harness/i.test(b.item)?patchM:Number(r.input[{server:'serverDistanceM',leafSpine:'leafSpineDistanceM',core:'coreDistanceM'}[b.segment]||b.segment+'DistanceM'])||null,fibers:p?.optical?(b.unit==='trunk'||b.unit==='reel'?p.trunkFiberCount||r.input.trunkFiberCount:p.installedChannelFibers):0,activeFibers:p?.fibers||0,profile:p,polarity:p?.optical?(x.fiberPolarity||'review'):'N/A',gender:p&&/MPO/.test(p.connector)?x.mpoGender||'review':'N/A',jacket:x.cableJacket||'project',candidates:[]};});
  for(const q of r.productRequirements)q.candidates=match(q);
  const opticalLinks=Object.values(r.optical||{}).filter(o=>o.profile.optical).reduce((s,o)=>s+o.links,0),copperLinks=Object.values(r.optical||{}).filter(o=>!o.profile.optical&&/RJ45/.test(o.profile.connector)).reduce((s,o)=>s+o.links,0),assemblies=r.productRequirements.filter(q=>['trunk','assembly','channel'].includes(q.unit)&&!/adapter/i.test(q.item));
  r.cablingHandover=[{item:'Cable end ID labels',qty:2*assemblies.reduce((s,q)=>s+q.installed,0),unit:'label',basis:'설치 케이블/패치 채널 1개당 양단 2개 · fanout 세부 라벨 별도'},{item:'Optical channel test records',qty:opticalLinks,unit:'record',basis:'논리 광링크당 IL/RL·극성·연속성 기록; 합격 기준은 선정 SKU/광 예산'},{item:'Copper channel test records',qty:copperLinks,unit:'record',basis:'관리망 링크별 인증 시험; Category·차폐·규격은 RFQ'},{item:'Rack/route/port map + packing manifest',qty:1,unit:'set',basis:'계산상 endpoint ID → 실제 rack/port/트레이 ID 매핑 필요'}];
  const racks=r.racks||[],sum=racks.reduce((s,z)=>s+Number(z.powerKw||0),0),compute=racks.filter(z=>z.role==='Compute'),computeKw=compute.reduce((s,z)=>s+z.powerKw,0),it=r.summary.totalItPowerKw,pue=Number(r.input.pue),head=1+Number(r.input.facilityHeadroomPct||0)/100;
  r.powerProof={itKw:it,rackSumKw:sum,residualKw:Math.abs(sum-it),totalRacks:racks.length,averageRackKw:racks.length?it/racks.length:0,computeRacks:compute.length,computeKw,averageComputeRackKw:compute.length?computeKw/compute.length:0,energizedRacks:racks.filter(z=>z.powerKw>0).length,pue,facilityKw:it*pue,overheadKw:it*(pue-1),upsKw:it*head,generatorKw:it*pue*head,transformerKva:it*pue*head/Number(r.input.powerFactor),isColo:r.input.scenario==='colo'};
 }
 window.DCProcurement={keys,catalog,media,enrich,match};
})();
