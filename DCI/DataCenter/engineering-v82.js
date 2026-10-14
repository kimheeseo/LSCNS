/* DC BOM engineering planning engine. All defaults are assumptions, not SKU qualification. */
(function(root){
'use strict';
const SYSTEMS={
 h200:{name:'H200 8-GPU planning node',gpus:8,ru:8,power:10.2,links:8,linkSpeed:400,cordCount:6},
 b200:{name:'B200 8-GPU planning node',gpus:8,ru:10,power:14.3,links:8,linkSpeed:400,cordCount:6},
 b300:{name:'B300 8-GPU planning node',gpus:8,ru:10,power:15,links:8,linkSpeed:800,cordCount:12},
 gb200:{name:'GB200 NVL72 planning rack',gpus:72,ru:48,power:120,links:72,linkSpeed:400,cordCount:0,rackScale:true,trays:18,rails:4},
 gb300:{name:'GB300 NVL72 planning rack',gpus:72,ru:48,power:140,links:72,linkSpeed:800,cordCount:0,rackScale:true,trays:18,rails:4}
};
const OPTICS=[
 {id:'dac',speed:400,reach:3,loss:null,fiber:'Copper',connector:'DAC',fibers:0},
 {id:'dac',speed:800,reach:2,loss:null,fiber:'Copper',connector:'DAC',fibers:0},
 {id:'aec',speed:400,reach:5,loss:null,fiber:'Copper',connector:'AEC',fibers:0},
 {id:'aec',speed:800,reach:5,loss:null,fiber:'Copper',connector:'AEC',fibers:0},
 {id:'aoc',speed:400,reach:30,loss:null,fiber:'Integrated',connector:'AOC',fibers:0},
 {id:'aoc',speed:800,reach:30,loss:null,fiber:'Integrated',connector:'AOC',fibers:0},
 {id:'sr',speed:400,standard:'400G-SR4',reach:100,loss:1.9,fiber:'OM4',connector:'MPO-12',polish:'UPC',fibers:8,base:12},
 {id:'sr',speed:800,standard:'800G-SR8',reach:100,loss:1.9,fiber:'OM4',connector:'MPO-16',polish:'UPC',fibers:16,base:16},
 {id:'dr',speed:400,standard:'400G-DR4',reach:500,loss:3,fiber:'OS2',connector:'MPO-12',polish:'APC',fibers:8,base:12},
 {id:'dr',speed:800,standard:'800G-DR8',reach:500,loss:3,fiber:'OS2',connector:'2×MPO-12',polish:'APC',fibers:16,base:12,groups:2},
 {id:'fr',speed:400,standard:'400G-FR4',reach:2000,loss:4,fiber:'OS2',connector:'Duplex LC',polish:'UPC',fibers:2,base:2},
 {id:'fr',speed:800,standard:'800G-2×FR4',reach:2000,loss:4,fiber:'OS2',connector:'2×Duplex LC',polish:'UPC',fibers:4,base:2,groups:2},
 {id:'lr',speed:400,standard:'400G-LR4-6',reach:6000,loss:6.3,fiber:'OS2',connector:'Duplex LC',polish:'UPC',fibers:2,base:2},
 {id:'lr',speed:800,standard:'800G-2×LR4 planning',reach:6000,loss:6.3,fiber:'OS2',connector:'2×Duplex LC',polish:'UPC',fibers:4,base:2,groups:2}
];
const DEFAULTS={targetGPU:1024,systemId:'b200',inputMode:'gpu',itInputMw:100,nvlRackCount:4,topology:'rail',fabricProtocol:'InfiniBand',fabricTier:2,oversub:1,
 rackRU:48,reservedRU:4,rackPowerKw:100,rackCoolingKw:100,rackHeadroomPct:20,coolingMode:'auto',sparePct:10,
 serverDistanceM:30,leafSpineDistanceM:100,coreDistanceM:100,serverCabling:'structured',leafSpineCabling:'structured',coreCabling:'structured',trunkFiberCount:16,
 switchRadix:64,switchKw:2,switchRU:2,switchLogicalPerCage:1,serverMedia:'auto',uplinkMedia:'auto',coreMedia:'auto',fiberPreference:'auto',
 slackPct:15,verticalRiseM:3,terminationM:0,matedPairs:4,p2pPairs:2,mpoLossDb:.35,lcLossDb:.25,lowLoss:false,
 fiberAttenDbKm:.35,mmfAttenDbKm:3,marginDb:0,spliceCount:0,spliceLossDb:.1,fiberPolarity:'B',connectorPolish:'auto',
 pue:1.2,facilityHeadroomPct:20,powerFactor:.9,upsEfficiency:.96,upsBlockKva:1000,upsUtilPct:80,generatorBlockKw:2500,generatorUtilPct:80,transformerBlockKva:3000,transformerUtilPct:80,facilityRedundancy:'N+1',facilityEnabled:false,
 liquidPct:75,cduBlockKw:600,cduRedundancy:'N+1',waterDensity:1000,waterCp:4186,deltaTK:10,airLimitKw:30,rearDoorLimitKw:50,
 panelChannels:24,cassetteChannels:12,pduRatedKw:46,pduUtilPct:80,pduOutlets:24,siteVoltage:400,
 storageEnabled:false,storageCapacityTb:1000,storageEncoding:'ec',storageReadGbpsPerGpu:0,storageRawTbPerNode:1000,storageReadGbpsPerNode:50,storageArrayKw:10,storageArrayRU:4,storageProtocol:'NVMe-oF/RoCE',storageEndpointPorts:2,storageSwitchPorts:64,storageSwitchKw:1,storageSwitchRU:2,
 frontendEnabled:true,frontendPortsPerServer:2,frontendRate:'100G',auxRadix:48,auxSwitchKw:.5,auxOversub:4,oobEnabled:true,consolePorts:32,
 trayWidthMm:300,trayDepthMm:50,cableOdMm:5,trayRouteM:0,trunkBundle:1,modulePack:1,assemblyPack:1,channelPack:1,twinPort400:false,
 enterpriseRackEnabled:false,firewallQty:2,routerQty:2,loadBalancerQty:2,kvmQty:1,nasQty:1,opsUnitKw:1,opsUnitRU:1
};
const ceil=Math.ceil, round=x=>Math.round(x*1000)/1000, sum=a=>a.reduce((s,x)=>s+x,0), red=(n,r)=>n===0?0:r==='2N'?2*n:r==='N+1'?n+1:n;
function calculate(raw={}){
 const i={...DEFAULTS,...raw},warnings=[],errors=[];
 // Invalid numeric values get a visible warning and a safe planning fallback. No NaN/Infinity BOMs.
 for(const [k,v] of Object.entries(DEFAULTS))if(typeof v==='number'){
  const x=Number(i[k]);if(!Number.isFinite(x)||x<0){warnings.push(k+': invalid input; default '+v+' used');i[k]=v;}else i[k]=x;
 }
 const positive=['rackRU','rackPowerKw','rackCoolingKw','switchRadix','powerFactor','upsEfficiency','upsBlockKva','generatorBlockKw','transformerBlockKva','cduBlockKw','waterDensity','waterCp','deltaTK','pduRatedKw','pduOutlets','panelChannels','cassetteChannels','storageRawTbPerNode','storageReadGbpsPerNode','auxRadix','consolePorts','trayWidthMm','trayDepthMm','siteVoltage'];
 for(const k of positive)if(i[k]<=0){warnings.push(k+': must be positive; default used');i[k]=DEFAULTS[k];}
 for(const k of ['rackHeadroomPct','facilityHeadroomPct','liquidPct','upsUtilPct','generatorUtilPct','transformerUtilPct','pduUtilPct'])if(i[k]>=(k==='liquidPct'?101:100)||(['upsUtilPct','generatorUtilPct','transformerUtilPct','pduUtilPct'].includes(k)&&i[k]===0)){warnings.push(k+': outside planning range; default used');i[k]=DEFAULTS[k];}
 for(const k of ['powerFactor','upsEfficiency'])if(i[k]>1){warnings.push(k+': cannot exceed 1; default used');i[k]=DEFAULTS[k];}
 if(i.pue<1){warnings.push('PUE < 1; 1 used');i.pue=1;}
 if(i.oversub<1){warnings.push('Oversubscription must be ≥1; 1 used');i.oversub=1;}
 if(i.topology==='clos3'){i.topology='rail';i.fabricTier=3;}
 if(!['rail','single','dual'].includes(i.topology)){warnings.push('Unknown topology; Rail used');i.topology='rail';}
 const selected=SYSTEMS[i.systemId];
 let p={...(selected||SYSTEMS.b200),...(raw.systemAssumptions?.[i.systemId]||{})};
 if(i.systemId==='custom'||(!selected&&i.customKw))p={name:i.customName||'Custom',gpus:i.customGpus||8,ru:i.customRU||10,power:i.customKw||14.3,links:i.customLinks||8,linkSpeed:Number(i.customSpeed)||400,cordCount:i.customCords||6};
 if(i.nicOverride){p.links=Number(i.nicLinks)||p.links;p.linkSpeed=Number(i.nicSpeed)||p.linkSpeed;}
 for(const k of ['gpus','ru','power','links','linkSpeed'])if(!Number.isFinite(Number(p[k]))||p[k]<=0){warnings.push('System '+k+': default used');p[k]=(selected||SYSTEMS.b200)[k];}else p[k]=Number(p[k]);
 p.links=ceil(p.links);p.gpus=ceil(p.gpus);p.ru=ceil(p.ru);
 if(![400,800].includes(p.linkSpeed)){warnings.push('Only 400/800G profiles supported; 400G used');p.linkSpeed=400;}
 if(p.rackScale&& !['dlc','immersion'].includes(i.coolingMode)){warnings.push('NVL72 requires liquid cooling; DLC used');i.coolingMode='dlc';}
 const mode=i.coolingMode==='auto'?(p.rackScale?'dlc':'air'):i.coolingMode;
 const head=1-i.rackHeadroomPct/100;
 const limits=p.rackScale?{RU:1,power:1,cooling:1}:{RU:Math.floor(Math.max(0,i.rackRU-i.reservedRU)/p.ru),power:Math.floor(i.rackPowerKw*head/p.power),cooling:Math.floor(i.rackCoolingKw*head/p.power)};
 if(!p.rackScale&&i.topology!=='rail'){const maxNicPerUnit=i.topology==='dual'?ceil(p.links/2):p.links;limits.NIC=Math.floor(Math.floor(i.switchRadix/2)/maxNicPerUnit);}
 const fit=Math.min(...Object.values(limits)),perRack=Math.max(1,fit),binding=Object.keys(limits).filter(k=>limits[k]===fit);
 if(fit<1)errors.push('No server fits rack limits; one per rack shown for review');
 if(i.rackPowerKw>i.rackCoolingKw)errors.push('Rack power rating exceeds rack cooling rating');
 let units,racks,target=i.targetGPU,powerBudget=null;
 if(p.rackScale){racks=i.inputMode==='power'?ceil(i.itInputMw*1000*head/p.power):i.inputMode==='racks'?Math.max(1,ceil(i.nvlRackCount)):Math.max(1,ceil(target/p.gpus));units=racks;target=i.inputMode==='gpu'?target:racks*p.gpus;}
 else if(i.inputMode==='power'){powerBudget=i.itInputMw*1000*head;racks=Math.max(1,ceil(powerBudget/i.rackPowerKw));units=Math.min(racks*perRack,Math.max(1,Math.floor(powerBudget/p.power)));target=units*p.gpus;}
 else {units=Math.max(1,ceil(target/p.gpus));racks=ceil(units/perRack);}
 if(units>200000){errors.push('Planning limit 200000 systems; reduce input');units=200000;racks=p.rackScale?units:ceil(units/perRack);target=units*p.gpus;}
 const computeRacks=Array.from({length:racks},(_,idx)=>{const n=p.rackScale?1:Math.max(0,Math.min(perRack,units-idx*perRack));return {id:'compute-'+(idx+1),role:'Compute',units:n,ru:p.rackScale?p.ru:n*p.ru,powerKw:round(n*p.power),cords:n*p.cordCount};});
 if(p.rackScale){if(p.power>i.rackPowerKw)errors.push('NVL72 rack kW exceeds configured rack power');if(p.power>i.rackCoolingKw)errors.push('NVL72 rack kW exceeds configured cooling');}
 const maxRackKw=computeRacks.reduce((v,r)=>Math.max(v,r.powerKw),0);
 if(mode==='air'&&maxRackKw>i.airLimitKw)warnings.push('Air cooling planning threshold exceeded: '+maxRackKw+' > '+i.airLimitKw+' kW');
 if(mode==='rear-door'&&maxRackKw>i.rearDoorLimitKw)warnings.push('Rear-door HX planning threshold exceeded: '+maxRackKw+' > '+i.rearDoorLimitKw+' kW');
 // Each domain has its own fabric. ToR modes preserve rack affinity; Dual splits existing NICs across A/B fabrics.
 const R=Math.max(4,Math.floor(i.switchRadix/2)*2),down=R/2,up=Math.max(1,ceil(down/i.oversub));
 if(i.twinPort400&&p.linkSpeed===400){i.switchLogicalPerCage=2;warnings.push('Twin-port option assumes two qualified 400G logical ports/switch cage');}
 const domains=i.topology==='rail'?(p.rackScale?p.rails||4:p.links):i.topology==='dual'?2:1;
 if(R!==i.switchRadix)warnings.push('Switch radix rounded to even '+R);i.switchRadix=R;
 const endpoints=units*p.links,portNodes=[],routes=[],domainRows=[],counts={leaf:0,spine:0,core:0,server:0,uplink:0,coreLink:0};
 const addNode=(id,role,capacity)=>{const node={id,role,capacity,logicalPerCage:i.switchLogicalPerCage,used:{server:0,leafSpine:0,core:0}};portNodes.push(node);return node;};
 function route(a,b,n,seg){if(!n)return;routes.push({a:a.id,b:b.id,links:n,segment:seg});const k=seg==='server'?'server':seg==='leafSpine'?'leafSpine':'core';a.used[k]+=n;b.used[k]+=n;}
 for(let domain=0;domain<domains;domain++){
  const assignedPerUnit=Math.floor(p.links/domains)+(domain<p.links%domains?1:0);
  if(p.links%domains)warnings.push('NIC ports are distributed unevenly across independent domains');
  const groups=i.topology==='rail'?[{id:'all',n:ceil(units*assignedPerUnit)}]:computeRacks.map(r=>({id:r.id,n:ceil(r.units*assignedPerUnit)}));
  const leaves=[];
  if(i.topology==='rail'){
   let leaf=null,remaining=0;
   for(const rack of computeRacks){let left=ceil(rack.units*assignedPerUnit);while(left>0){if(!remaining){leaf=addNode('D'+domain+'-L'+leaves.length,'Leaf',R);leaves.push(leaf);remaining=down;}const n=Math.min(left,remaining),host=addNode('D'+domain+'-'+rack.id+'-NIC'+leaves.length,'Endpoint',n);route(host,leaf,n,'server');left-=n;remaining-=n;}}
  }else for(const group of groups){if(group.n>down)warnings.push(i.topology+' ToR '+group.id+': requires multiple Leaves at current radix; choose larger switch or fewer endpoints');let remaining=group.n;while(remaining>0){const n=Math.min(down,remaining),leaf=addNode('D'+domain+'-L'+leaves.length,'Leaf',R),host=addNode('D'+domain+'-'+group.id+'-NIC'+leaves.length,'Endpoint',n);route(host,leaf,n,'server');leaves.push(leaf);remaining-=n;}}
  const hostCount=sum(groups.map(g=>g.n)),twoLimit=R*down,threeLimit=R*down*down;
  let tier=Number(i.fabricTier)===3?3:2;
  if(leaves.length>R||hostCount>twoLimit){tier=3;warnings.push('Domain '+domain+': 2-tier capacity exceeded; promoted to 3-tier');}
  if(hostCount>threeLimit||leaves.length>R*down)errors.push('Domain '+domain+': exceeds 3-tier capacity; additional independent fabrics required');
  let spineCount=0,coreCount=0,coreLinks=0;
  if(tier===2){
   let compactSpines=1;while(compactSpines<up&&(up%compactSpines!==0||leaves.length*(up/compactSpines)>R))compactSpines++;
   const spines=Array.from({length:compactSpines},(_,j)=>addNode('D'+domain+'-S'+j,'Spine',R));
   for(const leaf of leaves)for(const spine of spines)route(leaf,spine,up/compactSpines,'leafSpine');
   spineCount=spines.length;
  }else{
   // Regular folded Clos: <= R/2 leaves/pod; up spine planes/pod; each spine reserves leafCount upstream ports.
   const podLeafCap=Math.max(1,Math.min(down,Math.floor(i.podLeafLimit||down)));
   const pods=[];for(let at=0;at<leaves.length;at+=podLeafCap)pods.push(leaves.slice(at,at+podLeafCap));
   if(pods.length>R)errors.push('Domain '+domain+': configured pod Leaf limit creates more than R pods; split fabric');
   const maxLeaves=Math.max(1,...pods.map(x=>x.length));
   let podSpines=1;while(podSpines<up&&(up%podSpines!==0||maxLeaves*(up/podSpines)>down))podSpines++;
   const parallelLeafLinks=up/podSpines,spineUp=maxLeaves*parallelLeafLinks;
   // Equal parallel links within a core plane. Choose divisor so every core respects radix.
   let coresPerPlane=1;while(coresPerPlane<spineUp&&(spineUp%coresPerPlane!==0||pods.length*(spineUp/coresPerPlane)>R))coresPerPlane++;
   const planes=Array.from({length:podSpines},(_,plane)=>Array.from({length:coresPerPlane},(_,j)=>addNode('D'+domain+'-C'+plane+'-'+j,'Core',R)));
   pods.forEach((pod,j)=>{for(let plane=0;plane<podSpines;plane++){
    const spine=addNode('D'+domain+'-P'+j+'-S'+plane,'Spine',R);spineCount++;
    for(const leaf of pod)route(leaf,spine,parallelLeafLinks,'leafSpine');
    // Pad the small final pod's upstream reservations; explicit upstream excess is recorded.
    for(const core of planes[plane]){const n=spineUp/coresPerPlane;route(spine,core,n,'core');coreLinks+=n;}
   }});coreCount=podSpines*coresPerPlane;
  }
  counts.leaf+=leaves.length;counts.spine+=spineCount;counts.core+=coreCount;counts.server+=hostCount;counts.uplink+=leaves.length*up;counts.coreLink+=coreLinks;
  domainRows.push({domain,hosts:hostCount,leaf:leaves.length,spine:spineCount,core:coreCount,tier,twoLimit,threeLimit,down,up,allocatedOversub:down/up,actualOversub:hostCount/Math.max(1,leaves.length*up)});
 }
 for(const n of portNodes)if(sum(Object.values(n.used))>n.capacity)errors.push('Port capacity exceeded: '+n.id);
 const bom=[],procurement=[],optical={},requirements=[],guide=[],purchase=(n,pack=1)=>{pack=Math.max(1,ceil(pack));return ceil((n+ceil(n*i.sparePct/100))/pack)*pack;},row=(category,item,n,unit,basis,extra={})=>{const qty=ceil(n);bom.push({category,item,qty,installedQty:extra.installedQty??qty,spareQty:extra.installedQty===undefined?0:qty-extra.installedQty,unit,basis,...extra});};
 row('Compute',p.name,units,p.rackScale?'rack':'system',(p.rackScale?'ceil(GPU/72)':'ceil(GPU/'+p.gpus+')')+'; configured power '+p.power+' kW');
 row('Network','Compute NIC logical endpoints',endpoints,'port',units+' × '+p.links+' NIC ports');
 for(const [key,n]of[['Leaf',counts.leaf],['Spine',counts.spine],['Super-spine',counts.core]])if(n)row('Network',key+' '+p.linkSpeed+'G planning switch',n,'switch','Sum independent domains; R='+R+', down='+down+', up='+up+'; no spare');
 if(p.rackScale){row('Compute','NVLink copper backplane assembly',racks,'rack kit','1 OEM rack kit/rack; RFQ, not scale-out optics');row('Compute','NVLink switch trays (included OEM rack)',9*racks,'tray','9/rack reference, included in rack kit; do not double-order');}
 function link(segment,links,speed,distance,cabling,mediaChoice){
  if(!links)return;
  const routed=(distance+i.verticalRiseM+2*i.terminationM)*(1+i.slackPct/100),standard=[1,2,3,5,10,20,30,50,100].find(m=>m>=routed),length=standard||ceil(routed/10)*10;
  if(!standard)warnings.push(segment+': route >100m; custom length '+length+'m RFQ');
  const table=OPTICS.filter(x=>x.speed===speed).map(x=>({...x,reach:raw.opticsAssumptions?.[speed+'-'+x.id]?.reach??x.reach,loss:raw.opticsAssumptions?.[speed+'-'+x.id]?.loss??x.loss}));
  const lossOf=x=>{const optical=x.loss!==null,budget=optical&&i.opticalBudgetMode==='project'?Math.min(x.loss,Number(i.projectBudgetDb)||x.loss):x.loss,pairs=cabling==='structured'?Math.max(4,ceil(i.matedPairs)):Math.max(2,ceil(i.p2pPairs)),atten=x.fiber==='OM4'?i.mmfAttenDbKm:i.fiberAttenDbKm,connector=/LC/.test(x.connector)?i.lcLossDb:i.lowLoss?.25:i.mpoLossDb;return {applicable:optical,matedPairs:pairs,attenDbKm:atten,connectorDb:connector,estimatedDb:optical?round(atten*length/1000+pairs*connector+i.spliceCount*i.spliceLossDb):0,budgetDb:budget,marginDb:i.marginDb,pass:length<=x.reach&&(!optical||atten*length/1000+pairs*connector+i.spliceCount*i.spliceLossDb+i.marginDb<=budget)};};
  const available=table.filter(x=>cabling==='p2p'||x.loss!==null).filter(x=>i.fiberPreference==='smf'?x.loss===null||x.fiber==='OS2':i.fiberPreference==='mmf'?x.loss===null||x.fiber==='OM4':true);
  let chosen=mediaChoice&&mediaChoice!=='auto'?table.find(x=>x.id===mediaChoice):available.filter(x=>x.id!=='aec').find(x=>lossOf(x).pass);
  chosen=chosen||available.at(-1)||table.at(-1);const loss=lossOf(chosen),isOptical=chosen.loss!==null;
  const alternatives=table.filter(x=>lossOf(x).pass&&(cabling==='p2p'||x.loss!==null)).map(x=>x.standard||x.id.toUpperCase());
  if(!loss.pass)errors.push(segment+': '+(chosen.standard||chosen.id)+' FAIL; alternatives '+(alternatives.join(', ')||'reduce connector pairs / direct cabling / vendor-qualified long reach'));
  if(cabling==='structured'&&!isOptical)errors.push(segment+': DAC/AOC/AEC cannot traverse passive patch panels; select point-to-point');
  if(isOptical&&i.fiberPolarity!=='B'&&/MPO/.test(chosen.connector))errors.push(segment+': Type-B end-to-end polarity required by planning profile; verify assembly mapping');
  if(isOptical&&i.connectorPolish!=='auto'&&i.connectorPolish!==chosen.polish)errors.push(segment+': '+i.connectorPolish+' polish conflicts with '+chosen.polish+' profile; never mate APC/UPC');
  const groups=chosen.groups||1,base=chosen.base||0,channelFibers=base*groups,channelsPerTrunk=isOptical?Math.floor(i.trunkFiberCount/channelFibers):0;
  const actualTrunk=channelsPerTrunk>0?i.trunkFiberCount:channelFibers;
  if(isOptical&&cabling==='structured'&&!channelsPerTrunk)warnings.push(segment+': trunk too small; '+actualTrunk+'F required per channel; RFQ');
  // Share only qualified 2x400G switch modules, grouped per physical switch. Server-side NIC remains one module/port.
  let installedModules=isOptical?2*links:0;
  if(isOptical&&speed===400&&i.twinPort400){
   const field=segment==='server'?'server':segment==='leafSpine'?'leafSpine':'core';
   const eligible=portNodes.filter(n=>segment==='server'?n.role==='Leaf':segment==='leafSpine'?['Leaf','Spine'].includes(n.role):['Spine','Core'].includes(n.role));
   installedModules=segment==='server'?links:0;installedModules+=sum(eligible.map(n=>ceil(n.used[field]/2)));
   warnings.push(segment+': 2×400G modules assumed; matched breakout/FEC/cage/connector SKU required');
  }
  const trunkCapacity=Math.max(1,Math.floor(actualTrunk/channelFibers));
  const segmentRoutes=routes.filter(q=>q.segment===segment);
  const trunkInstalled=isOptical&&cabling==='structured'?sum(segmentRoutes.map(q=>ceil(q.links/trunkCapacity))):0;
  const selectedBase=i[{server:'serverBase',leafSpine:'uplinkBase',core:'coreBase'}[segment]];
  if(selectedBase&&selectedBase!=='auto'&&Number(selectedBase)!==base)warnings.push(segment+': Base-'+selectedBase+' conflicts with planned Base-'+base+'; RFQ interface redesign');
  const profile={speed,standard:(chosen.standard||'').replace(/^\d+G-/,''),protocol:i.fabricProtocol,package:'',media:chosen.standard||speed+'G '+chosen.id.toUpperCase(),optical:isOptical,optics:isOptical?2:0,connector:chosen.connector,polish:chosen.polish,fibers:chosen.fibers,installedChannelFibers:channelFibers,base,connectorGroups:groups,fiberType:chosen.fiber,trunkFiberCount:actualTrunk,requestedTrunkFiberCount:i.trunkFiberCount,compatibility:'REVIEW',rule:'Assumption; confirm vendor datasheet and host/FEC'};
  optical[segment]={links,lengthM:length,routeLengthM:routed,profile,installedModules,purchasedModules:purchase(installedModules,i.modulePack),loss:{...loss,alternatives},trunk:{installedCableCount:trunkInstalled,cableCount:purchase(trunkInstalled),installedChannelFibers:channelFibers,requiredFibers:links*chosen.fibers,provisionedFibers:trunkInstalled*actualTrunk}};
  guide.push({speed:speed+'G',media:profile.media,distanceM:length,connector:profile.connector,rule:'Loss '+loss.estimatedDb+'/'+loss.budgetDb+' dB; '+(loss.pass?'PASS planning':'FAIL')+'; assumption only',segment});
  if(isOptical){row('Connectivity',segment+' '+profile.media+' transceiver',purchase(installedModules,i.modulePack),'module',installedModules+' installed + ceil(spare '+i.sparePct+'%); '+(i.twinPort400&&speed===400?'qualified twin switch ports':'2 modules/link'),{segment,installedQty:installedModules});
   if(cabling==='structured'){
    row('Connectivity',segment+' '+actualTrunk+'F trunk '+length+' m',purchase(trunkInstalled,i.trunkBundle),'cable','ceil('+links+'/'+Math.max(1,Math.floor(actualTrunk/channelFibers))+') channels/trunk; reserve Base-'+base+' × '+groups,{segment,installedQty:trunkInstalled});
    row('Connectivity',segment+' optical patch leads',purchase(2*links,i.channelPack),'cable','2 endpoint patch leads/link; additional intermediate patching RFQ',{segment,installedQty:2*links});
    row('Passive',segment+' patch panels',2*ceil(links/i.panelChannels),'panel','2 sides × ceil('+links+'/'+i.panelChannels+')');
    if(base>2)row('Passive',segment+' MPO through adapter cassettes',2*ceil(links/i.cassetteChannels),'cassette','2 sides × ceil(channels/'+i.cassetteChannels+'); parallel MPO, not MPO→LC conversion');
   }else row('Connectivity',segment+' P2P fiber '+length+' m',purchase(links,i.channelPack),'cable','1 assembly/link; ceil spare',{segment,installedQty:links});
  }else row('Connectivity',segment+' '+profile.media+' '+length+' m',purchase(links,i.assemblyPack),'assembly','1 integrated cable/link; no discrete optics; ceil spare',{segment,installedQty:links});
  const cableQty=trunkInstalled||links;
  const buyCables=purchase(cableQty,trunkInstalled?i.trunkBundle:isOptical?i.channelPack:i.assemblyPack);
  const reelCapacity=i.cablePurchaseMode==='reel'&&trunkInstalled?Math.floor(i.reelM/length):0;
  if(i.cablePurchaseMode==='reel'&&trunkInstalled&&!reelCapacity)errors.push(segment+': reel shorter than route; factory assemblies shown for RFQ');
  const reels=reelCapacity?ceil(buyCables/reelCapacity):null,purchaseLengthM=reels?reels*i.reelM:buyCables*length;
  procurement.push({segment,lengthM:length,installedCables:cableQty,requiredCables:buyCables,reels,purchaseLengthM,wasteM:purchaseLengthM-buyCables*length});
  requirements.push({category:'Connectivity',item:profile.media,segment,installed:installedModules||links,purchase:purchase(installedModules||links),unit:isOptical?'module':'assembly',lengthM:length,fibers:channelFibers,activeFibers:chosen.fibers,profile,polarity:i.fiberPolarity,gender:'RFQ',jacket:i.cableJacket||'RFQ',candidates:[]});
 }
 link('server',counts.server,p.linkSpeed,i.serverDistanceM,i.serverCabling,i.serverMedia);
 link('leafSpine',counts.uplink,p.linkSpeed,i.leafSpineDistanceM,i.leafSpineCabling,i.uplinkMedia);
 link('core',counts.coreLink,p.linkSpeed,i.coreDistanceM,i.coreCabling,i.coreMedia);
 if(!optical.core)optical.core={links:0,profile:{optical:false},loss:{applicable:false},trunk:{cableCount:0,installedCableCount:0,provisionedFibers:0}};
 const allRacks=[...computeRacks],auxFabrics=[];
 const addEquipmentRacks=(role,count,ru,kw,cords=2)=>{const capacity=Math.max(1,Math.min(Math.floor(Math.max(1,i.rackRU-i.reservedRU)/ru),Math.floor(Math.min(i.rackPowerKw,i.rackCoolingKw)*head/kw)));for(let at=0;at<count;at+=capacity){const n=Math.min(capacity,count-at);allRacks.push({id:role+'-'+allRacks.length,role,units:n,ru:n*ru,powerKw:round(n*kw),cords:n*cords});}};
 const switchCount=counts.leaf+counts.spine+counts.core;addEquipmentRacks('Network',switchCount,i.switchRU,i.switchKw);
 const rawRatio=i.storageEncoding==='replica3'?1/3:.8;
 const storageNodes=i.storageEnabled?Math.max(ceil(i.storageCapacityTb/(i.storageRawTbPerNode*rawRatio)),ceil(target*i.storageReadGbpsPerGpu/i.storageReadGbpsPerNode)):0;
 const storageSwitches=storageNodes?2*ceil(storageNodes*i.storageEndpointPorts/i.storageSwitchPorts):0;
 if(storageNodes){row('Storage','Storage node',storageNodes,'node','max(ceil(usable TB/'+round(i.storageRawTbPerNode*rawRatio)+'), ceil('+target+'×'+i.storageReadGbpsPerGpu+'/'+i.storageReadGbpsPerNode+'))');row('Storage','Storage fabric switch',storageSwitches,'switch','2 fabrics × ceil(node ports/radix); '+i.storageProtocol+'; FC = legacy SAN');addEquipmentRacks('Storage',storageNodes,i.storageArrayRU,i.storageArrayKw);addEquipmentRacks('Network',storageSwitches,i.storageSwitchRU,i.storageSwitchKw);row('Connectivity','Storage endpoint cable',purchase(storageNodes*i.storageEndpointPorts*2),'cable','dual fabrics × endpoint ports; media/SKU RFQ',{installedQty:storageNodes*i.storageEndpointPorts*2});}
 function aux(key,endpoints,rate,copper){
  const d=Math.max(1,Math.floor(i.auxRadix/2)),u=Math.max(1,ceil(d/i.auxOversub)),leaf=ceil(endpoints/d),spine=leaf>1?ceil(leaf*u/i.auxRadix):0,sw=leaf+spine,links=endpoints+(spine?leaf*u:0);
  row('Network',key+' '+rate+' switch',sw,'switch','ceil(endpoints/'+d+') leaf + aggregate uplink capacity; RFQ full port map');
  row('Connectivity',key+' endpoint/uplink cable',purchase(links),'cable','endpoints + leaf uplinks; media length and optics RFQ',{installedQty:links});
  if(!copper)row('Connectivity',key+' optical module allowance',purchase(2*links),'module','2/link; speed, optics budget and host compatibility RFQ',{installedQty:2*links});
  addEquipmentRacks('Network',sw,1,i.auxSwitchKw);auxFabrics.push({key,title:key,endpoints,leafCount:leaf,spineCount:spine,switches:sw,links,rate,powerKw:sw*i.auxSwitchKw,scope:'Planning port allowance; detailed wiring / SKU RFQ'});return sw;
 }
 let frontSwitches=0;if(i.frontendEnabled)frontSwitches=aux('Frontend',units*(p.rackScale?p.trays:1)*i.frontendPortsPerServer,i.frontendRate,false);
 const bmc=p.rackScale?racks*27:units;
 const ops=i.enterpriseRackEnabled?sum(['firewallQty','routerQty','loadBalancerQty','kvmQty','nasQty'].map(k=>ceil(i[k]))):0;
 if(ops){for(const k of ['firewallQty','routerQty','loadBalancerQty','kvmQty','nasQty'])row('Ops',k.replace('Qty',''),ceil(i[k]),'device','User option');addEquipmentRacks('Ops',ops,i.opsUnitRU,i.opsUnitKw);}
 if(i.oobEnabled){const managed=bmc+switchCount+storageSwitches+storageNodes+frontSwitches+ops;aux('OOB',managed,'1G',true);row('Ops','BMC managed endpoints',managed,'port','Compute '+bmc+' + switches/storage/ops; OOB switch management upstream RFQ');row('Ops','Console server',ceil((switchCount+frontSwitches+storageSwitches)/i.consolePorts),'device','ceil(managed fabric switches/'+i.consolePorts+')');row('Connectivity','Console serial cable',purchase(switchCount+frontSwitches+storageSwitches),'cable','1/managed fabric switch; connector RFQ',{installedQty:switchCount+frontSwitches+storageSwitches});}
 if(i.regionPduEnabled){const phase=Number(i.sitePhases)===1?1:Math.sqrt(3);if(Number(i.siteCurrentA)>0)i.pduRatedKw=phase*i.siteVoltage*Number(i.siteCurrentA)*i.powerFactor/1000;else warnings.push('Regional PDU current invalid; configured PDU kW retained');}
 for(const r of allRacks){const q=Math.max(1,ceil(r.powerKw/(i.pduRatedKw*i.pduUtilPct/100)),ceil(ceil(r.cords/2)/i.pduOutlets));r.pduQty=2*q;r.phaseKw=round(r.powerKw/3);if(r.powerKw>i.rackCoolingKw)errors.push(r.id+': cooling capacity exceeded');}
 const mainSwitchNodes=portNodes.filter(n=>n.role!=='Endpoint');let switchIndex=0;for(const rack of allRacks.filter(r=>r.role==='Network'))for(let idx=0;idx<rack.units&&switchIndex<mainSwitchNodes.length;idx++)mainSwitchNodes[switchIndex++].rackId=rack.id;for(const node of portNodes.filter(n=>n.role==='Endpoint')){const match=node.id.match(/compute-(\d+)/);if(match)node.rackId='compute-'+match[1];}
 const it=round(sum(allRacks.map(r=>r.powerKw))),facilityKw=round(it*i.pue),facilityMargin=1+i.facilityHeadroomPct/100;
 const upsKva=it/i.powerFactor/i.upsEfficiency*facilityMargin,generatorKw=facilityKw*facilityMargin,transformerKva=facilityKw/i.powerFactor*facilityMargin;
 const facility={pue:i.pue,itPowerKw:it,facilityPowerKw:facilityKw,mechanicalAndLossKw:round(facilityKw-it),upsDesignKva:round(upsKva),upsDesignKw:round(upsKva*i.powerFactor),generatorDesignKw:round(generatorKw),transformerDesignKva:round(transformerKva),upsQty:red(ceil(upsKva/(i.upsBlockKva*i.upsUtilPct/100)),i.facilityRedundancy),generatorQty:red(ceil(generatorKw/(i.generatorBlockKw*i.generatorUtilPct/100)),i.facilityRedundancy),transformerQty:red(ceil(transformerKva/(i.transformerBlockKva*i.transformerUtilPct/100)),i.facilityRedundancy)};
 if(i.facilityEnabled){row('Facility','UPS '+i.upsBlockKva+' kVA module',facility.upsQty,'module','red(ceil(IT '+it+'/PF '+i.powerFactor+'/η '+i.upsEfficiency+'×'+facilityMargin+'/usable '+i.upsBlockKva*i.upsUtilPct/100+'))');row('Facility','Generator '+i.generatorBlockKw+' kW module',facility.generatorQty,'module','red(ceil(IT×PUE×headroom/usable kW)); mechanical load included');row('Facility','Transformer '+i.transformerBlockKva+' kVA module',facility.transformerQty,'module','red(ceil(IT×PUE/PF×headroom/(rated×derating)))');}
 const computeKw=units*p.power,liquidFraction=mode==='immersion'?1:mode==='dlc'?i.liquidPct/100:mode==='rear-door'?i.liquidPct/100:0;
 const liquidKw=round(computeKw*liquidFraction),airKw=round(it-liquidKw),flow=round(liquidKw*1000/(i.waterDensity*i.waterCp*i.deltaTK)*60000),cduN=ceil(liquidKw/i.cduBlockKw),cduQty=red(cduN,i.cduRedundancy);
 if(liquidKw){row('Cooling','CDU '+i.cduBlockKw+' kW module',cduQty,'module','red(ceil('+liquidKw+'/'+i.cduBlockKw+')); '+i.cduRedundancy);row('Cooling','Rack liquid manifold pair',2*racks,'manifold','Supply + return per compute rack; flow/pressure OEM RFQ');}
 row('Rack','IT rack',allRacks.length,'rack','Compute '+racks+' + equipment '+(allRacks.length-racks));row('Power','A/B rack PDU',sum(allRacks.map(r=>r.pduQty)),'PDU','Each side supports full rack kW and cord count; N-side redundancy');
 const cords=sum(allRacks.map(r=>r.cords));if(cords)row('Power','C19/C20 planning AC power cord',purchase(cords),'cable','sum(OEM cord count); connector/current verified by RFQ',{installedQty:cords});
 row('Power','A/B busway tap or RPP branch',2*allRacks.length,'branch','2/rack; A/B breaker, fault level and ampacity RFQ');
 row('Power','3-phase balance plan',3*allRacks.length,'phase schedule','Rack kW/3 per phase; discrete outlet mapping RFQ');
 row('Passive','Rack ground bonding kit',allRacks.length,'kit','1/rack; conductor sizing/site earth RFQ');
 row('Passive','Cable endpoint labels',2*sum(bom.filter(b=>b.category==='Connectivity'&&['cable','assembly'].includes(b.unit)).map(b=>b.installedQty)),'label','2 per installed cable; no spare stock labels');
 const maxBundle=sum(procurement.map(x=>x.installedCables)),fill=round(maxBundle*Math.PI*(i.cableOdMm/2)**2/(i.trayWidthMm*i.trayDepthMm)*100),trayLength=i.trayRouteM;
 if(fill>40)warnings.push('Tray shared-route upper-bound fill '+fill+'% >40%; split pathways or enlarge tray');
 row('Passive','Cable tray / optical raceway',trayLength||0,'m',trayLength?'Measured pathway '+trayLength+' m; upper-bound shared fill '+fill+'%':'RFQ: measured pathway required; cable lengths are not tray length',{coverage:trayLength?'calculated':'RFQ'});
 row('Facility','Aisle containment',racks,'rack allowance','Option: rack-face allowance only; dimensions/doors RFQ',{coverage:'option'});
 row('Facility','Fire detection / suppression',0,'RFQ','RFQ: floor area, hazard class, local fire design required',{coverage:'RFQ'});
 if(powerBudget!==null&&it>powerBudget)warnings.push('IT installed '+it+' kW exceeds reserved compute budget '+powerBudget+' kW after infrastructure loads; reduce GPU count');
 const cooling={mode,modeLabel:mode.toUpperCase(),rackDensityKw:maxRackKw,liquidKw,airKw,liquidFraction,flowLMin:flow,cduQty,cduDuty:cduN,rackLiquidCapacityKw:round(i.rackCoolingKw*liquidFraction),rackAirCapacityKw:round(i.rackCoolingKw*(1-liquidFraction)),rackLiquidKw:round(maxRackKw*liquidFraction),rackAirKw:round(maxRackKw*(1-liquidFraction)),notes:[{title:'Liquid / air heat',text:liquidKw+' / '+airKw+' kW; residual air includes network/storage'},{title:'Flow assumption',text:flow+' L/min = P/(ρ cp ΔT); ΔT '+i.deltaTK+' K; CDU '+cduN+' duty / '+cduQty+' installed; flow/pressure RFQ'},{title:'Rack binding constraint',text:binding.join(' + ')+'; '+perRack+' systems/rack; '+JSON.stringify(limits)}]};
 const coverage=bom.filter(x=>['Passive','Power','Facility','Ops'].includes(x.category)||/OOB/.test(x.item)).map(x=>({item:x.item,status:x.coverage||'calculated',statusClass:x.coverage==='RFQ'?'rfq':x.coverage==='option'?'optional':'direct',note:x.basis}));
 warnings.unshift('All coefficients are planning assumptions; manufacturer datasheet / OEM verification required.');
 const fabric={leafCount:counts.leaf,spineCount:counts.spine,coreCount:counts.core,serverLinks:counts.server,serverLeafLinks:counts.server,endpointLinks:counts.server,totalLeafUplinks:counts.uplink,leafSpineLinks:counts.uplink,spineToCoreLinks:counts.coreLink,spineCoreLinks:counts.coreLink,stageCount:Math.max(...domainRows.map(d=>d.tier)),linkSpeedGbps:p.linkSpeed,protocol:i.fabricProtocol,switchName:R+' logical-port '+p.linkSpeed+'G planning',serverFacingMedia:optical.server?.profile.media,leafSpineMedia:optical.leafSpine?.profile.media,spineDownPortsUsed:domainRows[0]?.leaf,coreUplinkPorts:counts.core?down:0,spineUnusedPorts:0,domains:domainRows};
 const allRequirements=bom.map(b=>{const o=optical[b.segment],isTransceiver=b.unit==='module'&&b.category==='Connectivity',isCable=['cable','assembly'].includes(b.unit),isTrunk=/trunk/.test(b.item);const profile=o?.profile||{optical:false,connector:'',protocol:b.category==='Network'?(/OOB|Frontend/.test(b.item)?'Ethernet':i.fabricProtocol):''};return {category:b.category,item:b.item,segment:b.segment||b.category,installed:b.installedQty,purchase:b.qty,unit:b.unit,lengthM:isCable?(o?(/patch leads/.test(b.item)?Number(i.fiberPatchM??2):o.lengthM):null):isTransceiver?o?.lengthM:null,fibers:isTrunk?o?.profile.trunkFiberCount:isCable?o?.profile.installedChannelFibers||0:0,activeFibers:o?.profile.fibers||0,profile,transceiver:isTransceiver,cable:isCable,polarity:o?.profile.optical?i.fiberPolarity:'',gender:o?.profile.optical?i.mpoGender||'review':'',jacket:isCable?i.cableJacket||'project':'',candidates:[]};});
 const products=bom.map(x=>({category:x.category,vendor:'RFQ',product:x.item,model:'RFQ',qty:x.qty,alternatives:[],evidence:'Exact SKU and compatibility unverified'}));
 const validation=Object.entries(optical).filter(([,o])=>o.links).map(([key,o])=>({label:key+' link conservation',actual:sum(routes.filter(x=>x.segment===key).map(x=>x.links)),expected:o.links,unit:'link',source:'',type:'Independent route sum',errorPct:0}));
 const networkTotals={switches:switchCount+storageSwitches+sum(auxFabrics.map(x=>x.switches)),modules:sum(bom.filter(x=>x.unit==='module'&&x.category==='Connectivity').map(x=>x.installedQty)),purchasedModules:sum(bom.filter(x=>x.unit==='module'&&x.category==='Connectivity').map(x=>x.qty)),bomLines:bom.length,powerKw:it};
 return {networkTotals,validation,version:'8.2.0-engineering',usable:true,status:errors.length?'ERROR / REVIEW':'REVIEW',input:i,warnings:[...new Set([...warnings,...errors])],errors,backend:{mode:'browser-engineering'},systemProfile:{...p,evidence:'Planning assumption',source:'https://docs.nvidia.com/dgx/'},summary:{targetGPU:target,installedGPU:units*p.gpus,systemUnits:units,servers:units,computeRacks:racks,networkRacks:allRacks.filter(r=>r.role==='Network').length,totalRacks:allRacks.length,totalItPowerKw:it,computePowerKw:computeKw,facilityPowerKw:facilityKw,systemsPerRack:perRack,bindingConstraint:binding,rackLimits:limits,inputItKw:i.inputMode==='power'?i.itInputMw*1000:null,computeBudgetKw:powerBudget},fabric,optical,facility,cooling,bom,products,coverage,interconnectGuide:guide,racks:allRacks,portAudit:{nodes:portNodes,routes,pass:!errors.some(x=>x.startsWith('Port capacity'))},additionalFabrics:auxFabrics,storage:{systems:storageNodes,switches:storageSwitches,rawRatio,rawTb:storageNodes*i.storageRawTbPerNode,usableTb:storageNodes*i.storageRawTbPerNode*rawRatio,readGbps:storageNodes*i.storageReadGbpsPerNode},planningScope:{frontend:i.frontendEnabled,oob:i.oobEnabled,storageFabric:i.storageEnabled},procurement,productRequirements:allRequirements,powerProof:{itKw:it,rackSumKw:it,residualKw:0,totalRacks:allRacks.length,averageRackKw:it/allRacks.length,computeKw,computeRacks:racks,averageComputeRackKw:computeKw/racks,pue:i.pue,overheadKw:facilityKw-it,upsKva,upsKw:upsKva*i.powerFactor,transformerKva,generatorKw,phaseBalance:'P/3 planning; outlet allocation RFQ'},cablingHandover:[],rackLayout:{rackRU:i.rackRU,components:p.rackScale?[{group:'Compute',label:'NVL72 OEM rack kit',startU:1,endU:Math.min(p.ru,i.rackRU)}]:Array.from({length:perRack},(_,j)=>({group:'Compute',label:p.name,startU:j*p.ru+1,endU:Math.min((j+1)*p.ru,i.rackRU)})),notes:cooling.notes},assumptions:{...i,system:p,optics:OPTICS,trayFillPct:fill}};
}
const api={calculate,SYSTEMS,OPTICS,DEFAULTS};if(typeof module!=='undefined'&&module.exports)module.exports=api;root.DCEngineering=api;
})(typeof window!=='undefined'?window:globalThis);
