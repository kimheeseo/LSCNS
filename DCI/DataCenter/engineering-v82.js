/* DC BOM engineering planning engine. All defaults are assumptions, not SKU qualification. */
(function(root){
'use strict';
const Calc=root.DCCalculations||(typeof require==='function'?require('./calculation-definitions.js'):null);
const RulesConfig=root.DCValidationConfig||(typeof require==='function'?require('./design-validation-config.js'):null);
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
// These entries constrain the cable-facing interface of a named module. The peer
// module, host cage/FEC and optical loss budget still need separate confirmation.
const OPTIC_SKUS={
 'nvidia-mma4z00-ns400':{name:'NVIDIA MMA4Z00-NS400',speed:400,mediaId:'sr',reach:50,connector:'MPO-12',polish:'APC',fiber:'OM4',protocols:['InfiniBand','Ethernet'],endGender:'female',source:'https://networking-docs.nvidia.com/mma4z00ns400'},
 'cisco-qsfp-400g-dr4':{name:'Cisco QSFP-400G-DR4',speed:400,mediaId:'dr',reach:500,connector:'MPO-12',polish:'APC',fiber:'OS2',protocols:['Ethernet'],source:'https://www.cisco.com/c/en/us/products/collateral/interfaces-modules/400g-qsfp-transceiver-modules-ds.html'},
 'cisco-qsfp-400g-fr4':{name:'Cisco QSFP-400G-FR4',speed:400,mediaId:'fr',reach:2000,connector:'Duplex LC',polish:'UPC',fiber:'OS2',protocols:['Ethernet'],source:'https://www.cisco.com/c/en/us/products/collateral/interfaces-modules/400g-qsfp-transceiver-modules-ds.html'}
};
const DEFAULTS={targetGPU:1024,systemId:'b200',inputMode:'gpu',itInputMw:100,nvlRackCount:4,topology:'rail',fabricProtocol:'InfiniBand',fabricTier:2,oversub:1,
 rackRU:48,reservedRU:4,rackPowerKw:100,rackCoolingKw:100,rackHeadroomPct:20,coolingMode:'auto',sparePct:10,
 serverDistanceM:30,leafSpineDistanceM:100,coreDistanceM:100,serverCabling:'structured',leafSpineCabling:'structured',coreCabling:'structured',trunkFiberCount:0,
 switchRadix:64,switchKw:2,switchRU:2,switchLogicalPerCage:1,serverLogicalPerModule:1,serverMedia:'auto',uplinkMedia:'auto',coreMedia:'auto',fiberPreference:'auto',
 slackPct:15,verticalRiseM:3,terminationM:0,matedPairs:4,p2pPairs:2,mpoLossDb:.35,lcLossDb:.25,lowLoss:false,
 fiberAttenDbKm:.35,mmfAttenDbKm:3,marginDb:0,spliceCount:0,spliceLossDb:.1,fiberPolarity:'B',connectorPolish:'auto',mpoEndGender:'review',serverOpticSku:'auto',uplinkOpticSku:'auto',coreOpticSku:'auto',
 pue:1.2,facilityHeadroomPct:20,powerFactor:.9,upsEfficiency:.96,upsBlockKva:1000,upsUtilPct:80,generatorBlockKw:2500,generatorUtilPct:80,transformerBlockKva:3000,transformerUtilPct:80,facilityRedundancy:'N+1',facilityEnabled:false,upsLoadScope:'it',upsInstalledQty:0,generatorInstalledQty:0,transformerInstalledQty:0,
 liquidPct:75,cduBlockKw:600,cduRedundancy:'N+1',waterDensity:1000,waterCp:4186,deltaTK:10,airLimitKw:RulesConfig.airLimitKw,rearDoorLimitKw:RulesConfig.rearDoorLimitKw,
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
 const positive=['rackRU','rackPowerKw','rackCoolingKw','switchRadix','powerFactor','upsEfficiency','upsBlockKva','generatorBlockKw','transformerBlockKva','cduBlockKw','waterDensity','waterCp','deltaTK','pduRatedKw','pduOutlets','panelChannels','cassetteChannels','storageRawTbPerNode','storageReadGbpsPerNode','auxRadix','consolePorts','trayWidthMm','trayDepthMm','siteVoltage','auxOversub','storageSwitchPorts'];
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
 const trace=Calc.createTrace(raw,i,DEFAULTS),F=(key,args,meta={})=>trace.run(key,args,meta);
 const field=name=>({field:name}),profileField=name=>({assumption:true,source:raw.systemAssumptions?.[i.systemId]?.[name]!==undefined?'Project system profile override':'engineering-v82.js / SYSTEMS.'+i.systemId+'.'+name});
 const red=(duty,mode)=>F('redundancy',{duty,mode});
 if(!['N','N+1','2N'].includes(i.facilityRedundancy))errors.push('Invalid facility redundancy: '+i.facilityRedundancy);
 if(!['it','facility'].includes(i.upsLoadScope)){warnings.push('Unknown UPS load scope; IT used');i.upsLoadScope='it';}
 const head=F('fraction',{reservePct:i.rackHeadroomPct},{id:'rack.headroom',inputs:{reservePct:field('rackHeadroomPct')}});
 const limits=p.rackScale?{RU:1,power:1,cooling:1}:{RU:F('rackRU',{rackRU:i.rackRU,reservedRU:i.reservedRU,systemRU:p.ru},{id:'rack.limitRU',inputs:{rackRU:field('rackRU'),reservedRU:field('reservedRU'),systemRU:profileField('ru')}}),power:F('rackPowerFit',{rackCapacity:i.rackPowerKw,headroom:head,systemPower:p.power},{id:'rack.limitPower',inputs:{rackCapacity:field('rackPowerKw'),headroom:'rack.headroom',systemPower:profileField('power')}}),cooling:F('rackPowerFit',{rackCapacity:i.rackCoolingKw,headroom:head,systemPower:p.power},{id:'rack.limitCooling',inputs:{rackCapacity:field('rackCoolingKw'),headroom:'rack.headroom',systemPower:profileField('power')}})};
 if(!p.rackScale&&i.topology!=='rail'){const maxNicPerUnit=i.topology==='dual'?ceil(p.links/2):p.links;limits.NIC=Math.floor(Math.floor(i.switchRadix/2)/maxNicPerUnit);}
 const fit=Math.min(...Object.values(limits)),perRack=F('rackFit',{limits:Object.values(limits)},{id:'rack.perRack',inputs:{limits:{source:'rack.limitRU / rack.limitPower / rack.limitCooling / topology NIC allocation',assumption:false}}}),binding=Object.keys(limits).filter(k=>limits[k]===fit);
 if(fit<1)errors.push('No server fits rack limits; one per rack shown for review');
 if(i.rackPowerKw>i.rackCoolingKw)errors.push('Rack power rating exceeds rack cooling rating');
 let units,racks,target=i.targetGPU,powerBudget=null;
 if(p.rackScale){const numerator=i.inputMode==='power'?F('budget',{inputMW:i.itInputMw,headroom:head},{id:'compute.budget',inputs:{inputMW:field('itInputMw'),headroom:'rack.headroom'}}):i.inputMode==='racks'?Math.max(1,i.nvlRackCount):target;racks=Math.max(1,F('ceilRatio',{numerator,denominator:i.inputMode==='power'?p.power:i.inputMode==='racks'?1:p.gpus},{id:'compute.units',inputs:{numerator:i.inputMode==='power'?'compute.budget':field(i.inputMode==='racks'?'nvlRackCount':'targetGPU'),denominator:profileField(i.inputMode==='power'?'power':'gpus')}}));units=racks;target=i.inputMode==='gpu'?target:racks*p.gpus;}
 else if(i.inputMode==='power'){powerBudget=F('budget',{inputMW:i.itInputMw,headroom:head},{id:'compute.budget',inputs:{inputMW:field('itInputMw'),headroom:'rack.headroom'}});racks=Math.max(1,ceil(powerBudget/i.rackPowerKw));units=F('powerSystems',{racks,perRack,budget:powerBudget,power:p.power},{id:'compute.units',inputs:{perRack:'rack.perRack',budget:'compute.budget',power:profileField('power')}});target=units*p.gpus;}
 else {units=Math.max(1,F('ceilRatio',{numerator:target,denominator:p.gpus},{id:'compute.units',inputs:{numerator:field('targetGPU'),denominator:profileField('gpus')}}));racks=F('ceilRatio',{numerator:units,denominator:perRack},{id:'compute.racks',inputs:{numerator:'compute.units',denominator:'rack.perRack'}});}
 if(units>200000){errors.push('Planning limit 200000 systems; reduce input');units=200000;racks=p.rackScale?units:ceil(units/perRack);target=units*p.gpus;}
 const computeRacks=Array.from({length:racks},(_,idx)=>{const n=p.rackScale?1:Math.max(0,Math.min(perRack,units-idx*perRack));return {id:'compute-'+(idx+1),role:'Compute',units:n,ru:p.rackScale?p.ru:n*p.ru,powerKw:round(F('multiply',{count:n,rating:p.power},{id:'rack.'+(idx+1)+'.power',scope:'rack',unit:'kW',inputs:{count:{source:'Compute rack distribution of compute.units / rack.perRack',assumption:false},rating:profileField('power')}})),cords:n*p.cordCount};});
 if(p.rackScale){if(p.power>i.rackPowerKw)errors.push('NVL72 rack kW exceeds configured rack power');if(p.power>i.rackCoolingKw)errors.push('NVL72 rack kW exceeds configured cooling');}
 const maxRackKw=computeRacks.reduce((v,r)=>Math.max(v,r.powerKw),0);
 if(mode==='air'&&maxRackKw>i.airLimitKw)warnings.push('Air cooling planning threshold exceeded: '+maxRackKw+' > '+i.airLimitKw+' kW');
 if(mode==='rear-door'&&maxRackKw>i.rearDoorLimitKw)warnings.push('Rear-door HX planning threshold exceeded: '+maxRackKw+' > '+i.rearDoorLimitKw+' kW');
 // Each domain has its own fabric. ToR modes preserve rack affinity; Dual splits existing NICs across A/B fabrics.
 const R=Math.max(4,Math.floor(i.switchRadix/2)*2),down=F('portsDown',{radix:R},{id:'fabric.down',inputs:{radix:field('switchRadix')}}),up=F('portsUp',{downPorts:down,oversubscription:i.oversub},{id:'fabric.up',inputs:{downPorts:'fabric.down',oversubscription:field('oversub')}});
 if(i.twinPort400&&p.linkSpeed===400){i.switchLogicalPerCage=2;warnings.push('Twin-port option assumes two qualified 400G logical ports/switch cage');}
 if(![1,2].includes(i.serverLogicalPerModule)){warnings.push('Server module port count must be 1 or 2; 1 used');i.serverLogicalPerModule=1;}
 if(i.serverLogicalPerModule===2&&!i.twinPort400)warnings.push('Server dual-port module count requires the qualified 2×400G physical module option; no sharing applied');
 const domains=i.topology==='rail'?(p.rackScale?p.rails||4:p.links):i.topology==='dual'?2:1;
 if(R!==i.switchRadix)warnings.push('Switch radix rounded to even '+R);i.switchRadix=R;
 const endpoints=F('multiply',{count:units,rating:p.links},{id:'fabric.endpoints',unit:'port',inputs:{count:'compute.units',rating:profileField('links')}}),portNodes=[],routes=[],domainRows=[],counts={leaf:0,spine:0,core:0,server:0,uplink:0,coreLink:0};
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
  F('leafGroups',{groups:groups.map(g=>g.n),downPorts:down},{id:'domain.'+domain+'.leaves',inputs:{groups:{source:'fabric.endpoints / independent domains; Rail pools endpoints; ToR preserves each rack',assumption:false},downPorts:'fabric.down'}});
  const hostCount=sum(groups.map(g=>g.n)),twoLimit=R*down,threeLimit=R*down*down;
  let tier=Number(i.fabricTier)===3?3:2;
  if(leaves.length>R||hostCount>twoLimit){tier=3;warnings.push('Domain '+domain+': 2-tier capacity exceeded; promoted to 3-tier');}
  if(hostCount>threeLimit||leaves.length>R*down)errors.push('Domain '+domain+': exceeds 3-tier capacity; additional independent fabrics required');
  let spineCount=0,coreCount=0,coreLinks=0;
  if(tier===2){
   const compactSpines=F('compactPlanes',{uplinks:up,fanIn:leaves.length,capacity:R},{id:'domain.'+domain+'.spinePlanes',inputs:{uplinks:'fabric.up',fanIn:'domain.'+domain+'.leaves',capacity:field('switchRadix')}});
   const spines=Array.from({length:compactSpines},(_,j)=>addNode('D'+domain+'-S'+j,'Spine',R));
   for(const leaf of leaves)for(const spine of spines)route(leaf,spine,up/compactSpines,'leafSpine');
   spineCount=spines.length;
  }else{
   // Regular folded Clos: <= R/2 leaves/pod; up spine planes/pod; each spine reserves leafCount upstream ports.
   const podLeafCap=Math.max(1,Math.min(down,Math.floor(i.podLeafLimit||down)));
   const pods=[];for(let at=0;at<leaves.length;at+=podLeafCap)pods.push(leaves.slice(at,at+podLeafCap));
   if(pods.length>R)errors.push('Domain '+domain+': configured pod Leaf limit creates more than R pods; split fabric');
   const maxLeaves=Math.max(1,...pods.map(x=>x.length));
   const podSpines=F('compactPlanes',{uplinks:up,fanIn:maxLeaves,capacity:down},{id:'domain.'+domain+'.podSpines',inputs:{uplinks:'fabric.up',fanIn:{source:'Maximum leaf count in each pod; podLeafLimit='+i.podLeafLimit,assumption:false},capacity:'fabric.down'}});
   const parallelLeafLinks=up/podSpines,spineUp=maxLeaves*parallelLeafLinks;
   // Equal parallel links within a core plane. Choose divisor so every core respects radix.
   const coresPerPlane=F('compactPlanes',{uplinks:spineUp,fanIn:pods.length,capacity:R},{id:'domain.'+domain+'.coresPerPlane',inputs:{uplinks:{source:'maxLeaves × fabric.up / domain.'+domain+'.podSpines',assumption:false},fanIn:{source:'ceil(domain leaves / podLeafLimit)',assumption:false},capacity:field('switchRadix')}});
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
 // Final quantities are derived from the same route/port graph shown to the user.
 for(const [k,role]of [['leaf','Leaf'],['spine','Spine'],['core','Core']])counts[k]=F('countRole',{nodes:portNodes,selectedRole:role},{id:'fabric.'+k,inputs:{nodes:{source:'Allocated port graph (topology, switchRadix, NIC endpoints, oversubscription, tier)',assumption:false},selectedRole:{source:'Topology role',assumption:false}}});
 for(const [k,segment]of [['server','server'],['uplink','leafSpine'],['coreLink','core']])counts[k]=F('routeLinks',{routes,selectedSegment:segment},{id:'fabric.links.'+segment,inputs:{routes:{source:'Allocated topology routes with endpoint port reservations',assumption:false},selectedSegment:{source:'Topology segment',assumption:false}}});
 const bom=[],procurement=[],optical={},requirements=[],guide=[],purchase=(n,pack=1)=>{const spare=F('spareUnits',{installed:n,sparePct:i.sparePct},{scope:'procurement',inputs:{installed:{source:'Installed quantity from current route / equipment graph',assumption:false},sparePct:field('sparePct')}}),spareId=trace.steps.at(-1).id;F('withSpare',{installed:n,spare},{scope:'procurement',inputs:{installed:{source:'Installed route / equipment count',assumption:false},spare:spareId}});return F('purchase',{installed:n,sparePct:i.sparePct,pack},{scope:'procurement',inputs:{installed:{source:'Installed quantity; before spare '+n,assumption:false},sparePct:field('sparePct'),pack:{source:'Project pack size (modulePack / trunkBundle / channelPack / assemblyPack); default = 1',assumption:true}}});},row=(category,item,n,unit,basis,extra={})=>{const previous=trace.steps.at(-1),qty=F('bomQuantity',{calculatedQuantity:n},{id:'bom.'+bom.length,scope:'bom',label:item,unit,inputs:{calculatedQuantity:previous?.value===n?previous.id:{source:'engineering-v82.js: '+basis,assumption:true}}});bom.push({category,item,qty,evidenceId:'bom.'+bom.length,installedQty:extra.installedQty??qty,spareQty:extra.installedQty===undefined?0:qty-extra.installedQty,unit,basis,...extra});};
 row('Compute',p.name,units,p.rackScale?'rack':'system',(p.rackScale?'ceil(GPU/72)':'ceil(GPU/'+p.gpus+')')+'; configured power '+p.power+' kW');
 row('Network','Compute NIC logical endpoints',endpoints,'port',units+' × '+p.links+' NIC ports');
 for(const [key,n]of[['Leaf',counts.leaf],['Spine',counts.spine],['Super-spine',counts.core]])if(n)row('Network',key+' '+p.linkSpeed+'G planning switch',n,'switch','Sum independent domains; R='+R+', down='+down+', up='+up+'; no spare');
 if(p.rackScale){row('Compute','NVLink copper backplane assembly',racks,'rack kit','1 OEM rack kit/rack; RFQ, not scale-out optics');row('Compute','NVLink switch trays (included OEM rack)',9*racks,'tray','9/rack reference, included in rack kit; do not double-order');}
 function link(segment,links,speed,distance,cabling,mediaChoice){
  if(!links)return;
  const routed=F('routeLength',{route:distance,rise:i.verticalRiseM,termination:i.terminationM,slackPct:i.slackPct},{id:segment+'.route',scope:'optical',inputs:{route:field({server:'serverDistanceM',leafSpine:'leafSpineDistanceM',core:'coreDistanceM'}[segment]),rise:field('verticalRiseM'),termination:field('terminationM'),slackPct:field('slackPct')}}),standard=RulesConfig.standardLengthsM.find(m=>m>=routed),length=F('standardLength',{routeLength:routed,standardLengths:RulesConfig.standardLengthsM,increment:RulesConfig.customLengthIncrementM},{id:segment+'.length',scope:'optical',inputs:{routeLength:segment+'.route',standardLengths:{source:RulesConfig.source,assumption:true},increment:{source:RulesConfig.source,assumption:true}}});
  if(!standard)warnings.push(segment+': route >100m; custom length '+length+'m RFQ');
  const table=OPTICS.filter(x=>x.speed===speed).map(x=>({...x,reach:raw.opticsAssumptions?.[speed+'-'+x.id]?.reach??x.reach,loss:raw.opticsAssumptions?.[speed+'-'+x.id]?.loss??x.loss}));
  const skuId=i[{server:'serverOpticSku',leafSpine:'uplinkOpticSku',core:'coreOpticSku'}[segment]];
  const requestedSku=skuId&&skuId!=='auto'?OPTIC_SKUS[skuId]:null;
  let sku=requestedSku;
  if(skuId&&skuId!=='auto'&&!sku){errors.push(segment+': unknown optic SKU '+skuId+'; generic profile is unqualified');}
  if(sku&&sku.speed!==speed){errors.push(segment+': '+sku.name+' requires '+sku.speed+'G, not '+speed+'G');sku=null;}
  if(sku&&!sku.protocols.includes(i.fabricProtocol))errors.push(segment+': '+sku.name+' has no published '+i.fabricProtocol+' protocol qualification');
  if(sku&&mediaChoice&&mediaChoice!=='auto'&&mediaChoice!==sku.mediaId)errors.push(segment+': requested '+mediaChoice+' media conflicts with '+sku.name);
  const lossOf=x=>{const optical=x.loss!==null,budget=optical&&i.opticalBudgetMode==='project'?Math.min(x.loss,Number(i.projectBudgetDb)||x.loss):x.loss,pairs=cabling==='structured'?Math.max(4,ceil(i.matedPairs)):Math.max(2,ceil(i.p2pPairs)),atten=x.fiber==='OM4'?i.mmfAttenDbKm:i.fiberAttenDbKm,connector=/LC/.test(x.connector)?i.lcLossDb:i.lowLoss?.25:i.mpoLossDb;return {applicable:optical,matedPairs:pairs,attenDbKm:atten,connectorDb:connector,estimatedDb:optical?round(Calc.evaluate('opticalLoss',{attenuation:atten,length,pairs,connectorLoss:connector,splices:i.spliceCount,spliceLoss:i.spliceLossDb})):0,budgetDb:budget,marginDb:i.marginDb,pass:length<=x.reach&&(!optical||Calc.evaluate('opticalLoss',{attenuation:atten,length,pairs,connectorLoss:connector,splices:i.spliceCount,spliceLoss:i.spliceLossDb})+i.marginDb<=budget)};};
  const available=table.filter(x=>cabling==='p2p'||x.loss!==null).filter(x=>i.fiberPreference==='smf'?x.loss===null||x.fiber==='OS2':i.fiberPreference==='mmf'?x.loss===null||x.fiber==='OM4':true);
  let chosen=sku?table.find(x=>x.id===sku.mediaId):mediaChoice&&mediaChoice!=='auto'?table.find(x=>x.id===mediaChoice):available.filter(x=>x.id!=='aec').find(x=>lossOf(x).pass);
  chosen=chosen||available.at(-1)||table.at(-1);
  if(sku)chosen={...chosen,reach:sku.reach,connector:sku.connector,polish:sku.polish,fiber:sku.fiber};
  const loss=lossOf(chosen),isOptical=chosen.loss!==null;
  const alternatives=table.filter(x=>!sku||x.id!==sku.mediaId).filter(x=>lossOf(x).pass&&(cabling==='p2p'||x.loss!==null)).map(x=>x.standard||x.id.toUpperCase());
  if(!loss.pass)errors.push(segment+': '+(sku?.name||chosen.standard||chosen.id)+' FAIL; '+(sku?'different qualified endpoint SKU required; generic media possibilities ':'alternatives ')+(alternatives.join(', ')||'reduce connector pairs / direct cabling / vendor-qualified long reach'));
  if(isOptical&&!sku)warnings.push(segment+': generic optical reach/connector only; select and verify both endpoint optic SKUs before procurement');
  if(isOptical&&i.marginDb===0)warnings.push(segment+': zero optical loss reserve; set a project margin before installation approval');
  if(sku&&((i.fiberPreference==='smf'&&chosen.fiber==='OM4')||(i.fiberPreference==='mmf'&&chosen.fiber==='OS2')))errors.push(segment+': '+sku.name+' fiber type '+chosen.fiber+' conflicts with '+i.fiberPreference+' preference');
  if(cabling==='structured'&&!isOptical)errors.push(segment+': DAC/AOC/AEC cannot traverse passive patch panels; select point-to-point');
  if(isOptical&&i.fiberPolarity!=='B'&&/MPO/.test(chosen.connector))errors.push(segment+': Type-B end-to-end polarity required by planning profile; verify assembly mapping');
  if(isOptical&&i.connectorPolish!=='auto'&&i.connectorPolish!==chosen.polish)errors.push(segment+': '+i.connectorPolish+' polish conflicts with '+chosen.polish+' profile; never mate APC/UPC');
  if(sku?.endGender&&i.mpoEndGender!=='review'&&i.mpoEndGender!==sku.endGender)errors.push(segment+': '+sku.name+' requires '+sku.endGender+' cable connector at the transceiver');
  if(sku?.endGender&&i.mpoEndGender==='review')warnings.push(segment+': '+sku.name+' transceiver end requires '+sku.endGender+' MPO cable; pinning/adapter mapping RFQ');
  const groups=chosen.groups||1,base=chosen.base||0,channelFibers=base*groups;
  const requestedTrunk=i.trunkFiberCount||channelFibers,channelsPerTrunk=isOptical?Math.floor(requestedTrunk/channelFibers):0;
  const actualTrunk=channelsPerTrunk>0?requestedTrunk:channelFibers;
  if(isOptical&&cabling==='structured'&&!channelsPerTrunk)warnings.push(segment+': trunk too small; '+actualTrunk+'F required per channel; RFQ');
  if(isOptical&&cabling==='structured'&&actualTrunk%channelFibers)warnings.push(segment+': '+actualTrunk+'F trunk leaves '+(actualTrunk%channelFibers)+'F outside complete '+channelFibers+'F channels; exact MPO breakout/termination RFQ');
  // Share qualified 2x400G switch modules per physical switch; count server
  // modules per actual system rather than pooling their ports across systems.
  let installedModules=isOptical?2*links:0;
  if(isOptical&&speed===400&&i.twinPort400){
   const field=segment==='server'?'server':segment==='leafSpine'?'leafSpine':'core';
   const eligible=portNodes.filter(n=>segment==='server'?n.role==='Leaf':segment==='leafSpine'?['Leaf','Spine'].includes(n.role):['Spine','Core'].includes(n.role));
   installedModules=segment==='server'?units*ceil(p.links/i.serverLogicalPerModule):0;installedModules+=sum(eligible.map(n=>ceil(n.used[field]/2)));
   warnings.push(segment+': 2×400G modules assumed; server '+i.serverLogicalPerModule+' logical ports/module; matched breakout/FEC/cage/connector SKU required');
  }
  const trunkCapacity=isOptical?Math.max(1,F('channelsPerTrunk',{trunkFibers:actualTrunk,allocatedFibersPerChannel:channelFibers},{id:segment+'.trunkCapacity',scope:'optical',inputs:{trunkFibers:field('trunkFiberCount'),allocatedFibersPerChannel:{source:sku?.source||'engineering-v82.js / OPTICS: connector base × groups',assumption:!sku}}})):1;
  const segmentRoutes=routes.filter(q=>q.segment===segment);
  const trunkInstalled=isOptical&&cabling==='structured'?F('trunks',{routeLinks:segmentRoutes.map(q=>q.links),channelsPerTrunk:trunkCapacity},{id:segment+'.trunksInstalled',scope:'optical',inputs:{routeLinks:{ref:'fabric.links.'+segment,source:'Independent endpoint routes; do not pool unrelated route fibers',assumption:false},channelsPerTrunk:segment+'.trunkCapacity'}}):0;
  const selectedBase=i[{server:'serverBase',leafSpine:'uplinkBase',core:'coreBase'}[segment]];
  if(selectedBase&&selectedBase!=='auto'&&Number(selectedBase)!==base)warnings.push(segment+': Base-'+selectedBase+' conflicts with planned Base-'+base+'; RFQ interface redesign');
  F('opticalLoss',{attenuation:loss.attenDbKm,length,pairs:loss.matedPairs,connectorLoss:loss.connectorDb,splices:i.spliceCount,spliceLoss:i.spliceLossDb},{id:segment+'.loss',scope:'optical',inputs:{length:segment+'.length',attenuation:field(chosen.fiber==='OM4'?'mmfAttenDbKm':'fiberAttenDbKm'),pairs:field(cabling==='structured'?'matedPairs':'p2pPairs'),connectorLoss:field(/LC/.test(chosen.connector)?'lcLossDb':'mpoLossDb'),splices:field('spliceCount'),spliceLoss:field('spliceLossDb')}});
  const profile={speed,mediaId:chosen.id,reachM:chosen.reach,standard:(chosen.standard||'').replace(/^\d+G-/,''),protocol:i.fabricProtocol,package:'',media:chosen.standard||speed+'G '+chosen.id.toUpperCase(),optical:isOptical,optics:isOptical?2:0,connector:chosen.connector,polish:chosen.polish,fibers:chosen.fibers,installedChannelFibers:channelFibers,base,connectorGroups:groups,fiberType:chosen.fiber,trunkFiberCount:actualTrunk,requestedTrunkFiberCount:i.trunkFiberCount,unusedTrunkFibers:actualTrunk%Math.max(1,channelFibers),opticSku:sku?.name||'',source:sku?.source||'',compatibility:'REVIEW',rule:sku?'SKU reach/connector only; loss budget, peer optic and host/FEC RFQ':'Generic assumption; select both endpoint optic SKUs and verify host/FEC'};
  optical[segment]={links,lengthM:length,routeLengthM:routed,profile,installedModules,purchasedModules:purchase(installedModules,i.modulePack),loss:{...loss,alternatives},trunk:{installedCableCount:trunkInstalled,cableCount:purchase(trunkInstalled,i.trunkBundle),installedChannelFibers:channelFibers,requiredFibers:links*chosen.fibers,provisionedFibers:trunkInstalled*actualTrunk}};
  guide.push({speed:speed+'G',media:profile.media,distanceM:length,connector:profile.connector,rule:(sku?sku.name+' · ':'Generic profile · ')+'Loss '+loss.estimatedDb+'/'+loss.budgetDb+' dB; '+(loss.pass?'PRELIMINARY / RFQ':'FAIL')+'; peer module and channel budget unverified',segment});
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
 const addEquipmentRacks=(role,count,ru,kw,cords=2)=>{const capacity=Math.max(1,Math.min(Math.floor(Math.max(1,i.rackRU-i.reservedRU)/ru),Math.floor(Math.min(i.rackPowerKw,i.rackCoolingKw)*head/kw)));for(let at=0;at<count;at+=capacity){const n=Math.min(capacity,count-at);allRacks.push({id:role+'-'+allRacks.length,role,units:n,ru:n*ru,powerKw:round(F('multiply',{count:n,rating:kw},{id:'rack.'+role+'-'+allRacks.length+'.power',unit:'kW',inputs:{count:{source:'Equipment rack allocation from equipment count and rack limits',assumption:false},rating:{source:'Project equipment power / DEFAULTS: switchKw, auxSwitchKw, storageArrayKw, storageSwitchKw, opsUnitKw',assumption:true}}})),cords:n*cords});}};
 const switchCount=counts.leaf+counts.spine+counts.core;addEquipmentRacks('Network',switchCount,i.switchRU,i.switchKw);
 const rawRatio=i.storageEncoding==='replica3'?1/3:.8;
 const storageNodes=i.storageEnabled?F('storageNodes',{capacity:i.storageCapacityTb,rawPerNode:i.storageRawTbPerNode,usableRatio:rawRatio,GPU:target,readPerGPU:i.storageReadGbpsPerGpu,readPerNode:i.storageReadGbpsPerNode},{id:'storage.nodes',inputs:{capacity:field('storageCapacityTb'),rawPerNode:field('storageRawTbPerNode'),usableRatio:{source:'engineering-v82.js: replica3 = 1/3; EC = 0.8 planning assumption',assumption:true},GPU:field('targetGPU'),readPerGPU:field('storageReadGbpsPerGpu'),readPerNode:field('storageReadGbpsPerNode')}}):0;
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
 const componentPower={};for(const role of ['Compute','Network','Storage','Ops'])componentPower[role]=F('sum',{values:allRacks.filter(r=>r.role===role).map(r=>r.powerKw)},{id:'power.component.'+role,unit:'kW',inputs:{values:{source:'Current rack schedule for '+role+'; component count × configured kW',assumption:false}}});
 const it=round(F('itPower',componentPower,{id:'power.IT',inputs:Object.fromEntries(Object.keys(componentPower).map(role=>[role,'power.component.'+role]))})),facilityKw=round(F('facilityPower',{ITpower:it,PUE:i.pue},{id:'power.facility',inputs:{ITpower:'power.IT',PUE:field('pue')}})),facilityMargin=1+i.facilityHeadroomPct/100;
 const protectedLoad=i.upsLoadScope==='facility'?facilityKw:it;
 const upsKva=F('upsDesign',{protectedLoad,powerFactor:i.powerFactor,efficiency:i.upsEfficiency,headroomPct:i.facilityHeadroomPct},{id:'facility.upsDesign',inputs:{protectedLoad:i.upsLoadScope==='facility'?'power.facility':'power.IT',powerFactor:field('powerFactor'),efficiency:field('upsEfficiency'),headroomPct:field('facilityHeadroomPct')}}),generatorKw=F('generatorDesign',{facilityPower:facilityKw,headroomPct:i.facilityHeadroomPct},{id:'facility.generatorDesign',inputs:{facilityPower:'power.facility',headroomPct:field('facilityHeadroomPct')}}),transformerKva=F('transformerDesign',{facilityPower:facilityKw,powerFactor:i.powerFactor,headroomPct:i.facilityHeadroomPct},{id:'facility.transformerDesign',inputs:{facilityPower:'power.facility',powerFactor:field('powerFactor'),headroomPct:field('facilityHeadroomPct')}});
 const facility={pue:i.pue,itPowerKw:it,facilityPowerKw:facilityKw,upsLoadScope:i.upsLoadScope,protectedLoadKw:protectedLoad,mechanicalAndLossKw:round(facilityKw-it),upsDesignKva:round(upsKva),upsDesignKw:round(upsKva*i.powerFactor),generatorDesignKw:round(generatorKw),transformerDesignKva:round(transformerKva)};
 for(const [kind,load,block,util]of [['ups',upsKva,i.upsBlockKva,i.upsUtilPct],['generator',generatorKw,i.generatorBlockKw,i.generatorUtilPct],['transformer',transformerKva,i.transformerBlockKva,i.transformerUtilPct]]){
  const usable=F('usableBlock',{blockRating:block,utilizationPct:util},{id:'facility.'+kind+'Block',unit:kind==='generator'?'kW':'kVA',inputs:{blockRating:field(kind+'Block'+(kind==='generator'?'Kw':'Kva')),utilizationPct:field(kind+'UtilPct')}});
  const duty=F('ceilRatio',{numerator:load,denominator:usable},{id:'facility.'+kind+'Duty',inputs:{numerator:'facility.'+kind+'Design',denominator:'facility.'+kind+'Block'}});
  facility[kind+'Duty']=duty;facility[kind+'UsableBlock']=usable;
  const autoQuantity=F('redundancy',{duty,mode:i.facilityRedundancy},{id:'facility.'+kind+'AutoQty',inputs:{duty:'facility.'+kind+'Duty',mode:field('facilityRedundancy')}});
  facility[kind+'Qty']=F('installedModules',{autoQuantity,override:i[kind+'InstalledQty']},{id:'facility.'+kind+'Qty',inputs:{autoQuantity:'facility.'+kind+'AutoQty',override:field(kind+'InstalledQty')}});
 }
 if(i.facilityEnabled){row('Facility','UPS '+i.upsBlockKva+' kVA module',facility.upsQty,'module','red(ceil(protected '+protectedLoad+'/PF '+i.powerFactor+'/η '+i.upsEfficiency+'×'+facilityMargin+'/usable '+i.upsBlockKva*i.upsUtilPct/100+'))');row('Facility','Generator '+i.generatorBlockKw+' kW module',facility.generatorQty,'module','red(ceil(IT×PUE×headroom/usable kW)); mechanical load included');row('Facility','Transformer '+i.transformerBlockKva+' kVA module',facility.transformerQty,'module','red(ceil(IT×PUE/PF×headroom/(rated×derating)))');}
 const computeKw=F('multiply',{count:units,rating:p.power},{id:'power.compute',unit:'kW',inputs:{count:'compute.units',rating:profileField('power')}}),liquidFraction=mode==='immersion'?1:mode==='dlc'?i.liquidPct/100:mode==='rear-door'?i.liquidPct/100:0;
 const liquidKw=round(F('multiply',{count:computeKw,rating:liquidFraction},{id:'cooling.liquid',unit:'kW',inputs:{count:'power.compute',rating:{source:'coolingMode and liquidPct (project / DEFAULTS)',assumption:!Object.hasOwn(raw,'liquidPct')}}})),airKw=round(it-liquidKw),flow=round(F('flow',{liquidKW:liquidKw,density:i.waterDensity,heatCapacity:i.waterCp,deltaT:i.deltaTK},{id:'cooling.flow',inputs:{liquidKW:'cooling.liquid',density:field('waterDensity'),heatCapacity:field('waterCp'),deltaT:field('deltaTK')}})),cduN=F('ceilRatio',{numerator:liquidKw,denominator:i.cduBlockKw},{id:'cooling.cduDuty',inputs:{numerator:'cooling.liquid',denominator:field('cduBlockKw')}}),cduQty=red(cduN,i.cduRedundancy);
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
 return {evidence:{...trace.snapshot(),rackSchedule:allRacks,systemProfile:p},networkTotals,validation,version:'8.2.0-engineering',usable:true,status:errors.length?'ERROR / REVIEW':'REVIEW',input:i,warnings:[...new Set([...warnings,...errors])],errors,backend:{mode:'browser-engineering'},systemProfile:{...p,evidence:'Planning assumption',source:'https://docs.nvidia.com/dgx/'},summary:{targetGPU:target,installedGPU:units*p.gpus,systemUnits:units,servers:units,computeRacks:racks,networkRacks:allRacks.filter(r=>r.role==='Network').length,totalRacks:allRacks.length,totalItPowerKw:it,computePowerKw:computeKw,facilityPowerKw:facilityKw,systemsPerRack:perRack,bindingConstraint:binding,rackLimits:limits,inputItKw:i.inputMode==='power'?i.itInputMw*1000:null,computeBudgetKw:powerBudget},fabric,optical,facility,cooling,bom,products,coverage,interconnectGuide:guide,racks:allRacks,portAudit:{nodes:portNodes,routes,pass:!errors.some(x=>x.startsWith('Port capacity'))},additionalFabrics:auxFabrics,storage:{systems:storageNodes,switches:storageSwitches,rawRatio,rawTb:storageNodes*i.storageRawTbPerNode,usableTb:storageNodes*i.storageRawTbPerNode*rawRatio,readGbps:storageNodes*i.storageReadGbpsPerNode},planningScope:{frontend:i.frontendEnabled,oob:i.oobEnabled,storageFabric:i.storageEnabled},procurement,productRequirements:allRequirements,powerProof:{itKw:it,rackSumKw:it,residualKw:0,totalRacks:allRacks.length,averageRackKw:it/allRacks.length,computeKw,computeRacks:racks,averageComputeRackKw:computeKw/racks,pue:i.pue,overheadKw:facilityKw-it,upsKva,upsKw:upsKva*i.powerFactor,transformerKva,generatorKw,phaseBalance:'P/3 planning; outlet allocation RFQ'},cablingHandover:[],rackLayout:{rackRU:i.rackRU,components:p.rackScale?[{group:'Compute',label:'NVL72 OEM rack kit',startU:1,endU:Math.min(p.ru,i.rackRU)}]:Array.from({length:perRack},(_,j)=>({group:'Compute',label:p.name,startU:j*p.ru+1,endU:Math.min((j+1)*p.ru,i.rackRU)})),notes:cooling.notes},assumptions:{...i,system:p,optics:OPTICS,trayFillPct:fill}};
}
const api={calculate,SYSTEMS,OPTICS,OPTIC_SKUS,DEFAULTS};if(typeof module!=='undefined'&&module.exports)module.exports=api;root.DCEngineering=api;
})(typeof window!=='undefined'?window:globalThis);
