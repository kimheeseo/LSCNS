// Universal per-component picking for the educational data-center rack model.
// Only modeled meshes can be selected. Does not invent physical parts or specifications.
const supported=/^(rack-shape:|rack-server:|rack-link:|rack-optical:|rack-optical-patch:|rack-lc:|odf-small:|hardware:)/;
const preserve=new Set(['cassette-cable','fiber-port','mpo-input','patch-internal','lc-adapter-port','odf-small-port','fiber-patch-cable','dac-cable','power-cable','patch-cord','optical-connector','tor-port']);
const aliases=[
 [/^fan-(cartridge|grille|blade|hub)/,'냉각 팬 부품'],
 [/^(green-mainboard|gpu-carrier-board|nic-board|drive-controller-pcb)/,'PCB 회로기판'],
 [/^(memory-dimm|dimm-socket)/,'메모리 DIMM'],
 [/^(chip-package|metal-package-cap)/,'프로세서 패키지'],
 [/^(heatsink|gpu-coldplate)/,'방열판·냉각판'],
 [/^(redundant-psu|psu-|pdu-)/,'전원 공급 부품'],
 [/^(optical|mpo-|lc-|adapter-|transceiver-)/,'광통신 부품'],
 [/^(server-port|fabric-port|tor-port|nic-cage|nic-latch|io-port)/,'네트워크 포트'],
 [/^(storage-drive|hot-swap-drive|drive-release)/,'스토리지 부품'],
 [/^(front-door|rear-door|rack-ear|upright)/,'랙 기구부품'],
 [/^(rail-|sled-rail|fixed-guide|cassette-pull)/,'슬라이딩 레일']
];
const readable=raw=>String(raw||'부품').replace(/^(rack-shape:|rack-server:|rack-link:|rack-optical:|rack-optical-patch:|rack-lc:|odf-small:|hardware:)/,'').replace(/[-_]/g,' ').replace(/\s+/g,' ').trim();
const labelFor=(name,parent)=>{
 const raw=name.split(':').slice(1).join(':')||name;
 const match=aliases.find(([rx])=>rx.test(raw));
 return (parent?.label?parent.label+' · ':'')+(match?match[1]+' / ':'')+readable(raw);
};
export function registerRackSurfaceParts(asset,record){
 if(!record?.rackParts||!record.partMap||!record.rackLOD?.high)return {added:0,removed:0,total:0};
 if(!record.universalParts)record.universalParts=new Map();
 const active=record.universalParts,seen=new Set(),groups=new Map(),original=record.partMap;
 let added=0,removed=0,index=0;
 record.rackLOD.high.traverse(object=>{
  if(!object.isMesh&&!object.isInstancedMesh)return;
  if(object.userData?.noPick||object.userData?.cablePickProxy||object.userData?.cassetteCable||!supported.test(object.name||''))return;
  // Existing optical port, MPO connector and grouped cable IDs must retain their
  // physical identity. Do not turn individual optical cores into pick targets.
  const sourceId=object.userData.universalSourcePartId??object.userData.partId;
  const known=sourceId&&original.get(sourceId);
  if(known&&preserve.has(known.kind))return;
  let group=null;
  for(let p=object.parent;p&&p!==record.g;p=p.parent){
   if(p.userData?.rackSelectableGroup){group=p;break;}
  }
  const key=group?'group:'+group.uuid:'mesh:'+object.uuid;
  let info=groups.get(key);
  if(!info){
   const id=asset.id+':surface:'+key;
   const parent=known||null;
   info=active.get(id)?.info||{
    id,assetId:asset.id,kind:'rack-component',
    label:group?.userData?.rackSelectableLabel||labelFor(object.name,parent),
    description:'랙 3D 교육용 모델에서 개별 선택할 수 있는 부품입니다. 모델에 표현된 형상이며 제조사 검증 사양이나 실제 BOM이 아닙니다.',
    role:'랙 내부 구성 부품의 위치와 주변 연결 관계를 확인합니다.',
    component:group?.userData?.rackSelectableLabel||readable(object.name),
    parentAssembly:parent?.label||'랙 내부',
    insideHardwareDetail:String(object.name||'').startsWith('hardware:'),
    uStart:parent?.uStart,uHeight:parent?.uHeight,
    universal:true
   };
   if(!original.has(id)){record.rackParts.push(info);original.set(id,info);added++;}
   active.set(id,{info,object:group||object});
   groups.set(key,info);
  }
  if(object.userData.universalSourcePartId===undefined)object.userData.universalSourcePartId=sourceId||null;
  object.userData.assetId=asset.id;object.userData.partId=info.id;
  seen.add(info.id);index++;
 });
 for(const [id] of active){
  if(seen.has(id))continue;
  active.delete(id);original.delete(id);removed++;
 }
 if(removed)record.rackParts=record.rackParts.filter(p=>!p.universal||seen.has(p.id));
 return {added,removed,total:active.size,meshes:index};
}
