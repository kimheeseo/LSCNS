import * as THREE from 'three';
import { createArchitectureKit } from './architecture.mjs?v=20261015-step04';
import { createEquipmentFactory } from './equipment-factory.mjs?v=20261015-step05';
import { createNpcCharacterFactory } from './npc-characters.mjs?v=20261015-step09b';
import { createNpcNavigation } from './npc-navigation.mjs?v=20261015-step07';
import { createInspectionFSM } from './inspection-fsm.mjs?v=20261015-step08c';
import { createVisitorSecurity } from './visitor-security.mjs?v=20261015-step09c';
import { createRackComponentKit } from './rack-components.mjs?v=20261015-step11to13';
import { createRackDoorControls } from './rack-door-controls.mjs?v=20261015-step11';
import { createRackServerRails } from './rack-server-rails.mjs?v=20261015-step12';
import { createRackOpticalCassette } from './rack-optical-cassette.mjs?v=20261015-step13';

// STEP 03: real Three.js scene graph, not 2D image planes.
// All world coordinates, equipment identity and simulation states are owned by the v4.7 host.
export function createSceneAdapter(T, canvas) {
  if (!T || !T.Z || !T.visibleAssets || !T.pickAsset) throw new Error('LS3D host API unavailable');
  const renderer = new THREE.WebGLRenderer({ canvas, antialias: true, alpha: false, powerPreference: 'high-performance' });
  renderer.setPixelRatio(1);
  renderer.outputColorSpace = THREE.SRGBColorSpace;
  renderer.toneMapping = THREE.ACESFilmicToneMapping;
  renderer.toneMappingExposure = 1.28;
  renderer.shadowMap.enabled = true;
  renderer.shadowMap.type = THREE.PCFSoftShadowMap;
  const scene = new THREE.Scene();
  scene.background = new THREE.Color('#183247');
  scene.fog = new THREE.Fog('#183247', 115, 310);
  const camera = new THREE.PerspectiveCamera(45, 1, 0.08, 1200);
  camera.up.set(0, 1, 0);
  const world = new THREE.Group(); scene.add(world);
  const zoneRoot = new THREE.Group(), assetRoot = new THREE.Group(), routeRoot = new THREE.Group(), peopleRoot = new THREE.Group();
  world.add(zoneRoot, assetRoot, routeRoot, peopleRoot);
  const hemi = new THREE.HemisphereLight(0xddf3ff, 0x243448, 3.4); scene.add(hemi);
  const sun = new THREE.DirectionalLight(0xfff0d6, 3.1);
  sun.position.set(-75, 130, 64);
  sun.castShadow = true;
  sun.shadow.mapSize.set(1024, 1024);
  sun.shadow.camera.left = -120; sun.shadow.camera.right = 120;
  sun.shadow.camera.top = 110; sun.shadow.camera.bottom = -110;
  sun.shadow.camera.near = 1; sun.shadow.camera.far = 340;
  sun.shadow.bias = -0.00025;
  scene.add(sun, sun.target);
  const cool = new THREE.DirectionalLight(0x4aa3ed, 1.3); cool.position.set(55, 35, -45); scene.add(cool);
  const boxGeometry = new THREE.BoxGeometry(1, 1, 1);
  const cylinderGeometry = new THREE.CylinderGeometry(1, 1, 1, 18);
  const materialCache = new Map();
  const registry = new Map(), people = new Map();
  const raycaster = new THREE.Raycaster(), mouse = new THREE.Vector2();
  let zoneSignature = '', routeSignature = '', disposed = false, frames = 0, failReason = null;
  const diagnostic = { engine: 'three', version: THREE.REVISION, frames: 0, assets: 0, zones: 0, npc: 0, calls: 0, triangles: 0, memoryGeometries: 0, memoryTextures: 0 };
  function material(color, metal = 0.28, rough = 0.48, transparent = false) {
    const key = String(color) + '|' + metal + '|' + rough + '|' + transparent;
    if (!materialCache.has(key)) materialCache.set(key, new THREE.MeshStandardMaterial({
      color, metalness: metal, roughness: rough, transparent, opacity: transparent ? 0.38 : 1,
      side: transparent ? THREE.DoubleSide : THREE.FrontSide, depthWrite: !transparent
    }));
    return materialCache.get(key);
  }
  function cube(parent, x, baseY, z, w, h, d, color, metal = 0.3, rough = 0.5, transparent = false) {
    const m = new THREE.Mesh(boxGeometry, material(color, metal, rough, transparent));
    m.position.set(x, baseY + h * .5, z); m.scale.set(Math.max(.001, w), Math.max(.001, h), Math.max(.001, d));
    m.castShadow = h > .3; m.receiveShadow = true; parent.add(m); return m;
  }
  function cylinder(parent, x, baseY, z, r, height, color) {
    const m = new THREE.Mesh(cylinderGeometry, material(color, .56, .36));
    m.position.set(x, baseY + height / 2, z); m.scale.set(Math.max(.001, r), Math.max(.001, height), Math.max(.001, r));
    m.castShadow = true; m.receiveShadow = true; parent.add(m); return m;
  }
  function clearGroup(g) { while (g.children.length) { const n = g.children[0]; g.remove(n); n.traverse(obj => {
    if (obj.geometry && obj.geometry !== boxGeometry && obj.geometry !== cylinderGeometry) obj.geometry.dispose();
    if (obj.material && !Array.from(materialCache.values()).includes(obj.material)) {
      if (Array.isArray(obj.material)) obj.material.forEach(m => m.dispose());
      else obj.material.dispose();
    }
  }); } }
  // STEP 04: real structural building envelope. Dynamic architecture shares
  // the existing source-of-truth zone data and never creates a second RAF loop.
  const architecture = createArchitectureKit(THREE, {cube});
  function zoneScene(flags, dt) {
    const mode=T.facilityMode, current=T.currentZone;
    const available=T.visibleZones();
    const sig=[mode,current,!!flags.showRoof,!!flags.showSection,
      ...available.map(z=>[z.id,z.pos.join(','),z.size.join(',')].join(':'))].join('|');
    zoneRoot.visible=!flags.inspect;
    if(sig!==zoneSignature) {
      zoneSignature=sig;
      clearGroup(zoneRoot);
      const result=architecture.build(zoneRoot,available,mode,current,flags);
      diagnostic.zones=result.zones;
    }
    if(!flags.inspect)architecture.tick(dt,current,flags);
    diagnostic.architecture=architecture.summary;
  }
  const equipmentFactory=createEquipmentFactory(THREE,{cube,cylinder,material,boxGeometry});
  function disposeAsset(item) {
    if(!item)return;
    assetRoot.remove(item.g);
    const id=item.g.userData.assetId;
    rackDoors.drop(id);rackRails.drop(id);rackOptics.drop(id);
    item.g.traverse(obj=>{
      if(obj.geometry && obj.geometry!==boxGeometry && obj.geometry!==cylinderGeometry) obj.geometry.dispose();
      // LOD levels use cached materials; selected-ring material is per-object.
      if(obj.material && !Array.from(materialCache.values()).includes(obj.material)) {
        if(Array.isArray(obj.material))obj.material.forEach(m=>m.dispose());
        else obj.material.dispose();
      }
    });
  }
  // STEP 10: replace the undifferentiated rack visual with separately tagged parts.
  // The original front-door/server-slide/fiber-cassette motion contract is preserved.
  const rackKit=createRackComponentKit(THREE,{cube});
  const rackDoors=createRackDoorControls();
  const rackRails=createRackServerRails(THREE,{cube});
  const rackOptics=createRackOpticalCassette(THREE,{cube});
  let selectedRackPart=null, rackHud=null, rackControls=null, rackControlsLabel=null;
  let focusedRack=null;
  function currentRackAsset() {
    const a=T.selected;
    return a&&T.isMechanicalRack(a)&&registry.has(a.id)?a:null;
  }
  function focusRackSide(assetId,side='front') {
    const a=T.assetMap.get(assetId),desired=T.cameraDesired;
    if(!a||!T.isMechanicalRack(a)||!desired||T.workerViewId||T.walkActive)return false;
    focusedRack={id:a.id,side:side==='rear'?'rear':'front'};
    desired.target=[a.x,a.h*.56,a.z];
    desired.distance=Math.max(7,Math.min(19,Math.max(a.w,a.d)*2.7));
    desired.pitch=.24;
    desired.yaw=focusedRack.side==='rear'?0:Math.PI;
    return true;
  }
  function setRackServerSelection(assetId,partId) {
    const selected=rackRails.select(assetId,partId);
    if(selected)selectRackPart(assetId,selected);
    return selected;
  }
  function cycleRackServer(assetId,step=1) {
    const s=registry.get(assetId)?.serverSession;
    if(!s)return null;
    const ids=[...s.groups.keys()],index=ids.indexOf(s.selected);
    if(!ids.length)return null;
    return setRackServerSelection(assetId,ids[(index+step+ids.length)%ids.length]);
  }
  function toggleRearRackDoor(assetId){return rackDoors.toggleRear(assetId);}
  function toggleRackServerSlide(assetId) {
    const a=T.assetMap.get(assetId);
    return !!(a&&T.isMechanicalRack(a)&&T.setMechanical?.(a,'slide'));
  }
  function toggleRackCassette(assetId) {
    const a=T.assetMap.get(assetId);
    return !!(a&&T.isMechanicalRack(a)&&T.setMechanical?.(a,'cassette'));
  }
  function toggleRackFrontDoor(assetId) {
    const a=T.assetMap.get(assetId);
    return !!(a&&T.isMechanicalRack(a)&&T.setMechanical?.(a,'door'));
  }
  function showControls() {
    if(typeof document==='undefined'||!canvas.parentElement)return;
    const a=currentRackAsset();
    if(!a){if(rackControls)rackControls.hidden=true;return;}
    if(!rackControls){
      rackControls=document.createElement('div');
      rackControls.id='ls3dRackMotionControls';
      rackControls.style.cssText='position:absolute;right:10px;bottom:10px;z-index:30;width:min(302px,calc(100% - 20px));padding:10px;background:rgba(13,29,45,.94);border:1px solid #6aa8bc;border-radius:10px;color:#f0fbff;font:12px system-ui;box-shadow:0 4px 20px #0008;pointer-events:auto;';
      rackControlsLabel=document.createElement('div');
      rackControlsLabel.id='ls3dRackActionStatus';
      rackControlsLabel.setAttribute('aria-live','polite');
      rackControlsLabel.style.cssText='font-weight:600;margin-bottom:7px;line-height:1.4;';
      rackControls.appendChild(rackControlsLabel);
      const grid=document.createElement('div');
      grid.style.cssText='display:grid;grid-template-columns:1fr 1fr;gap:5px;';
      const actions=[
        ['front-camera','전면 카메라'],['rear-camera','후면 카메라'],
        ['front-door','전면 문 열기/닫기'],['rear-door','후면 문 열기/닫기'],
        ['previous-server','이전 서버'],['next-server','다음 서버'],
        ['server-slide','선택 서버 인출/복귀'],['cassette-slide','광 카세트 인출/복귀'],
        ['next-fiber','다음 광 커넥터'],['clear-selection','부품 선택 해제']
      ];
      for(const [code,label] of actions){
        const button=document.createElement('button');
        button.type='button';button.dataset.rackAction=code;button.textContent=label;
        button.style.cssText='padding:7px 5px;border-radius:6px;border:1px solid #508497;background:#24465a;color:#eefaff;cursor:pointer;text-align:center;font:inherit;';
        grid.appendChild(button);
      }
      grid.addEventListener('click',e=>{
        const action=e.target.closest?.('button[data-rack-action]')?.dataset.rackAction;
        const asset=currentRackAsset();if(!asset||!action)return;
        e.preventDefault();e.stopPropagation();
        if(action==='front-camera')focusRackSide(asset.id,'front');
        else if(action==='rear-camera')focusRackSide(asset.id,'rear');
        else if(action==='front-door')toggleRackFrontDoor(asset.id);
        else if(action==='rear-door')toggleRearRackDoor(asset.id);
        else if(action==='previous-server')cycleRackServer(asset.id,-1);
        else if(action==='next-server')cycleRackServer(asset.id,1);
        else if(action==='server-slide')toggleRackServerSlide(asset.id);
        else if(action==='cassette-slide')toggleRackCassette(asset.id);
        else if(action==='next-fiber'){
          const item=registry.get(asset.id);
          const ports=(item?.opticsSession?.portDefs||[]);
          const current=ports.findIndex(p=>p.id===selectedRackPart?.id);
          if(ports.length)selectRackPart(asset.id,ports[(current+1)%ports.length].id);
        }else if(action==='clear-selection')showRackPart(null);
        showControls();
      });
      rackControls.appendChild(grid);
      canvas.parentElement.appendChild(rackControls);
    }
    rackControls.hidden=false;
    const server=rackRails.state(a.id),door=rackDoors.state(a.id),fiber=rackOptics.state(a.id);
    rackControlsLabel.textContent='랙 '+a.id+' · '+(focusedRack?.id===a.id?(focusedRack.side==='rear'?'후면 시점':'전면 시점'):'기본 시점')+
      ' · 서버 '+(server?.selected?.split(':').pop()||'—')+
      ' · 후면문 '+(door.targetRear?'열림':'닫힘')+
      ' · 광포트 '+(fiber?.ports.length||0)+'개';
  }
  function showRackPart(info) {
    selectedRackPart=info?{...info}:null;
    if(typeof document==='undefined'||!canvas.parentElement)return selectedRackPart;
    if(!rackHud) {
      rackHud=document.createElement('div');
      rackHud.id='ls3dRackPartInfo';
      rackHud.style.cssText='position:absolute;left:12px;bottom:12px;z-index:25;max-width:min(330px,85%);padding:11px 13px;border:1px solid #64aebc;border-radius:10px;background:rgba(15,35,54,.95);color:#f3fbff;font:12px system-ui;pointer-events:none;box-shadow:0 4px 20px #0007';
      canvas.parentElement.appendChild(rackHud);
    }
    rackHud.hidden=!selectedRackPart;
    if(selectedRackPart)rackHud.textContent=
      '선택 부품: '+selectedRackPart.label+' · '+selectedRackPart.assetId+
      ' · U '+selectedRackPart.uStart+'-'+(selectedRackPart.uStart+selectedRackPart.uHeight-1);
    return selectedRackPart;
  }
  function selectRackPart(assetId,partId) {
    const record=registry.get(assetId);
    const part=record?.partMap?.get(partId)||null;
    return showRackPart(part);
  }
  function assetMesh(a) {
    const g=new THREE.Group();
    g.name='equipment:'+a.id;g.userData.assetId=a.id;
    const {w,d,h}=a;
    let moving=null,rackLOD=null,equipmentLOD=null,kind='rack',rackParts=[],partMap=null,rackUnits=null;
    let serverSession=null,opticsSession=null;
    if(T.isMechanicalRack(a)){
      const high=new THREE.Group();high.name='rack-high';g.add(high);
      const built=rackKit.build(high,a);
      serverSession=rackRails.decorate(a,built,high);
      opticsSession=rackOptics.decorate(a,built);
      moving=built.moving;rackParts=built.parts;partMap=built.partMap;
      rackUnits=built.rackUnits;
      const low=new THREE.Group();low.name='rack-low';g.add(low);
      cube(low,0,.06,0,w*.92,h*.94,d*.90,'#53728b',.52,.44);
      cube(low,0,h*.72,-d*.48,w*.75,.26,.08,'#56a7b7',.32,.34);
      rackLOD={high,low};
    } else {
      const result=equipmentFactory.build(a);
      kind=result.kind;equipmentLOD=result.lod;g.add(equipmentLOD);
    }
    cube(g,0,-.09,0,w*1.08,.11,d*1.08,'#355269',.12,.88);
    g.position.set(a.x,0,a.z);
    const selection=new THREE.Mesh(
      new THREE.RingGeometry(Math.max(w,d)*.56,Math.max(w,d)*.66,36),
      new THREE.MeshBasicMaterial({color:0x63e4d8,side:THREE.DoubleSide,transparent:true,opacity:.94})
    );
    selection.rotation.x=-Math.PI/2;selection.position.y=.027;
    selection.name='selection-ring';selection.userData.noPick=true;g.add(selection);
    g.traverse(obj=>{if(obj.isMesh)obj.userData.assetId=a.id});
    return {g,moving,selection,kind,equipmentLOD,rackLOD,rackParts,partMap,rackUnits,serverSession,opticsSession};
  }
  function refreshAssets(flags,dt=.016) {
    const showDetail = !!flags.inspect;
    const visible = T.visibleAssets().filter(a => (T.currentZone === 'campus' || a.zone === T.currentZone) && (!showDetail || a.id === T.selected?.id));
    const present = new Set();
    for (const a of visible) {
      present.add(a.id);
      const sig = [a.zone,a.type,a.x,a.z,a.w,a.d,a.h].join('|');
      let item = registry.get(a.id);
      if (!item || item.sig !== sig) {
        if(item)disposeAsset(item);
        const model = assetMesh(a);
        item = { ...model, sig }; registry.set(a.id, item); assetRoot.add(item.g);
      }
      item.selection.visible=T.selected?.id===a.id;
      if(selectedRackPart && selectedRackPart.assetId===a.id && !item.partMap?.has(selectedRackPart.id))
        showRackPart(null);
      if(item.rackLOD) {
        const dx=camera.position.x-a.x, dy=camera.position.y-a.h*.5, dz=camera.position.z-a.z;
        const far=T.selected?.id!==a.id && Math.hypot(dx,dy,dz)>96;
        item.rackLOD.high.visible=!far;item.rackLOD.low.visible=far;
      }
      if(item.equipmentLOD)item.equipmentLOD.update(camera);
      const fault = !!a.down || !!(flags.scenario !== 'normal' && (
        (flags.scenario === 'cooling' && /tower/i.test(a.id)) ||
        (flags.scenario === 'network' && a.id === 'core') ||
        (flags.scenario === 'power' && a.id === 'switchgear')
      ));
      if (fault) item.g.scale.y = 1 + Math.sin((T.simTime || 0) * 5) * .008;
      else item.g.scale.y = 1;
      if (item.moving) {
        const m=T.mechanical(a);
        rackDoors.tick(a.id,item.moving,m,dt);
        rackRails.update(a.id,item.moving,m);
        rackOptics.update(a.id,item.moving,m);
        // When using the STEP 11 independent side camera, keep world-space rack
        // orientation fixed. Original legacy rear flip remains otherwise available.
        item.g.rotation.y=focusedRack?.id===a.id?0:(m.rear?Math.PI:0);
      }
    }
    for(const [id,item] of registry)if(!present.has(id)){disposeAsset(item);registry.delete(id);}
    if(selectedRackPart&&!registry.has(selectedRackPart.assetId))showRackPart(null);
    if(focusedRack&&!registry.has(focusedRack.id))focusedRack=null;
    showControls();
    diagnostic.assets=registry.size;
    diagnostic.rackInteraction={
      focus:focusedRack?{...focusedRack}:null,
      doors:{...rackDoors.stats},
      rails:{...rackRails.stats},
      optics:{...rackOptics.stats},
      selectedServer:currentRackAsset()?rackRails.state(currentRackAsset().id):null,
      selectedOptics:currentRackAsset()?rackOptics.state(currentRackAsset().id):null
    };
    diagnostic.rackComponents={
      rackCount:[...registry.values()].filter(x=>x.rackUnits).length,
      counts:{...rackKit.stats},
      selectedPart:selectedRackPart,
      visibleParts:[...registry.values()].reduce((n,x)=>n+(x.rackParts?.length||0),0)
    };
    diagnostic.equipment={
      kinds:{...equipmentFactory.diagnostics.kinds},
      instancedBatches:equipmentFactory.diagnostics.instancedBatches,
      instancedParts:equipmentFactory.diagnostics.instancedParts,
      activeKinds:[...new Set([...registry.values()].map(v=>v.kind))].sort(),
      lodAssets:[...registry.values()].filter(v=>!!v.equipmentLOD || !!v.rackLOD).length
    };
  }
  function routeScene(flags) {
    const on = flags.flow || {};
    // Server-room route filter is strict; avoid external world-space MEP paths.
    if (T.facilityMode === 'server-room' || flags.inspect) { if (routeSignature !== 'none') { clearGroup(routeRoot); routeSignature = 'none'; } return; }
    const sig = JSON.stringify([T.currentZone, on, T.flowPaths]);
    if (sig === routeSignature) return;
    routeSignature = sig; clearGroup(routeRoot);
    const keys = ['power','cooling','fiber'];
    const colors = { power:'#f8ce75', cooling:'#59c8ed', fiber:'#6debc2' };
    for (const key of keys) if (on[key] && Array.isArray(T.flowPaths[key])) {
      for (const path of T.flowPaths[key]) {
        if (!Array.isArray(path) || path.length < 2) continue;
        const points = path.filter(p => Array.isArray(p) && p.length >= 2).map(p => new THREE.Vector3(p[0], key === 'power' ? .22 : key === 'cooling' ? .45 : 6.38, p[1]));
        if (points.length < 2) continue;
        const geometry = new THREE.BufferGeometry().setFromPoints(points);
        const line = new THREE.Line(geometry, new THREE.LineBasicMaterial({ color: colors[key], transparent:true, opacity:.77 }));
        line.userData.noPick = true; routeRoot.add(line);
      }
    }
  }
  // STEP 06 replaces simple block people with reusable articulated Three.Group rigs.
  // Position/heading/action are still sampled from original T.npcPose; pathfinding is STEP 07.
  const npcFactory=createNpcCharacterFactory(THREE,{cube,cylinder});
  const navigation=createNpcNavigation(T,{cellSize:1.25,agentRadius:.39});
  const inspection=createInspectionFSM(T);
  const security=createVisitorSecurity(THREE,T,{cube,characterFactory:npcFactory});
  world.add(security.root);
  let securityState={active:false};
  function refreshPeople(flags) {
    const all=T.npcs || [],visible=new Set(),states={walk:0,inspect:0,pause:0,idle:0};
    const roles={},staff=T.facilityProfiles?.[T.facilityMode]?.staff || [];
    for(const n of all){
      if(T.facilityMode!=='aidc' && n.role==='guard')continue;
      if(n.role!=='guard' && !staff.includes(n.id))continue;
      if(T.currentZone!=='campus' && n.zone!==T.currentZone)continue;
      if(flags.inspect)continue;
      const rawPose=T.npcPose(n);
      if(!rawPose||!Number.isFinite(rawPose.x)||!Number.isFinite(rawPose.z))continue;
      const checked=inspection.update(n,rawPose,T.simTime||0,flags.scenario||'normal');
      const extra=n.id==='guard-a'?securityState.guardAAction:
        n.id==='guard-b'?securityState.guardBAction:null;
      const pose=extra?{...checked,action:extra}:checked;
      visible.add(n.id);
      let g=people.get(n.id);
      if(!g){g=npcFactory.create(n);people.set(n.id,g);peopleRoot.add(g);}
      g.position.set(pose.x,.04,pose.z);
      g.rotation.y=Number.isFinite(pose.heading)?pose.heading:0;
      const action=npcFactory.update(g,n,pose,T.simTime||0);
      states[action]++;
      const role=g.userData.npcRole;
      roles[role]=(roles[role]||0)+1;
    }
    for(const [id,g] of people)if(!visible.has(id)){
      peopleRoot.remove(g);
      people.delete(id); // all primitive geometries/materials are cached/shared by the adapter
    }
    diagnostic.npc=people.size;
    diagnostic.navigation=navigation.diagnostic();
    diagnostic.inspections=inspection.diagnostic();
    diagnostic.npcAnimation={
      visible:people.size,
      ids:[...people.keys()],
      roles,
      actions:states,
      articulated:people.size,
      totalBuilt:npcFactory.metrics.built,
      totalAnimationUpdates:npcFactory.metrics.animated
    };
  }
  let viewportWidth=0, viewportHeight=0, viewportDpr=0;
  function render(dt, view) {
    if (disposed) throw new Error('Three.js adapter was disposed');
    const rect = canvas.getBoundingClientRect();
    if (rect.width < 1 || rect.height < 1) return;
    const quality = window.LS3D_CONFIG?.quality || 'medium';
    const dpr = Math.min(window.devicePixelRatio || 1, quality === 'low' ? .8 : quality === 'high' ? 1.55 : 1.1);
    const pixelWidth=Math.floor(rect.width),pixelHeight=Math.floor(rect.height);
    if(viewportWidth!==pixelWidth || viewportHeight!==pixelHeight || viewportDpr!==dpr) {
      renderer.setPixelRatio(dpr);
      renderer.setSize(pixelWidth, pixelHeight, false);
      viewportWidth=pixelWidth;viewportHeight=pixelHeight;viewportDpr=dpr;
    }
    const flags = view.flags || {};
    const eye = view.eye, direction = view.forward, up = view.up;
    camera.position.set(eye[0],eye[1],eye[2]);
    camera.up.set(up[0],up[1],up[2]);
    camera.lookAt(eye[0] + direction[0] * 20, eye[1] + direction[1] * 20, eye[2] + direction[2] * 20);
    camera.fov = view.fov * 180 / Math.PI;
    camera.aspect = rect.width / rect.height;
    camera.near = Math.max(.035,view.near || .08);
    camera.far = Math.max(600, view.far || 600);
    camera.updateProjectionMatrix();
    zoneScene(flags, dt);
    refreshAssets(flags,dt);
    routeScene(flags);
    securityState=security.simulate(T.simTime||0);
    security.root.visible=securityState.active&&!flags.inspect;
    refreshPeople(flags);
    diagnostic.security=security.diagnostic();
    renderer.render(scene,camera);
    diagnostic.frames = ++frames;
    diagnostic.calls = renderer.info.render.calls;
    diagnostic.triangles = renderer.info.render.triangles;
    diagnostic.memoryGeometries = renderer.info.memory.geometries;
    diagnostic.memoryTextures = renderer.info.memory.textures;
  }
  function pick(clientX,clientY) {
    const rect=canvas.getBoundingClientRect();
    if (!rect.width || !rect.height) return null;
    mouse.x=((clientX-rect.left)/rect.width)*2-1;
    mouse.y=-((clientY-rect.top)/rect.height)*2+1;
    raycaster.setFromCamera(mouse,camera);
    for(const hit of raycaster.intersectObjects(assetRoot.children,true)) {
      if (hit.object.userData.noPick || hit.object.name === 'selection-ring') continue;
      let obj=hit.object,part=null;
      for(let p=obj;p&&p!==assetRoot;p=p.parent) {
        if(!part&&p.userData.rackPart)part=p.userData.rackPart;
      }
      while(obj && !obj.userData.assetId) obj=obj.parent;
      const id=obj?.userData.assetId, asset=T.assetMap.get(id);
      if(asset){
        if(part) {
          selectRackPart(id,part.id);
          if(part.kind?.startsWith('server-')||part.kind?.startsWith('network-switch-')||part.kind?.startsWith('storage-'))
            rackRails.select(id,part.id);
        } else showRackPart(null);
        T.pickAsset(asset);return id;
      }
    }
    showRackPart(null);
    return null;
  }
  function dispose() {
    if(disposed)return;disposed=true;
    navigation.reset();inspection.reset();
    if(rackHud){rackHud.remove();rackHud=null;}selectedRackPart=null;
    if(rackControls){rackControls.remove();rackControls=null;}
    rackDoors.reset();rackRails.reset();rackOptics.reset();focusedRack=null;
    security.root.removeFromParent?.();
    clearGroup(security.root);
    clearGroup(routeRoot);clearGroup(zoneRoot);clearGroup(peopleRoot);
    for(const item of registry.values())disposeAsset(item);
    registry.clear();
    for(const m of materialCache.values())m.dispose();
    materialCache.clear();boxGeometry.dispose();cylinderGeometry.dispose();
    renderer.dispose();
  }
  return { render, pick, dispose, diagnostic, renderer, scene, camera, THREE, get architecture(){return architecture.summary;}, get equipment(){return diagnostic.equipment;}, get npcAnimation(){return diagnostic.npcAnimation;},
    get inspections(){return diagnostic.inspections;},get security(){return diagnostic.security;},
    get rackComponents(){return diagnostic.rackComponents;},get selectedPart(){return selectedRackPart;},
    selectRackPart,focusRackSide,setRackServerSelection,cycleRackServer,
    toggleRearRackDoor,toggleRackFrontDoor,toggleRackServerSlide,toggleRackCassette,
    get rackInteraction(){return diagnostic.rackInteraction;},
    inspection,securityController:security,navigation,
    get disposed(){return disposed;} };
}
