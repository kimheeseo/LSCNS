import {isVisiblePickObject,describeEquipment} from './scene-inspection.mjs?v=20261024-product-click';
import { applyRackCamera, rackFocusActive } from './rack-camera-policy.mjs?v=20261101-rack-bottom-space-v1';
import * as THREE from 'three';
import { createArchitectureKit } from './architecture.mjs?v=20261029-tour-gate-v1';
import { createEquipmentFactory } from './equipment-factory.mjs?v=20261031-loading-crew-v1';
import { createNpcCharacterFactory } from './npc-characters.mjs?v=20261025-visitor-motion';
import { createVisitorSecurity } from './visitor-security.mjs?v=20261025-visitor-checkin';
import { createNpcNavigation } from './npc-navigation.mjs?v=20261015-step07';
import { createInspectionFSM } from './inspection-fsm.mjs?v=20261015-step08c';
import { createRackComponentKit } from './rack-components.mjs?v=20261102-rack-part-name-role-v1';
import { createRackDoorControls } from './rack-door-controls.mjs?v=20261015-step11';
import { createRackServerRails } from './rack-server-rails.mjs?v=20261027-rack-roles-v1';
import { createRackOpticalCassette } from './rack-optical-cassette.mjs?v=20261102-rack-part-name-role-v1';
import { createRackNetworkCabling } from './rack-network-cabling.mjs?v=20261026-leaf-spine-racks-v1';
import { createRackPatchCord } from './rack-patch-cord.mjs?v=20261031-cable-connect-v1';
import { createRackOpticalPatchDesign } from './rack-optical-patch-design.mjs?v=20261102-patch-component-guide-v1';
import { createRackLcAdapterDesign } from './rack-lc-adapter-design.mjs?v=20261101-lc-odf-design-v1';
import { createOdfSmallFiberBox } from './odf-small-fiber-box.mjs?v=20261101-odf-small-design-v4';
import { createZoneSpecializations } from './zone-specializations.mjs?v=20261026-black-mesh-cage-v2';
import { createCampusMEPRoutes } from './campus-mep-routes.mjs?v=20261015-step17';
import { createViewportPolicy, createCameraOcclusion, layoutSceneOverlays } from './scene-viewport-policy.mjs?v=20261015-rack-collapse';
import { createRackPanelDisclosure } from './rack-panel-disclosure.mjs?v=20261015-rack-collapse';

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
  let zoneSignature = '', disposed = false, frames = 0, failReason = null;
  const diagnostic = { engine: 'three', version: THREE.REVISION, frames: 0, assets: 0, zones: 0, npc: 0, calls: 0, triangles: 0, memoryGeometries: 0, memoryTextures: 0 };
  const viewportPolicy=createViewportPolicy();
  const cameraOcclusion=createCameraOcclusion(THREE,{zoneRoot});
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
  function clearGroup(g) {
    // Snapshot cached materials just once per teardown, not once per mesh.
    const cached=new Set(materialCache.values());
    while(g.children.length){
      const child=g.children[0];g.remove(child);
      child.traverse(obj=>{
        if(obj.geometry&&obj.geometry!==boxGeometry&&obj.geometry!==cylinderGeometry)
          obj.geometry.dispose();
        if(obj.material){
          const materials=Array.isArray(obj.material)?obj.material:[obj.material];
          for(const mat of materials)if(!cached.has(mat))mat.dispose?.();
        }
      });
    }
  }
  // STEP 04: real structural building envelope. Dynamic architecture shares
  // the existing source-of-truth zone data and never creates a second RAF loop.
  const architecture = createArchitectureKit(THREE, {cube});
  function zoneScene(flags, dt) {
    const mode=T.facilityMode, current=T.currentZone;
    const available=T.visibleZones();
    const sig=[mode,current,!!flags.showRoof,!!flags.showSection,
      ...available.map(z=>[z.id,z.pos.join(','),z.size.join(',')].join(':'))].join('|');
    zoneRoot.visible=!flags.inspect&&!flags.coloTour;
    if(sig!==zoneSignature) {
      zoneSignature=sig;
      cameraOcclusion.restore();
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
    rackDoors.drop(id);rackRails.drop(id);rackOptics.drop(id);rackNetwork.drop(id);rackPatch.drop(id);rackPatchDesign.drop(id);rackLcDesign.drop(id);odfBox.drop(id);
    item.g.traverse(obj=>{
      if(obj.isInstancedMesh)obj.dispose?.(); // Release per-instance GPU buffers on rack removal.
      if(obj.userData?.ownedTexture)obj.userData.ownedTexture.dispose();
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
  const rackNetwork=createRackNetworkCabling(THREE,{cube});
  const rackPatch=createRackPatchCord(THREE);
  const rackPatchDesign=createRackOpticalPatchDesign(THREE,{cube});
  const rackLcDesign=createRackLcAdapterDesign(THREE,{cube});
  const odfBox=createOdfSmallFiberBox(THREE,{cube});
  const zoneExtras=createZoneSpecializations(THREE,T,{cube,world,canvas});
  const mepRoutes=createCampusMEPRoutes(THREE,T,{cube,root:routeRoot,canvas});
  let selectedRackPart=null, rackHud=null, rackControls=null, rackControlsLabel=null, rackPanelDisclosure=null, rackPanelHeading=null, rackControlsActiveRack=null;
  let opticalScaleAutoCollapsed=false;
  let focusedRack=null;
  let inspectedScenePart=null,sceneInspectionContext='',scenePartMaterials=[];
  function clearSceneInspection(){
    for(const entry of scenePartMaterials){if(entry.object.material===entry.highlighted)entry.object.material=entry.original;entry.highlighted.dispose?.();}
    scenePartMaterials=[];inspectedScenePart=null;sceneInspectionContext='';
  }
  function clearInspectionSelection(){showRackPart(null);}
  function selectScenePart(info){
    clearRackPartEmphasis();clearSceneInspection();
    if(selectedRackPart){
      rackPatchDesign.isolate(selectedRackPart.assetId,false);
      rackPatchDesign.select(selectedRackPart.assetId,false);
      rackOptics.isolate(selectedRackPart.assetId,false);
      rackLcDesign.isolate(selectedRackPart.assetId,false);
      rackLcDesign.highlight(selectedRackPart.assetId,-1);
      odfBox.isolate(selectedRackPart.assetId,false);
      odfBox.highlight(selectedRackPart.assetId,-1);
    }
    selectedRackPart=null;if(rackHud)rackHud.hidden=true;
    inspectedScenePart={...info};sceneInspectionContext=T.currentZone+'|'+T.facilityMode+'|'+T.selected?.id;
    const highlightRoot=info.assetId?registry.get(info.assetId)?.g:zoneRoot;
    highlightRoot?.traverse(object=>{
      if(!object.isMesh||object.userData?.inspectionPickSurface||(!info.assetId&&object.userData?.inspectionPart?.id!==info.id)||object.userData?.noPick||!object.material?.clone)return;
      const original=object.material,highlighted=original.clone();
      highlighted.emissive?.set(0xffd65a);highlighted.emissiveIntensity=1.2;
      object.material=highlighted;scenePartMaterials.push({object,original,highlighted});
    });
    window.LS3D_INSPECTOR?.setSelection(info);return info;
  }
  function selectInspectionPart(id){
    if(id.startsWith('equipment:')){const assetId=id.slice(10),record=registry.get(assetId),asset=T.assetMap.get(assetId);if(record&&asset){T.pickAsset(asset,false);return selectScenePart(describeEquipment(asset,record.kind));}}
    for(const [assetId,record] of registry){if(record.partMap?.has(id)){const part=selectRackPart(assetId,id);const asset=T.assetMap.get(assetId);if(asset)T.pickAsset(asset,false);return part;}}
    let info=null;zoneRoot.traverse(o=>{if(!info&&o.userData?.inspectionPart?.id===id)info=o.userData.inspectionPart;});
    return info?selectScenePart(info):null;
  }
  function syncInspectionCatalog(){
    const selectedRecord=registry.get(T.selected?.id);
    const equipment=selectedRecord&&describeEquipment(T.selected,selectedRecord.kind);
    const parts=[...(selectedRecord?.rackParts||[]),...(equipment?[equipment]:[])],ids=new Set(parts.map(p=>p.id));
    if(zoneRoot.visible)zoneRoot.traverse(o=>{const p=o.userData?.inspectionPart;if(p&&isVisiblePickObject(o)&&!ids.has(p.id)){ids.add(p.id);parts.push(p);}});
    window.LS3D_INSPECTOR?.setParts(parts);window.LS3D_INSPECTOR?.sync();
  }
  function currentRackAsset() {
    const a=T.selected;
    return a&&T.isMechanicalRack(a)&&registry.has(a.id)?a:null;
  }
  function focusRackSide(assetId,side='front') {
    const a=T.assetMap.get(assetId),desired=T.cameraDesired;
    if(!a||!T.isMechanicalRack(a)||!desired||T.workerViewId||T.walkActive)return false;
    const rect=canvas.getBoundingClientRect();
    focusedRack=applyRackCamera(T,a,{width:rect.width,height:rect.height,side});
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
  function cycleRackGpu(assetId,step=1){
    const a=T.assetMap.get(assetId),gpu=a&&/gpu/i.test(a.type)?rackRails.cycleGpu(assetId,step):null;
    if(!gpu)return null;
    const motion=T.mechanical(a);
    if(!motion.targetSlide)T.setMechanical(a,'slide'); // opens the door as well
    selectRackPart(assetId,gpu.id);
    return gpu;
  }
  function toggleRackGpu(assetId){
    const a=T.assetMap.get(assetId);
    if(!a||!/gpu/i.test(a.type))return null;
    const motion=T.mechanical(a);
    if(!motion.targetSlide)T.setMechanical(a,'slide');
    const gpu=rackRails.toggleGpu(assetId);
    if(gpu)selectRackPart(assetId,gpu.id);
    return gpu;
  }
  function cycleRackDrive(assetId){
    const drive=rackRails.cycleDrive(assetId,1);
    if(drive)selectRackPart(assetId,drive.id);
    return drive;
  }
  function toggleRackDrive(assetId){
    const a=T.assetMap.get(assetId);
    if(!a||a.type!=='Storage rack')return null;
    const motion=T.mechanical(a);
    if(!motion.targetDoor)T.setMechanical(a,'door');
    const drive=rackRails.toggleDrive(assetId);
    if(drive)selectRackPart(assetId,drive.id);
    return drive;
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
  function focusRackOpticalAssembly(assetId,action){
    const asset=T.assetMap.get(assetId),item=registry.get(assetId);
    if(!asset||!item)return false;
    const kindByAction={'optical-patch':'patch-panel','optical-mtp':'optical-module',
      'optical-lc':'lc-adapter-panel','optical-transceiver':'optical-transceiver-'};
    const kind=kindByAction[action];
    const part=item.rackParts.find(p=>p.kind===kind||((action==='optical-transceiver')&&p.kind.startsWith('optical-transceiver-')));
    if(!part)return false;
    focusRackSide(assetId,'front');
    const motion=T.mechanical(asset);
    if(!motion.targetDoor)T.setMechanical(asset,'door');
    if(action==='optical-mtp'&&!motion.targetCassette)T.setMechanical(asset,'cassette');
    selectRackPart(assetId,part.id);
    if(action==='optical-patch'||action==='optical-mtp'||action==='optical-lc')return true; // Each assembly has its own inspection view.
    const d=T.cameraDesired;
    if(d){
      const unit=(asset.h-.46)/(item.rackUnits||42);
      const yy=.23+(part.uStart-1+part.uHeight*.50)*unit;
      d.target=[asset.x,yy,asset.z-asset.d*.55-(action==='optical-mtp'?.40:0)];
      d.distance=Math.max(3.15,asset.w*1.30);d.yaw=Math.PI;d.pitch=.10;
    }
    return true;
  }
  function focusOdfSmallBox(assetId){
    const asset=T.assetMap.get(assetId),item=registry.get(assetId);
    if(asset?.id!=='odf'||!item?.odfSmallSession)return false;
    // The network zone turns this rack by PI; its enclosure front faces world +Z.
    focusRackSide(assetId,'rear');
    if(!T.mechanical(asset).targetDoor)T.setMechanical(asset,'door');
    selectRackPart(assetId,assetId+':odf-small:housing');
    return true;
  }
  function focusRackStoragePart(assetId,action){
    const asset=T.assetMap.get(assetId),item=registry.get(assetId);
    if(!asset||!item||asset.type!=='Storage rack')return false;
    focusRackSide(assetId,'front');
    const motion=T.mechanical(asset);
    if(!motion.targetDoor)T.setMechanical(asset,'door');
    let part=null;
    if(action==='storage-controller')part=item.rackParts.find(p=>p.kind==='storage-controller');
    else if(action==='storage-psu')part=item.rackParts.find(p=>p.kind==='storage-psu-fan');
    else if(action==='storage-shelf')part=item.rackParts.find(p=>p.kind==='storage-1');
    if(part){if(action==='storage-shelf')rackRails.select(assetId,part.id);selectRackPart(assetId,part.id);}
    const d=T.cameraDesired;
    if(d&&part){const u=(asset.h-.46)/(item.rackUnits||42);
      d.target=[asset.x,.23+(part.uStart-1+part.uHeight*.5)*u,asset.z-asset.d*.48];
      d.distance=Math.max(3.2,asset.w*1.35);d.yaw=Math.PI;d.pitch=.10;}
    return !!part;
  }
  function showControls() {
    if(typeof document==='undefined'||!canvas.parentElement)return;
    const a=currentRackAsset();
    // Rack-only motion controls are shown exclusively inside ‘랙 단독 보기’.
    if(!a||!T.rackIsolateActive){if(rackControls)rackControls.hidden=true;rackControlsActiveRack=null;return;}
    if(!rackControls){
      rackControls=document.createElement('div');
      rackControls.id='ls3dRackMotionControls';
      rackControls.style.cssText='position:absolute;right:10px;bottom:10px;z-index:30;width:min(325px,calc(100% - 20px));max-height:min(72dvh,650px);overflow-y:auto;padding:10px;background:rgba(13,29,45,.94);border:1px solid #6aa8bc;border-radius:10px;color:#f0fbff;font:12px system-ui;box-shadow:0 4px 20px #0008;pointer-events:auto;';
      const bar=document.createElement('div');
      bar.style.cssText='display:flex;align-items:center;justify-content:space-between;gap:8px;min-height:31px;';
      rackPanelHeading=document.createElement('strong');
      rackPanelHeading.id='ls3dRackPanelTitle';
      rackPanelHeading.textContent='랙 조작';
      rackPanelHeading.style.cssText='font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;min-width:0;';
      bar.appendChild(rackPanelHeading);
      const disclosureButton=document.createElement('button');
      disclosureButton.id='ls3dRackPanelToggle';
      disclosureButton.type='button';
      disclosureButton.style.cssText='flex:none;padding:6px 10px;border:1px solid #74b4c6;border-radius:7px;background:#22516a;color:#effcff;font:600 12px system-ui;cursor:pointer;touch-action:manipulation;';
      bar.appendChild(disclosureButton);rackControls.appendChild(bar);
      const controlsBody=document.createElement('div');
      controlsBody.id='ls3dRackMotionBody';
      controlsBody.style.cssText='padding-top:5px;';
      rackControls.appendChild(controlsBody);
      rackControlsLabel=document.createElement('div');
      rackControlsLabel.id='ls3dRackActionStatus';
      rackControlsLabel.setAttribute('aria-live','polite');
      rackControlsLabel.style.cssText='font-weight:600;margin-bottom:7px;line-height:1.4;';
      controlsBody.appendChild(rackControlsLabel);
      const grid=document.createElement('div');
      grid.style.cssText='display:grid;grid-template-columns:1fr 1fr;gap:5px;';
      const actions=[
        ['front-camera','전면 카메라'],['rear-camera','후면 카메라'],
        ['front-door','전면 문 열기/닫기'],['rear-door','후면 문 열기/닫기'],
        ['previous-server','이전 모듈'],['next-server','다음 모듈'],
        ['server-slide','선택 모듈 인출/복귀'],['next-gpu','다음 GPU 슬롯 선택'],['gpu-toggle','GPU 모듈 인출/장착'],
        ['next-drive','다음 드라이브 선택'],['drive-toggle','선택 드라이브 인출/장착'],['cassette-slide','광 카세트 인출/복귀'],
        ['patch-select','광 패치패널 선택'],['switch-select','네트워크 스위치 선택'],
        ['cassette-select','MPO/MTP 카세트 선택'],['gpu-select','GPU 모듈 선택'],
        ['cable-patch','케이블 · 패치패널 연결/분리'],
        ['optical-detail','LS 광패널 상세'],['optical-explode','모듈 분해 보기'],['next-fiber','광 커넥터 선택·확대'],['clear-selection','부품 선택 해제']
      ];
      for(const [code,label] of actions){
        const button=document.createElement('button');
        button.type='button';button.dataset.rackAction=code;button.textContent=label;
        if(code==='next-gpu'||code==='gpu-toggle'||code==='gpu-select')button.dataset.gpuOnly='1';
        if(code==='next-drive'||code==='drive-toggle')button.dataset.storageOnly='1';
        button.style.cssText='padding:7px 5px;border-radius:6px;border:1px solid #508497;background:#24465a;color:#eefaff;cursor:pointer;text-align:center;font:inherit;';
        grid.appendChild(button);
      }
      rackControls.addEventListener('click',e=>{
        const action=e.target.closest?.('button[data-rack-action]')?.dataset.rackAction;
        const asset=currentRackAsset();if(!asset||!action)return;
        e.preventDefault();e.stopPropagation();
        if(action==='patch-select')focusRackOpticalAssembly(asset.id,'optical-patch');
        else if(action==='switch-select'){
          const switchPart=registry.get(asset.id)?.rackParts.find(p=>p.kind==='tor');
          if(switchPart)selectRackPart(asset.id,switchPart.id);
        }
        else if(action==='cassette-select')focusRackOpticalAssembly(asset.id,'optical-mtp');
        else if(action==='gpu-select')cycleRackGpu(asset.id,0);
        else if(action==='cable-patch'){
          if(focusRackOpticalAssembly(asset.id,'optical-patch')){
            rackPatch.toggle(asset.id);
            rackPatchDesign.includeCable(asset.id,true);
          }
        }
        else if(['optical-patch','optical-mtp','optical-lc','optical-transceiver'].includes(action))focusRackOpticalAssembly(asset.id,action);
        else if(action==='lc-next-port'){
          if(!rackLcDesign.owns(selectedRackPart))focusRackOpticalAssembly(asset.id,'optical-lc');
          const next=((rackLcDesign.state(asset.id)?.selectedPort??-1)+1)%12;
          selectRackPart(asset.id,asset.id+':lc-adapter-port-'+next);
        }else if(action==='lc-rear'){
          if(!rackLcDesign.owns(selectedRackPart))focusRackOpticalAssembly(asset.id,'optical-lc');
          const s=registry.get(asset.id)?.lcSession,desired=T.cameraDesired;
          if(s&&desired){
            focusRackSide(asset.id,'rear');
            const point=new THREE.Vector3(0,s.lower+s.height*.6,s.rear+.20);
            registry.get(asset.id).g.updateMatrixWorld(true);registry.get(asset.id).g.localToWorld(point);
            desired.target=[point.x,point.y,point.z];desired.distance=Math.max(2.15,asset.w*1.05);
            desired.yaw=0;desired.pitch=.16;
            if(T.camera)Object.assign(T.camera,{target:desired.target.slice(),distance:desired.distance,yaw:desired.yaw,pitch:desired.pitch});
          }
        }
        else if(action==='odf-select')focusOdfSmallBox(asset.id);
        else if(action==='odf-cover'){
          if(!odfBox.owns(selectedRackPart))focusOdfSmallBox(asset.id);
          odfBox.toggle(asset.id);
        }else if(action==='odf-next-port'){
          if(!odfBox.owns(selectedRackPart))focusOdfSmallBox(asset.id);
          const next=((odfBox.state(asset.id)?.selectedPort??-1)+1)%6;
          selectRackPart(asset.id,asset.id+':odf-small:lc-port-'+next);
        }
        else if(['storage-controller','storage-psu','storage-shelf'].includes(action))focusRackStoragePart(asset.id,action);
        else if(action==='storage-shelf-pull'){
          const selected=rackRails.state(asset.id)?.selected;
          const part=selected&&registry.get(asset.id)?.partMap?.get(selected);
          if(part&&part.kind?.startsWith('storage-')){
            focusRackSide(asset.id,'front');
            if(!T.mechanical(asset).targetDoor)T.setMechanical(asset,'door');
            selectRackPart(asset.id,part.id);
          }else focusRackStoragePart(asset.id,'storage-shelf');
          toggleRackServerSlide(asset.id);
        }else if(action==='storage-drive-pull'){
          if(!rackRails.state(asset.id)?.selectedDrive)cycleRackDrive(asset.id);
          toggleRackDrive(asset.id);
        }else if(action==='storage-rear'){
          focusRackSide(asset.id,'rear');
          if(!rackDoors.state(asset.id)?.targetRear)toggleRearRackDoor(asset.id);
        }
        else if(action==='front-camera')focusRackSide(asset.id,'front');
        else if(action==='rear-camera')focusRackSide(asset.id,'rear');

        else if(action==='front-door')toggleRackFrontDoor(asset.id);
        else if(action==='rear-door')toggleRearRackDoor(asset.id);
        else if(action==='previous-server')cycleRackServer(asset.id,-1);
        else if(action==='next-server')cycleRackServer(asset.id,1);
        else if(action==='server-slide')toggleRackServerSlide(asset.id);
        else if(action==='next-gpu')cycleRackGpu(asset.id,1);
        else if(action==='gpu-toggle')toggleRackGpu(asset.id);
        else if(action==='next-drive')cycleRackDrive(asset.id);
        else if(action==='drive-toggle')toggleRackDrive(asset.id);
        else if(action==='cassette-slide')toggleRackCassette(asset.id);
        else if(action==='optical-detail'||action==='optical-explode'){
          focusRackSide(asset.id,'front');
          const motion=T.mechanical(asset);if(!motion.targetCassette)toggleRackCassette(asset.id);
          if(action==='optical-explode')rackOptics.explode(asset.id);
          const info=registry.get(asset.id)?.rackParts.find(p=>p.kind==='optical-module');
          if(info)selectRackPart(asset.id,info.id);
          const desired=T.cameraDesired;
          desired.target=[asset.x,asset.h*.89,asset.z-asset.d*.65];
          desired.distance=Math.max(3.4,asset.w*1.5);desired.pitch=.62;desired.yaw=Math.PI+.22;
        }
        else if(action==='link-all')rackNetwork.focus(asset.id,'all');
        else if(action==='link-fiber')rackNetwork.focus(asset.id,'fiber');
        else if(action==='link-dac')rackNetwork.focus(asset.id,'dac');
        else if(action==='link-power-a')rackNetwork.focus(asset.id,'power-A');
        else if(action==='link-power-b')rackNetwork.focus(asset.id,'power-B');
        else if(action==='next-fiber'){
          const item=registry.get(asset.id);
          const ports=(item?.opticsSession?.portDefs||[]);
          const current=ports.findIndex(p=>p.id===selectedRackPart?.id);
          if(ports.length)selectRackPart(asset.id,ports[(current+1)%ports.length].id);
        }else if(action==='clear-selection')showRackPart(null);
        showControls();
      });
      controlsBody.appendChild(grid);
      const createAccessoryButtons=(id,title,description,items)=>{
        const panel=document.createElement('details');panel.id=id;panel.open=true;
        panel.style.cssText='border-top:1px solid #477384;margin-top:9px;padding-top:8px;';
        const heading=document.createElement('summary');heading.textContent=title;
        heading.style.cssText='cursor:pointer;font-size:12px;font-weight:800;color:#bdeefa;';
        panel.appendChild(heading);
        const note=document.createElement('p');note.textContent=description;
        note.style.cssText='color:#bad4e2;font-size:11px;line-height:1.65;margin:6px 0;';
        panel.appendChild(note);
        const buttons=document.createElement('div');
        buttons.style.cssText='display:grid;grid-template-columns:1fr 1fr;gap:5px;';
        for(const [code,label] of items){
          const button=document.createElement('button');button.type='button';button.dataset.rackAction=code;button.textContent=label;
          button.style.cssText='padding:8px 5px;border-radius:6px;border:1px solid #4a91aa;background:#173d51;color:#edfaff;cursor:pointer;text-align:center;font:600 11px "Pretendard Variable",Pretendard,"Malgun Gothic",sans-serif;line-height:1.5;';
          buttons.appendChild(button);
        }
        panel.appendChild(buttons);controlsBody.appendChild(panel);
      };
      createAccessoryButtons('ls3dNetworkOpticalAccessories','◎ 네트워크 랙 광배선 장치',
        '문을 열어 패치패널·MPO/MTP 카세트·LC 어댑터·광모듈을 선택하면 오른쪽 설명에 구성과 역할이 나옵니다.',
        [['optical-patch','1 · 광 패치패널'],['optical-mtp','2 · MPO/MTP 카세트'],
         ['optical-lc','3 · LC 어댑터 패널'],['lc-rear','LC 패널 후면 보기'],
         ['lc-next-port','다음 LC 포트·피그테일'],['optical-transceiver','4 · 광 트랜시버']]);
      createAccessoryButtons('ls3dStorageRackAccessories','▣ 스토리지 랙 내부 확인',
        '8-bay 스토리지 쉘프, 컨트롤러, PSU/팬을 구분해 확인하고 쉘프/드라이브를 각각 인출해 보세요.',
        [['storage-shelf','스토리지 쉘프 선택'],['storage-shelf-pull','쉘프 꺼내기/넣기'],
         ['storage-controller','컨트롤러 확인'],['storage-psu','PSU / 팬 확인'],
         ['storage-drive-pull','선택 드라이브 인출'],['storage-rear','후면 연결 확인']]);
      createAccessoryButtons('ls3dODFSmallBoxAccessories','◎ ODF 소형 광단자함',
        '검은 힌지 커버를 열고 6색 피그테일 루프·접속 홀더와 측면 LC 어댑터를 확인합니다.',
        [['odf-select','소형 단자함 선택'],['odf-cover','커버 열기 / 닫기'],['odf-next-port','다음 LC 포트·광경로']]);
      const details=document.createElement('details');
      details.style.cssText='border-top:1px solid #477384;margin-top:7px;padding-top:7px;';
      const title=document.createElement('summary');
      title.textContent='◎ ToR / 후면 PDU A·B / 서버 NIC 배선';
      title.style.cssText='cursor:pointer;font-weight:650;padding:3px 0;';
      details.appendChild(title);
      const cablesGrid=document.createElement('div');
      cablesGrid.style.cssText='display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:5px;margin-top:5px;';
      for(const [code,label] of [
        ['link-all','모든 랙 내부 배선'],['link-fiber','광 점퍼 → ToR'],['link-dac','DAC → ToR'],
        ['link-power-a','전원 A / PDU-A'],['link-power-b','전원 B / PDU-B']
      ]){
        const button=document.createElement('button');
        button.type='button';button.dataset.rackAction=code;button.textContent=label;
        button.style.cssText='padding:7px 5px;border-radius:6px;border:1px solid #508497;background:#24465a;color:#eefaff;cursor:pointer;text-align:center;font:inherit;';
        cablesGrid.appendChild(button);
      }
      details.appendChild(cablesGrid);controlsBody.appendChild(details);
      rackPanelDisclosure=createRackPanelDisclosure({
        panel:rackControls,body:controlsBody,button:disclosureButton,
        title:rackPanelHeading,
        collapsed:canvas.getBoundingClientRect().width<=640
      });
      disclosureButton.addEventListener('click',e=>{
        e.preventDefault();e.stopPropagation();
        rackPanelDisclosure.toggle();
      });
      canvas.parentElement.appendChild(rackControls);
    }
    // Restore the old open/close workflow: entering rack-only view must reveal
    // all mechanical controls, even when the canvas width is phone-sized.
    // Preserve manual collapse within the same inspection session.
    if(rackControlsActiveRack!==a.id){
      rackPanelDisclosure?.setCollapsed(false);
      rackControlsActiveRack=a.id;
    }
    rackControls.hidden=false;
    const gpuRack=/gpu/i.test(a.type),storageRack=a.type==='Storage rack',networkRack=a.type==='Top-of-Rack switch'||['Core','Spine','Leaf','MMR rack'].includes(a.type);
    const networkPanel=rackControls.querySelector('#ls3dNetworkOpticalAccessories');
    const storagePanel=rackControls.querySelector('#ls3dStorageRackAccessories');
    const odfPanel=rackControls.querySelector('#ls3dODFSmallBoxAccessories');
    if(networkPanel)networkPanel.hidden=!networkRack;
    if(storagePanel)storagePanel.hidden=!storageRack;
    if(odfPanel)odfPanel.hidden=a.id!=='odf';
    for(const code of ['lc-rear','lc-next-port']){
      const button=rackControls.querySelector('[data-rack-action="'+code+'"]');
      if(button)button.hidden=!rackLcDesign.state(a.id);
    }
    rackControls.querySelectorAll('[data-gpu-only]').forEach(button=>button.hidden=!gpuRack);
    rackControls.querySelectorAll('[data-storage-only]').forEach(button=>button.hidden=!storageRack);
    if(rackPanelHeading)rackPanelHeading.textContent=(gpuRack?'GPU 서버':storageRack?'스토리지':networkRack?'네트워크':'서버')+' 랙 조작 · '+a.id;
    for(const code of ['previous-server','next-server','server-slide']){
      const button=rackControls.querySelector('[data-rack-action="'+code+'"]');if(!button)continue;
      const noun=networkRack?'스위치':storageRack?'스토리지 쉘프':'서버';
      button.textContent=code==='server-slide'?'선택 '+noun+' 인출/복귀':(code==='previous-server'?'이전 ':'다음 ')+noun;
    }
    const patchStatus=rackPatch.state(a.id);
    const patchAction=rackControls.querySelector('[data-rack-action="cable-patch"]');
    if(patchAction)patchAction.textContent=patchStatus?.connected?'케이블 분리 · 패치패널에서 빼기':'케이블 연결 · 패치패널에 꽂기';
    const patchSelect=rackControls.querySelector('[data-rack-action="patch-select"]');
    if(patchSelect)patchSelect.hidden=!patchStatus;
    if(patchAction)patchAction.hidden=!patchStatus;
    const odfCover=rackControls.querySelector('[data-rack-action="odf-cover"]');
    if(odfCover)odfCover.textContent=odfBox.state(a.id)?.target?'커버 닫기':'커버 열기';
    const server=rackRails.state(a.id),door=rackDoors.state(a.id),fiber=rackOptics.state(a.id);
    const serverLabel=registry.get(a.id)?.partMap?.get(server?.selected)?.label||'—';
    const gpuStatus=gpuRack&&server?.gpuCount?' · GPU '+server.gpuSlot+'/'+server.gpuCount+(server.gpuTarget?' 인출':' 장착'):'';
    const gpuButton=rackControls.querySelector('[data-rack-action="gpu-toggle"]');
    if(gpuButton)gpuButton.textContent=server?.gpuTarget?'선택 GPU 다시 장착':'선택 GPU 인출';
    const driveButton=rackControls.querySelector('[data-rack-action="drive-toggle"]');
    if(driveButton)driveButton.textContent=server?.driveTarget?'선택 드라이브 다시 장착':'선택 드라이브 인출';
    const fiberSelected=selectedRackPart?.assetId===a.id&&selectedRackPart.kind==='fiber-port';
    const storageSelected=storageRack?selectedRackPart?.label||'스토리지 부품을 선택하세요':'';
    const storageInfo=storageRack&&server?.driveCount?' · 드라이브 '+server.driveSlot+'/'+server.driveCount+(server.driveTarget?' 인출':' 장착'):'';
    const storageSummary=storageRack?' · '+storageSelected:'';

    rackControlsLabel.textContent='랙 '+a.id+' · '+(focusedRack?.id===a.id?(focusedRack.side==='rear'?(a.id==='odf'?'광단자함 전면 · 3인칭':'후면 시점 · 3인칭'):'전면 시점 · 3인칭'):'기본 시점')+
      ' · '+(networkRack?'스위치 ':storageRack?'쉘프 ':'서버 ')+serverLabel+gpuStatus+
      storageInfo+storageSummary+
      ' · 후면문 '+(door.targetRear?'열림':'닫힘')+
      ' · 광포트 '+(fiber?.ports.length||0)+'개'+
      (fiberSelected?' · 포트 '+(selectedRackPart.portIndex+1)+'/24 선택 · 노란색/광섬유 경로 강조':' · 광 포트 클릭 또는 선택·확대로 확인')+
      ' · 배선 '+(rackNetwork.diagnostic(a.id)?.filter||'all')+
      (patchStatus?' · 패치패널 LC 케이블 '+(patchStatus.plugged?'체결':patchStatus.connected?'체결 중':'분리'):'');
  }
  function expandRackControls(){
    if(!T.rackIsolateActive||!currentRackAsset())return false;
    showControls();
    rackPanelDisclosure?.setCollapsed(false);
    return !!rackControls&&!rackControls.hidden;
  }
  function runRackAction(action){
    // Single action implementation: sidebar shortcuts reuse the same
    // proven 3D rack buttons; avoid divergent state between two panels.
    if(!T.rackIsolateActive||!currentRackAsset())return false;
    showControls();
    // A sidebar shortcut must work even if the floating controls were folded manually.
    rackPanelDisclosure?.setCollapsed(false);
    const allowed=/^(front-camera|rear-camera|front-door|rear-door|previous-server|next-server|server-slide|cassette-slide|next-gpu|gpu-toggle|next-drive|drive-toggle|storage-shelf-pull|storage-drive-pull|optical-detail|optical-explode|patch-select|switch-select|cassette-select|gpu-select|cable-patch|optical-lc|lc-rear|lc-next-port|odf-select|odf-cover|odf-next-port)$/;
    if(!allowed.test(action))return false;
    const button=[...(rackControls?.querySelectorAll('[data-rack-action]')||[])].find(x=>x.dataset.rackAction===action);
    if(!button||button.hidden||button.closest('[hidden]'))return false;
    button.click();
    return true;
  }
  let selectedPartVisuals=[], selectedPartRing=null;
  function clearRackPartEmphasis() {
    for(const entry of selectedPartVisuals){
      if(entry.object.material===entry.highlighted)entry.object.material=entry.original;
      const list=Array.isArray(entry.highlighted)?entry.highlighted:[entry.highlighted];
      for(const material of list)material?.dispose?.();
    }
    selectedPartVisuals=[];
    if(selectedPartRing){
      selectedPartRing.parent?.remove(selectedPartRing);
      selectedPartRing.geometry?.dispose?.();selectedPartRing.material?.dispose?.();
      selectedPartRing=null;
    }
  }
  function emphasizeRackPart(assetId,partId,{ring=true,intensity=2.2,colorMix=.72}={}) {
    clearRackPartEmphasis();
    const item=registry.get(assetId);if(!item)return null;
    const rearMpo=partId.includes(':mpo-input-');
    let anchor=null;
    item.g.updateMatrixWorld(true);
    item.g.traverse(object=>{
      if(!object.isMesh||object.userData?.partId!==partId||!object.material)return;
      if(rearMpo&&!object.name.includes('mpo-keyed-receptacle-'))return;
      const original=object.material;
      const highlighted=Array.isArray(original)?original.map(m=>m.clone()):original.clone();
      const mats=Array.isArray(highlighted)?highlighted:[highlighted];
      for(const material of mats){
        if(material.emissive){material.emissive.set(0xffdf55);material.emissiveIntensity=intensity;}
        if(material.color)material.color.lerp(new THREE.Color(0xffef8a),colorMix);
      }
      object.material=highlighted;
      selectedPartVisuals.push({object,original,highlighted});
      if(!anchor||object.name.includes('connector-boot'))anchor=object;
    });
    if(ring&&anchor&&THREE.TorusGeometry&&THREE.MeshBasicMaterial){
      selectedPartRing=new THREE.Mesh(
        new THREE.TorusGeometry(.105,.018,8,24),
        new THREE.MeshBasicMaterial({color:0xffef68,depthTest:false,toneMapped:false})
      );
      selectedPartRing.name='rack-optical:selected-port-ring';
      selectedPartRing.userData.assetId=assetId;selectedPartRing.userData.partId=partId;
      selectedPartRing.position.set(0,0,rearMpo?.058:-.058);
      selectedPartRing.renderOrder=1001;
      anchor.add(selectedPartRing);
    }
    return anchor;
  }
  function focusOpticalPort(assetId,anchor) {
    const asset=T.assetMap.get(assetId),desired=T.cameraDesired,item=registry.get(assetId);
    if(!asset||!desired||!anchor||T.workerViewId||T.walkActive)return false;
    focusRackSide(assetId,'front');
    const motion=T.mechanical?.(asset),wasOpen=!!(motion?.targetCassette||motion?.cassette>.5);
    if(!wasOpen)toggleRackCassette(assetId);
    item?.g?.updateMatrixWorld?.(true);
    const p=new THREE.Vector3();anchor.getWorldPosition(p);
    const extension=wasOpen?0:(Number(item?.moving?.reach)||1)*.70;
    desired.target=[p.x,p.y,p.z-extension];
    desired.distance=Math.max(2.7,Math.min(4.2,asset.w*.75));
    desired.pitch=.12;desired.yaw=Math.PI;
    return true;
  }
  function showRackPart(info) {
    clearSceneInspection();
    clearRackPartEmphasis();
    if(selectedRackPart){
      if(!rackPatchDesign.owns(info)){
        rackPatchDesign.select(selectedRackPart.assetId,false);
        rackPatchDesign.isolate(selectedRackPart.assetId,false);
      }
      if(!rackOptics.owns(info))rackOptics.isolate(selectedRackPart.assetId,false);
      if(!rackLcDesign.owns(info)){
        rackLcDesign.isolate(selectedRackPart.assetId,false);
        rackLcDesign.highlight(selectedRackPart.assetId,-1);
      }
      if(!odfBox.owns(info)){
        odfBox.isolate(selectedRackPart.assetId,false);
        odfBox.highlight(selectedRackPart.assetId,-1);
      }
    }
    selectedRackPart=info?{...info}:null;
    if(rackPatchDesign.owns(selectedRackPart)){
      rackPatchDesign.select(selectedRackPart.assetId,true);
      rackPatchDesign.isolate(selectedRackPart.assetId,true);
      if(selectedRackPart.kind==='patch-panel')rackPatchDesign.includeCable(selectedRackPart.assetId,false);
    }
    if(rackOptics.owns(selectedRackPart))rackOptics.isolate(selectedRackPart.assetId,true);
    if(rackLcDesign.owns(selectedRackPart)){
      rackLcDesign.isolate(selectedRackPart.assetId,true);
      rackLcDesign.highlight(selectedRackPart.assetId,selectedRackPart.kind==='lc-adapter-port'?selectedRackPart.portIndex:-1);
    }
    if(odfBox.owns(selectedRackPart)){
      odfBox.isolate(selectedRackPart.assetId,true);
      odfBox.highlight(selectedRackPart.assetId,selectedRackPart.kind==='odf-small-port'?selectedRackPart.portIndex:-1);
      if(['odf-small-port','odf-small-splice','odf-small-guide'].includes(selectedRackPart.kind))odfBox.open(selectedRackPart.assetId);
    }
    const opticalInspect=selectedRackPart?.kind==='patch-panel'||rackOptics.owns(selectedRackPart)||rackLcDesign.owns(selectedRackPart)||odfBox.owns(selectedRackPart);
    if(opticalInspect)
      rackPanelDisclosure?.setCollapsed(true);
    // Keep the mobile lens ruler from covering one of the four cassettes.
    // Its own toggle remains usable; restore only when we collapsed it.
    const scaleButton=typeof document!=='undefined'?document.getElementById('phase1-toggle'):null;
    const scale=scaleButton?.closest('.phase1-scale');
    if(opticalInspect&&scaleButton&&!scale?.classList.contains('is-collapsed')){
      scaleButton.click();opticalScaleAutoCollapsed=true;
    }else if(!opticalInspect&&opticalScaleAutoCollapsed){
      if(scale?.classList.contains('is-collapsed'))scaleButton?.click();
      opticalScaleAutoCollapsed=false;
    }
    window.LS3D_INSPECTOR?.setSelection(selectedRackPart);
    if(typeof document==='undefined'||!canvas.parentElement)return selectedRackPart;
    if(!rackHud) {
      rackHud=document.createElement('div');
      rackHud.id='ls3dRackPartInfo';
      rackHud.style.cssText='display:none'; // Selection details live in the right inspector.
      canvas.parentElement.appendChild(rackHud);
    }
    rackHud.hidden=!selectedRackPart;
    if(selectedRackPart)rackHud.textContent=
      '선택 부품: '+selectedRackPart.label+' · '+selectedRackPart.assetId+
      ' · U '+selectedRackPart.uStart+'-'+(selectedRackPart.uStart+selectedRackPart.uHeight-1)+
      (selectedRackPart.kind==='fiber-port'?' · 포트 '+(selectedRackPart.portIndex+1)+'/24 · 광 경로 노란색 강조':'')+
      (selectedRackPart.description?' · '+selectedRackPart.description:'');
    return selectedRackPart;
  }
  function selectRackPart(assetId,partId) {
    const record=registry.get(assetId);
    const part=record?.partMap?.get(partId)||null;
    const selected=showRackPart(part);
    if(part?.kind==='fiber-port'){
      const anchor=emphasizeRackPart(assetId,partId);
      focusOpticalPort(assetId,anchor);
    }else if(part?.kind==='optical-module'||part?.kind==='mpo-input'){
      const asset=T.assetMap.get(assetId),desired=T.cameraDesired;
      if(asset&&desired&&T.rackIsolateActive&&!T.workerViewId&&!T.walkActive){
        const rear=part.kind==='mpo-input';focusRackSide(assetId,rear?'rear':'front');
        const unit=(asset.h-.46)/(record.rackUnits||42),u=record.opticsSession?.base||.23+(part.uStart-1)*unit;
        const mid=rear?record.opticsSession?.back:(record.opticsSession?.front+record.opticsSession?.back)*.5;
        const extension=(record.moving?.reach||asset.d*.67)*.70;
        const module=record.opticsSession?.modules?.[part.moduleIndex??0];
        const point=new THREE.Vector3(rear?(module?.x||0):0,rear?u+unit*.25:u+unit*.9,
          mid-extension+(rear?.035:0));
        record.g.updateMatrixWorld(true);record.g.localToWorld(point);
        const rect=canvas.getBoundingClientRect();
        const usable=(rect.width-(rect.width>640?330:20))/Math.max(1,rect.height);
        desired.target=[point.x,point.y,point.z];
        desired.distance=rear?Math.max(1.45,asset.w*.62):
          Math.max(2.4,asset.w*.78/(2*Math.tan(Math.PI/8)*Math.max(.55,usable))*1.30);
        desired.yaw=rear?0:Math.PI;desired.pitch=rear?.14:.88;
        if(T.camera)Object.assign(T.camera,{target:desired.target.slice(),distance:desired.distance,yaw:desired.yaw,pitch:desired.pitch});
      }
      if(part.kind==='mpo-input')emphasizeRackPart(assetId,partId);
    }else if(rackLcDesign.owns(part)){
      const asset=T.assetMap.get(assetId),desired=T.cameraDesired,s=record.lcSession;
      if(asset&&desired&&s&&T.rackIsolateActive&&!T.workerViewId&&!T.walkActive){
        focusRackSide(assetId,'front');
        const point=new THREE.Vector3(0,s.lower+s.height*.5,s.front);
        record.g.updateMatrixWorld(true);record.g.localToWorld(point);
        desired.target=[point.x,point.y,point.z];desired.distance=Math.max(2.15,asset.w*1.03);
        desired.yaw=Math.PI;desired.pitch=.13;
        if(T.camera)Object.assign(T.camera,{target:desired.target.slice(),distance:desired.distance,yaw:desired.yaw,pitch:desired.pitch});
      }
      if(part.kind==='lc-adapter-port')emphasizeRackPart(assetId,partId);
    }else if(odfBox.owns(part)){
      const asset=T.assetMap.get(assetId),desired=T.cameraDesired,s=record.odfSmallSession;
      if(asset&&desired&&s&&T.rackIsolateActive&&!T.workerViewId&&!T.walkActive){
        focusRackSide(assetId,'rear');
        const point=new THREE.Vector3(0,s.height*.5,0);
        s.root.updateMatrixWorld(true);s.root.localToWorld(point);
        desired.target=[point.x,point.y,point.z];desired.distance=canvas.getBoundingClientRect().width<640?3.05:2.35;
        desired.yaw=0;desired.pitch=.24;
        if(T.camera)Object.assign(T.camera,{target:desired.target.slice(),distance:desired.distance,yaw:desired.yaw,pitch:desired.pitch});
      }
      if(part.kind==='odf-small-port')emphasizeRackPart(assetId,partId);
    }else if(part?.kind==='patch-panel'){
      const asset=T.assetMap.get(assetId),desired=T.cameraDesired;
      if(asset&&desired&&T.rackIsolateActive&&!T.workerViewId&&!T.walkActive){
        focusRackSide(assetId,'front');
        const unit=(asset.h-.46)/(record.rackUnits||42);
        const point=new THREE.Vector3(0,.23+(part.uStart-.5)*unit+asset.d*.09,-asset.d*.347);
        record.g.updateMatrixWorld(true);record.g.localToWorld(point);
        const rect=canvas.getBoundingClientRect();
        const usable=(rect.width-(rect.width>640?330:20))/Math.max(1,rect.height);
        desired.target=[point.x,point.y,point.z];
        desired.distance=Math.max(2.85,asset.w*.78/(2*Math.tan(Math.PI/8)*Math.max(.55,usable))*1.38);
        desired.yaw=Math.PI-.19;desired.pitch=.83;
        if(T.camera)Object.assign(T.camera,{target:desired.target.slice(),distance:desired.distance,yaw:desired.yaw,pitch:desired.pitch});
      }
    }else if(part?.kind==='patch-internal'){
      // Keep the open-panel view while indicating the selected internal part.
      emphasizeRackPart(assetId,partId,{ring:false,intensity:.75,colorMix:.28});
    }else if(part){
      emphasizeRackPart(assetId,partId);
    }
    if(part?.kind?.startsWith('storage-drive-'))rackRails.selectDrive(assetId,partId);
    return selected;
  }
  function visitorCaption(a) {
    const explicit={
      'entry-desk':['SECURITY DESK','VISITOR CHECK-IN'],
      'site-sign':['LS','DATA CENTER'],
      'visitor-lobby':['LOBBY · BRIEFING ROOM','SAFETY ORIENTATION'],
      'visitor-staging':['STAGING · 서버 설치 전 준비실','입고·개봉 → 검수·설정·시험 → 서버홀'],
      'visitor-loading':['LS CABLE · LOADING BAY','DELIVERY / STAGING'],
      'visitor-locker':['VISITOR ID CAMERA','ID / CAMERA · USB · LAPTOP'],
      'visitor-mantrap':['BIOMETRIC MANTRAP','IRIS · FINGERPRINT · CARD'],
      'generator':['EMERGENCY GENERATOR','DIESEL · RADIATOR · EXHAUST'],
      'tower-a':['COOLING TOWER A','FAN · DRIFT ELIMINATOR'],
      'tower-b':['COOLING TOWER B','FAN · DRIFT ELIMINATOR'],
      'chiller':['CHILLER','COMPRESSOR · CONDENSER'],
      'xfmr-a':['MAIN TRANSFORMER A','UTILITY INTAKE'],
      'xfmr-b':['MAIN TRANSFORMER B','UTILITY INTAKE'],
      'switchgear':['SWITCHGEAR','PROTECTION · DISTRIBUTION'],
      'ats':['ATS','UTILITY / GENERATOR TRANSFER'],
      'battery-a':['BATTERY BANK A','DC STRING · UPS BACKUP'],
      'ups-a':['UPS A','ONLINE UPS · 500 kVA'],
      'ups-b':['UPS B','STANDBY UPS · 500 kVA']
    };
    const text=explicit[a.id];if(!text||typeof document==='undefined')return null;
    const c=document.createElement('canvas');c.width=640;c.height=150;
    const ctx=c.getContext('2d');if(!ctx)return null;
    ctx.fillStyle='rgba(8,29,43,.92)';ctx.strokeStyle='#68d9d1';ctx.lineWidth=7;
    ctx.beginPath();ctx.roundRect?.(8,8,624,134,18);ctx.fill();ctx.stroke();
    ctx.fillStyle='#f0fbff';ctx.font='700 42px "Pretendard Variable","Malgun Gothic","Apple SD Gothic Neo",sans-serif' ;ctx.textAlign='center';
    ctx.fillText(text[0],320,64);ctx.fillStyle='#83e1d5';ctx.font='600 26px "Pretendard Variable","Malgun Gothic","Apple SD Gothic Neo",sans-serif' ;
    ctx.fillText(text[1],320,108);
    const texture=new THREE.CanvasTexture(c);texture.colorSpace=THREE.SRGBColorSpace;
    const sprite=new THREE.Sprite(new THREE.SpriteMaterial({map:texture,transparent:true,depthWrite:false}));
    sprite.name='visitor-caption:'+a.id;sprite.scale.set(Math.max(3.8,a.w*.88),Math.max(.9,a.h*.23),1);
    sprite.position.set(0,a.h+Math.max(.55,a.h*.15),-a.d*.16);
    sprite.userData.noPick=true;return sprite;
  }
  function assetMesh(a) {
    const g=new THREE.Group();
    g.name='equipment:'+a.id;g.userData.assetId=a.id;
    const {w,d,h}=a;
    let moving=null,rackLOD=null,equipmentLOD=null,kind='rack',rackParts=[],partMap=null,rackUnits=null,fuelSmoke=[],loadingBayActors=[];
    let serverSession=null,opticsSession=null,networkSession=null,odfSmallSession=null,lcSession=null;
    if(T.isMechanicalRack(a)){
      const high=new THREE.Group();high.name='rack-high';g.add(high);
      const built=rackKit.build(high,a);
      serverSession=rackRails.decorate(a,built,high);
      opticsSession=rackOptics.decorate(a,built);
      networkSession=rackNetwork.decorate(a,built,high);
      rackPatchDesign.decorate(a,built,high);
      rackPatch.decorate(a,built,high);
      lcSession=rackLcDesign.decorate(a,built,high);
      odfSmallSession=odfBox.decorate(a,built,high);
      moving=built.moving;rackParts=built.parts;partMap=built.partMap;
      rackUnits=built.rackUnits;
      const low=new THREE.Group();low.name='rack-low';g.add(low);
      cube(low,0,.06,0,w*.92,h*.94,d*.90,'#53728b',.52,.44);
      cube(low,0,h*.72,-d*.48,w*.75,.26,.08,'#56a7b7',.32,.34);
      rackLOD={high,low};
    } else if(a.id==='colo'&&a.type==='Customer cage'){
      // STEP 16 draws the real three-tenant fence in a separate world layer.
      // Suppress the old single-cage generic equipment shell to avoid overlaying it.
      kind='customer-cage-anchor';
      cube(g,0,.01,0,w*.93,.065,d*.9,'#618996',.24,.62);
    } else {
      const result=equipmentFactory.build(a);
      kind=result.kind;equipmentLOD=result.lod;g.add(equipmentLOD);
    }
    cube(g,0,-.09,0,w*1.08,.11,d*1.08,'#355269',.12,.88);
    const caption=visitorCaption(a);if(caption)g.add(caption);
    // Loading Bay is approached from its cargo/carton side (+Z); unlike the
    // other Gray Zone rooms it must face the default visitor camera directly.
    const baseYaw=a.id==='visitor-mantrap'?Math.PI:(a.zone==='gray'||a.zone==='network')&&!['gray-cdu','visitor-loading','visitor-staging'].includes(a.id)?Math.PI:0;
    g.position.set(a.x,0,a.z);g.rotation.y=baseYaw;
    const selection=new THREE.Mesh(
      new THREE.RingGeometry(Math.max(w,d)*.56,Math.max(w,d)*.66,36),
      new THREE.MeshBasicMaterial({color:0x63e4d8,side:THREE.DoubleSide,transparent:true,opacity:.94})
    );
    selection.rotation.x=-Math.PI/2;selection.position.y=.027;
    selection.name='selection-ring';selection.userData.noPick=true;g.add(selection);
    const mantrapLights={red:null,green:null},mantrapDoors=[],maintenanceActors=[];
    g.traverse(obj=>{
      if(obj.isMesh)obj.userData.assetId=a.id;
      if(obj.name?.startsWith('equipment-mantrap-open-door:'))mantrapDoors.push({joint:obj,openAngle:obj.rotation.y});
      if(obj.name==='equipment-part:mantrap-card-led-red')mantrapLights.red=obj;
      if(obj.name==='equipment-part:mantrap-card-led-green')mantrapLights.green=obj;
      if(obj.userData?.fuelSmoke)fuelSmoke.push(obj);
      if(obj.userData?.maintenanceAnimation)maintenanceActors.push(obj);
      if(obj.userData?.loadingBayAnimation)loadingBayActors.push(obj);
    });
    return {g,moving,selection,kind,equipmentLOD,rackLOD,rackParts,partMap,rackUnits,serverSession,opticsSession,networkSession,odfSmallSession,lcSession,mantrapLights,mantrapDoors,fuelSmoke,maintenanceActors,loadingBayActors,baseYaw};
  }
  // Twelve-second educational operation. One employee loads a box, boards the
  // 50%-size forklift, drives forward, reverses, and dismounts. The vehicle
  // never moves until the employee has boarded. The existing frame drives it.
  const clamp01=x=>Math.max(0,Math.min(1,x));
  const smooth01=x=>{const t=clamp01(x);return t*t*(3-2*t);};
  const lerp=(a,b,t)=>a+(b-a)*t;
  function animateLoadingBay(forklift,now) {
    const m=forklift.userData.loadingBayAnimation;if(!m)return;
    const t=(now%m.loopSeconds+m.loopSeconds)%m.loopSeconds;
    const approach=smooth01(t/1.5);
    const load=smooth01((t-1.5)/1.15);
    const board=smooth01((t-2.65)/1.20);
    const forward=smooth01((t-3.90)/2.15);
    const reverse=smooth01((t-6.35)/2.15);
    const exit=smooth01((t-8.55)/1.20);
    const travel=forward-reverse;
    const vehicleX=m.homeX-m.travel*travel;
    forklift.position.set(vehicleX,m.homeY,m.parkZ);
    // Employee remains alongside the fork during loading, and moves to the
    // driver's seat before forward movement begins. The same rig is reused.
    const worker=m.worker;
    let x,y,z,scale;
    if(t<2.65){
      x=lerp(m.workerHomeX,m.workerNearX,approach);
      y=.12+Math.sin(Math.PI*approach)*.018;
      z=m.workerSideZ;scale=1;
    }else if(t<3.85){
      x=lerp(m.workerNearX,m.homeX+m.seatOffsetX,board);
      y=lerp(.12,.39,board);
      z=lerp(m.workerSideZ,m.parkZ,board);
      scale=lerp(1,.67,board);
    }else if(t<8.55){
      x=vehicleX+m.seatOffsetX;y=.39;z=m.parkZ;scale=.67;
    }else{
      x=lerp(vehicleX+m.seatOffsetX,m.workerHomeX,exit);
      y=lerp(.39,.12,exit);
      z=lerp(m.parkZ,m.workerSideZ,exit);
      scale=lerp(.67,1,exit);
    }
    worker.position.set(x,y,z);
    worker.scale.setScalar(scale);
    worker.rotation.y=-Math.PI/2;
    const holding=t<2.65;
    m.armLeft.rotation.x=holding?-.9:t<8.55?-.46:-.22;
    m.armRight.rotation.x=holding?-.9:t<8.55?-.46:-.22;
    m.carriedBox.visible=holding;
    m.loadedBox.visible=t>=2.65&&t<11.70;
    if(holding){
      const from=[x-.30,y+.90,z+.10];
      const to=[m.homeX-m.loadOffsetX,m.homeY+.39*.5,m.parkZ];
      m.carriedBox.position.set(
        lerp(from[0],to[0],load),lerp(from[1],to[1],load),lerp(from[2],to[2],load)
      );
      m.carriedBox.rotation.y=lerp(-.18,0,load);
    }
    return t<1.5?'직원 접근':t<2.65?'포크 적재':t<3.85?'운전석 탑승':
      t<6.05?'전진':t<6.35?'정지':t<8.50?'후진':
      t<9.75?'하차':'다음 작업 준비';
  }
  function refreshAssets(flags,dt=.016) {
    const showDetail = !!flags.inspect;
    if(selectedRackPart&&selectedRackPart.assetId!==T.selected?.id)showRackPart(null);
    const visible = T.visibleAssets().filter(a => (T.currentZone === 'campus' || a.zone === T.currentZone) &&
      (!showDetail || a.id === T.selected?.id) &&
      (!flags.coloTour || a.id==='colo' || /^colo-rack-[abc]$/.test(a.id)));
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
      const opticalIsolated=selectedRackPart?.assetId===a.id&&
        (rackOptics.owns(selectedRackPart)||rackPatchDesign.owns(selectedRackPart)||rackLcDesign.owns(selectedRackPart)||odfBox.owns(selectedRackPart));
      item.selection.visible=T.selected?.id===a.id&&!opticalIsolated;
      if(selectedRackPart && selectedRackPart.assetId===a.id && !item.partMap?.has(selectedRackPart.id))
        showRackPart(null);
      if(item.rackLOD) {
        const dx=camera.position.x-a.x, dy=camera.position.y-a.h*.5, dz=camera.position.z-a.z;
        const far=T.selected?.id!==a.id && Math.hypot(dx,dy,dz)>96;
        item.rackLOD.high.visible=!far;item.rackLOD.low.visible=far;
      }
      if(item.equipmentLOD){
        // Selected/tour-target equipment must never collapse into the low-detail box model.
        if(T.selected?.id===a.id){
          item.equipmentLOD.levels.forEach((level,index)=>{level.object.visible=index===0;});
        }else item.equipmentLOD.update(camera);
      }
      if(item.loadingBayActors?.length){
        const now=performance.now()/1000;
        // Both high- and mid-LOD copies receive the same deterministic pose.
        // No independent animation timers; the hidden tab does not advance frames.
        for(const forklift of item.loadingBayActors)
          diagnostic.loadingBayPhase=animateLoadingBay(forklift,now);
      }
      if(item.kind==='biometric-mantrap'){
        const tour=window.__LS3D_VISITOR__?.state;
        const demo=tour?.gateDemo;
        const approved=!!(demo?.active&&demo.openness>.5);
        for(const door of item.mantrapDoors||[])door.joint.rotation.y=door.openAngle*(demo?.active?demo.openness:1);
        if(item.mantrapLights?.red)item.mantrapLights.red.visible=!approved;
        if(item.mantrapLights?.green)item.mantrapLights.green.visible=approved;
      }
      const fault = !!a.down || !!(flags.scenario !== 'normal' && (
        (flags.scenario === 'cooling' && /tower/i.test(a.id)) ||
        (flags.scenario === 'network' && a.id === 'core') ||
        (flags.scenario === 'power' && a.id === 'switchgear')
      ));
      if (fault) item.g.scale.y = 1 + Math.sin((T.simTime || 0) * 5) * .008;
      else item.g.scale.y = 1;
      if(item.fuelSmoke?.length){
        const now=(T.simTime||performance.now()/1000);
        for(const puff of item.fuelSmoke){
          const smoke=puff.userData.fuelSmoke, rise=(now*.17+smoke.phase)%1;
          const scale=smoke.scale*(.82+rise*.72);
          puff.position.set(
            smoke.baseX+Math.sin(now*.70+smoke.phase)*smoke.drift*(.5+rise),
            smoke.baseY+rise*smoke.rise,
            smoke.baseZ+Math.cos(now*.52+smoke.phase)*smoke.drift*(.35+rise)
          );
          puff.scale.set(scale,scale*.72,scale*.91);
          if(puff.material)puff.material.opacity=smoke.opacity*(1-rise*.58);
        }
      }
      if(item.maintenanceActors?.length){
        const now=(T.simTime||performance.now()/1000);
        for(const worker of item.maintenanceActors){
          const motion=worker.userData.maintenanceAnimation;
          const wave=Math.sin(now*.72+motion.phase), pace=Math.cos(now*2.16+motion.phase);
          worker.position.x=motion.baseX+wave*motion.range;
          worker.rotation.y=wave>=0?Math.PI/2:-Math.PI/2;
          const carton=worker.getObjectByName('maintenance-carrying-carton');
          if(carton){carton.position.y=a.h*.40+Math.abs(pace)*.045;carton.rotation.z=pace*.075;}
        }
      }
      if (item.moving) {
        const m=T.mechanical(a);
        rackDoors.tick(a.id,item.moving,m,dt);
        rackRails.update(a.id,item.moving,m,dt);
        rackOptics.update(a.id,item.moving,m);
        odfBox.update(a.id,dt);
        rackNetwork.update(a.id,rackRails.state(a.id)?.selected||null);
        if(rackOptics.owns(selectedRackPart))rackOptics.isolate(a.id,true);
        if(rackLcDesign.owns(selectedRackPart))rackLcDesign.isolate(a.id,true);
        if(odfBox.owns(selectedRackPart))odfBox.isolate(a.id,true);
        rackPatchDesign.update(a.id,dt);
        rackPatch.update(a.id,dt);
        const cord=rackPatch.state(a.id);
        if(cord&&!cord.connected&&cord.progress<.01)rackPatchDesign.includeCable(a.id,false);
        // When using the STEP 11 independent side camera, keep world-space rack
        // orientation fixed. Original legacy rear flip remains otherwise available.
        item.g.rotation.y=item.baseYaw+(m.rear?Math.PI:0);
      }
    }
    for(const [id,item] of registry)if(!present.has(id)){disposeAsset(item);registry.delete(id);}
    if(selectedRackPart&&!registry.has(selectedRackPart.assetId))showRackPart(null);
    if(focusedRack&&!registry.has(focusedRack.id))focusedRack=null;
    showControls();
    diagnostic.assets=registry.size;
    diagnostic.rackNetwork={
      stats:{...rackNetwork.stats},
      selected:currentRackAsset()?rackNetwork.diagnostic(currentRackAsset().id):null
    };
    diagnostic.rackInteraction={
      focus:focusedRack?{...focusedRack}:null,
      doors:{...rackDoors.stats},
      rails:{...rackRails.stats},
      optics:{...rackOptics.stats},
      selectedServer:currentRackAsset()?rackRails.state(currentRackAsset().id):null,
      selectedOptics:currentRackAsset()?rackOptics.state(currentRackAsset().id):null,
      selectedOdfBox:currentRackAsset()?odfBox.state(currentRackAsset().id):null,
      selectedLcPanel:currentRackAsset()?rackLcDesign.state(currentRackAsset().id):null
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
  // STEP 17: calibrated as *example geometry*, unlike the old WebGL2 flat lines.
  // This layer anchors curves to actual equipment IDs and shares the scene RAF.
  function routeScene(flags) {
    diagnostic.mep=mepRoutes.update(flags.coloTour?{...flags,inspect:true}:flags);
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
      if(flags.inspect||flags.coloTour)continue;
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
  let viewportWidth=0, viewportHeight=0, viewportDpr=0, rackAutoFocused=null;
  function render(dt, view) {
    if (disposed) throw new Error('Three.js adapter was disposed');
    const rect = canvas.getBoundingClientRect();
    if (rect.width < 1 || rect.height < 1) return;
    // Frame only once on entering an isolated rack or switching rack.
    // Do not fight the visitor's mouse/keyboard movements by re-centering
    // every time the camera passes the old focusedRack distance threshold.
    const inspectedId=T.rackIsolateActive&&T.selected&&T.isMechanicalRack(T.selected)?T.selected.id:null;
    if(inspectedId!==rackAutoFocused){
      rackAutoFocused=inspectedId;
      if(inspectedId)focusRackSide(inspectedId,'front');
    }
    // STEP 18: sustained slow frames reduce MEDIUM DPR only; explicit low/high stays fixed.
    viewportPolicy.sample(dt);
    const quality = window.LS3D_CONFIG?.quality || 'medium';
    const {dpr}=viewportPolicy.settings({
      width:rect.width,quality,devicePixelRatio:window.devicePixelRatio||1
    });
    // Avoid mobile shadow-map cost on LOW quality or sustained MEDIUM slowdown.
    const reducedShadows=quality==='low'||
      (rect.width<=640&&quality==='medium'&&viewportPolicy.summary().degraded);
    renderer.shadowMap.enabled=!reducedShadows;
    const pixelWidth=Math.floor(rect.width),pixelHeight=Math.floor(rect.height);
    if(viewportWidth!==pixelWidth || viewportHeight!==pixelHeight || viewportDpr!==dpr) {
      renderer.setPixelRatio(dpr);
      renderer.setSize(pixelWidth, pixelHeight, false);
      viewportWidth=pixelWidth;viewportHeight=pixelHeight;viewportDpr=dpr;
    }
    // Dedicated rack inspection never lets a room wall, roof, neighboring rack,
    // or a campus MEP/cage layer fill the front/rear camera view.
    if(focusedRack&&!rackFocusActive(focusedRack,T))focusedRack=null;
    const isolatedRack=!!focusedRack;
    const flags = {...(view.flags||{}),visitorGate:window.__LS3D_VISITOR__?.state?.gateDemo,inspect:!!view.flags?.inspect||isolatedRack};
    diagnostic.isolatedRack=isolatedRack?focusedRack.id:null;
    const eye = view.eye, direction = view.forward, up = view.up;
    camera.position.set(eye[0],eye[1],eye[2]);
    camera.up.set(up[0],up[1],up[2]);
    camera.lookAt(eye[0] + direction[0] * 20, eye[1] + direction[1] * 20, eye[2] + direction[2] * 20);
    camera.fov = view.fov * 180 / Math.PI;
    camera.aspect = rect.width / rect.height;
    camera.near=Math.max(.055,Math.min(.15,view.near||.08));
    camera.far=Math.max(600,view.far||600);
    camera.updateProjectionMatrix();
    if(inspectedScenePart&&((flags.inspect&&!inspectedScenePart.assetId)||sceneInspectionContext!==T.currentZone+'|'+T.facilityMode+'|'+T.selected?.id))clearInspectionSelection();
    zoneScene(flags, dt);
    refreshAssets(flags,dt);
    const rooms=zoneExtras.update(flags,dt);
    diagnostic.mmr=rooms.mmr;
    diagnostic.colocation=rooms.colocation;
    routeScene(flags);
    syncInspectionCatalog();
    // Hide only those room envelope meshes that actually block the near selected
    // rack/asset; all are restored on zoom-out, zone changes or dispose.
    diagnostic.occlusion=cameraOcclusion.update({
      eye,selected:T.selected,currentZone:T.currentZone,
      facilityMode:T.facilityMode,inspect:!!flags.inspect&&focusedRack?.id!==T.selected?.id,showSection:!!flags.showSection
    });
    diagnostic.viewport=viewportPolicy.summary();
    diagnostic.overlay=layoutSceneOverlays({
      width:rect.width,height:rect.height,
      mmr:document.getElementById('ls3dMMRControls'),
      colo:document.getElementById('ls3dColoControls'),
      mep:document.getElementById('ls3dMEPControls'),
      rack:rackControls
    });
    securityState=security.simulate(T.simTime||0);
    security.root.visible=securityState.active&&!flags.inspect&&!flags.coloTour;
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
    const targets=[...assetRoot.children,
      ...(zoneRoot.visible?zoneRoot.children:[]),
      ...(zoneExtras.mmr.root.visible?zoneExtras.mmr.root.children:[]),
      ...(zoneExtras.colocation.root.visible?zoneExtras.colocation.root.children:[]),
      ...(routeRoot.visible?routeRoot.children:[])];
    for(const hit of raycaster.intersectObjects(targets,true)) {
      // Raycaster also intersects hidden descendants; only rendered surfaces may win.
      if (!isVisiblePickObject(hit.object)) continue;
      const structure=hit.object.userData?.inspectionPart;
      if(structure){selectScenePart(structure);return structure.id;}
      const routeSelection=routeRoot.visible?mepRoutes.pick(hit.object):null;
      if(routeSelection){selectScenePart({id:routeSelection.id,kind:'mep-route',label:routeSelection.label+' 연결 구간',zoneId:routeSelection.fromZone,description:routeSelection.from+' → '+routeSelection.to+' · '+routeSelection.name+' · 교육용 연결 경로입니다.'});return routeSelection.id;}
      const areaSelection=zoneExtras.pick(hit.object);
      if(areaSelection){
        // MMR/ODF room meshes previously stopped here, so their bay state
        // changed without activating the corresponding equipment inspector.
        const areaAsset=T.assetMap.get(areaSelection.assetId);
        if(areaAsset){
          showRackPart(null);
          T.pickAsset(areaAsset);
          if(!T.isMechanicalRack(areaAsset))
            selectScenePart(describeEquipment(areaAsset,registry.get(areaAsset.id)?.kind));
        }
        return areaSelection.assetId||areaSelection.id;
      }
      let obj=hit.object,part=null;
      for(let p=obj;p&&p!==assetRoot;p=p.parent) {
        if(!part&&p.userData.rackPart)part=p.userData.rackPart;
      }
      while(obj && !obj.userData.assetId) obj=obj.parent;
      const id=obj?.userData.assetId, asset=T.assetMap.get(id);
      // STEP 13: a connector mesh carries a more specific partId than its
      // parent cassette drawer. Resolve by the owning rack's canonical partMap.
      const clickedPartId=hit.object.userData?.partId;
      const exactPart=clickedPartId&&registry.get(id)?.partMap?.get(clickedPartId);
      if(exactPart)part=exactPart;
      if(asset){
        // A tap on any visible rack surface selects the rack as a whole.
        // Components remain available through the dedicated rack-internal view.
         if(T.isMechanicalRack(asset)&&!focusedRack&&!T.rackIsolateActive&&!document.getElementById('detailModal')?.classList.contains('show'))part=null;
        if(part) {
          selectRackPart(id,part.id);
          if(part.kind?.startsWith('gpu-module-'))rackRails.selectGpu(id,part.id);
          if(part.kind?.startsWith('storage-drive-'))rackRails.selectDrive(id,part.id);
          if(part.kind?.startsWith('server-')||part.kind?.startsWith('network-switch-')||part.kind?.startsWith('storage-'))
            rackRails.select(id,part.id);
        } else showRackPart(null);
        T.pickAsset(asset);
        if(!T.isMechanicalRack(asset))selectScenePart(describeEquipment(asset,registry.get(id)?.kind));
        return id;
      }
    }
    showRackPart(null);
    return null;
  }
  function dispose() {
    if(disposed)return;disposed=true;
    clearSceneInspection();
    navigation.reset();inspection.reset();
    cameraOcclusion.dispose();
    zoneExtras.dispose();
    mepRoutes.dispose();
    if(rackHud){rackHud.remove();rackHud=null;}selectedRackPart=null;
    if(rackControls){rackControls.remove();rackControls=null;}
    rackPanelDisclosure=null;rackPanelHeading=null;rackControlsLabel=null;rackControlsActiveRack=null;
    rackDoors.reset();rackRails.reset();rackOptics.reset();rackNetwork.reset();rackPatch.reset();rackPatchDesign.reset();rackLcDesign.reset();odfBox.reset();focusedRack=null;
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
    get rackNetwork(){return diagnostic.rackNetwork;},
    get rackPatch(){return currentRackAsset()?rackPatch.state(currentRackAsset().id):null;},
    get mmr(){return diagnostic.mmr;},get colocation(){return diagnostic.colocation;},
    get mep(){return diagnostic.mep;},mepController:mepRoutes,
    get viewportPerformance(){return viewportPolicy.summary();},
    get cameraOcclusion(){return cameraOcclusion.summary();},
    rackCabling:rackNetwork,mmrController:zoneExtras.mmr,colocationController:zoneExtras.colocation,
     selectRackPart,selectInspectionPart,clearInspectionSelection,focusRackSide,clearRackFocus(){focusedRack=null;},setRackServerSelection,cycleRackServer,
    toggleRearRackDoor,toggleRackFrontDoor,toggleRackServerSlide,toggleRackCassette,cycleRackGpu,toggleRackGpu,expandRackControls,runRackAction,
    get rackPanel(){return rackPanelDisclosure?.summary()||null;},
    get rackInteraction(){return diagnostic.rackInteraction;},
    inspection,securityController:security,navigation,
    get disposed(){return disposed;} };
}


