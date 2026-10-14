import * as THREE from 'three';
import { createArchitectureKit } from './architecture.mjs?v=20261015-step04';

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
  function genericEquipment(root, a) {
    const type = String(a.type || ''), w = a.w, d = a.d, h = a.h;
    const rack = /rack|top-of-rack|core|spine|leaf|odf|switch/i.test(type) && !/switchgear/i.test(type);
    const color = /cool|chiller|tower|pump|CRAH|CDU/i.test(type) ? '#65a5bf'
      : /transform|UPS|PDU|switchgear|generator|battery/i.test(type) ? '#9ca9a5'
      : rack ? '#57748e' : '#6a94a9';
    cube(root, 0, .05, 0, w, Math.max(.5, h - .17), d, color, .52, .38);
    cube(root, 0, h * .16, d * .506, w * .79, h * .38, .13, '#1a354d', .34, .34);
    for (let j = 0; j < 4; j++) cube(root, -w * .31 + j * w * .205, h * .39, d * .589, w * .13, .075, .065, '#b7d0d9', .6, .3);
  }
  function makeRack(g, a, motion) {
    const w = a.w, d = a.d, h = a.h;
    const dark = '#1c334a', steel = '#7695ad';
    cube(g, -w * .45, 0, -d * .42, w * .085, h, d * .075, steel, .75, .3);
    cube(g, w * .45, 0, -d * .42, w * .085, h, d * .075, steel, .75, .3);
    cube(g, -w * .45, 0, d * .42, w * .085, h, d * .075, steel, .75, .3);
    cube(g, w * .45, 0, d * .42, w * .085, h, d * .075, steel, .75, .3);
    cube(g, 0, h - .16, 0, w, .16, d, '#90acc2', .78, .29);
    cube(g, 0, .06, 0, w, .16, d, steel, .74, .3);
    const slide = new THREE.Group(); slide.name = 'server-tray'; g.add(slide);
    for (let i = 0; i < 5; i++) {
      const y = .7 + (i * (h - 1.9)) / 5;
      const col = i === 0 ? '#5a7690' : '#263d53';
      cube(slide, 0, y, -d * .26, w * .77, .42, d * .49, col, .7, .32);
      for (let j = 0; j < 6; j++) cube(slide, -w * .31 + j * w * .125, y + .18, -d * .52, w * .07, .05, .052, '#8facc3', .65, .35);
    }
    cube(g, 0, h - .67, -d * .33, w * .79, .32, d * .65, '#367b88', .66, .32);
    for (let i = 0; i < 9; i++) cube(g, -w * .34 + i * w * .085, h - .5, -d * .678, w * .05, .105, .06, i % 2 ? '#5fe8bd' : '#8dd6fd', .1, .2);
    const cassette = new THREE.Group(); cassette.name = 'fiber-cassette'; g.add(cassette);
    cube(cassette, 0, h - 1.17, -d * .46, w * .79, .25, d * .2, '#d6ae66', .6, .36);
    const hinge = new THREE.Group(); hinge.name = 'front-door-hinge';
    hinge.position.set(-w * .5, .06, -d * .55);
    g.add(hinge);
    cube(hinge, w * .5, 0, 0, w, h - .12, .08, '#5f829a', .54, .4);
    for (let i = 0; i < 12; i++) cube(hinge, w * .5, .38 + i * ((h - .94) / 12), -.052, w * .75, .045, .055, dark);
    cube(g, w * .4, .42, d * .53, .13, h - .84, .14, '#f3d07c', .61, .33);
    cube(g, -w * .4, .42, d * .53, .13, h - .84, .14, '#f3d07c', .61, .33);
    return { hinge, slide, cassette, reach: d * .67, door: motion || null };
  }
  function assetMesh(a) {
    const g = new THREE.Group();
    g.name = 'equipment:' + a.id;
    g.userData.assetId = a.id;
    const t = String(a.type || '').toLowerCase();
    const w = a.w, d = a.d, h = a.h;
    let moving = null;
    if (T.isMechanicalRack(a)) moving = makeRack(g, a, T.mechanical(a));
    else if (/cooling tower/.test(t)) {
      cube(g, 0, 0, 0, w, .3, d, '#668ea5');
      cylinder(g, 0, .3, 0, w * .32, h - .7, '#83bbcd');
      cylinder(g, 0, h - .55, 0, w * .34, .25, '#3d6074');
      for (let i = 0; i < 4; i++) {
        const ang = i * Math.PI / 2;
        cube(g, Math.cos(ang) * w * .16, h - .3, Math.sin(ang) * w * .16, .15, .05, w * .25, '#a0bdc7');
      }
    } else if (/transformer/.test(t)) {
      genericEquipment(g, a);
      for (let i = -1; i <= 1; i++) cylinder(g, i * w * .26, h - .18, 0, .23, .8, '#ccd8d9');
      for (let i = 0; i < 5; i++) cube(g, -w * .36 + i * w * .18, h * .36, d * .56, .06, h * .32, .16, '#24415b');
    } else if (/chiller/.test(t)) {
      genericEquipment(g, a);
      for (let i = -1; i <= 1; i++) cylinder(g, i * w * .25, h - .15, 0, .45, .2, '#2f5367');
    } else if (/x-ray|xray/i.test(a.id + ' ' + a.type)) {
      genericEquipment(g, a);
      cube(g, -w * .72, .72, 0, w * .5, .16, d * .52, '#879ca9');
      cube(g, w * .72, .72, 0, w * .5, .16, d * .52, '#879ca9');
    } else if (/metal.detector/i.test(a.id + ' ' + a.type)) {
      cube(g, -w * .34, 0, 0, w * .17, h, d * .4, '#c6d5d9');
      cube(g, w * .34, 0, 0, w * .17, h, d * .4, '#c6d5d9');
      cube(g, 0, h - .23, 0, w * .85, .22, d * .4, '#92bccb');
    } else genericEquipment(g, a);
    cube(g, 0, -.09, 0, w * 1.08, .11, d * 1.08, '#355269', .12, .88);
    g.position.set(a.x, 0, a.z);
    const selection = new THREE.Mesh(new THREE.RingGeometry(Math.max(w,d) * .56, Math.max(w,d) * .66, 36), new THREE.MeshBasicMaterial({ color: 0x63e4d8, side: THREE.DoubleSide, transparent: true, opacity: .94 }));
    selection.rotation.x = -Math.PI / 2; selection.position.y = .027; selection.name = 'selection-ring'; selection.userData.noPick = true; g.add(selection);
    g.traverse(obj => { if (obj.isMesh) obj.userData.assetId = a.id; });
    return { g, moving, selection };
  }
  function refreshAssets(flags) {
    const showDetail = !!flags.inspect;
    const visible = T.visibleAssets().filter(a => (T.currentZone === 'campus' || a.zone === T.currentZone) && (!showDetail || a.id === T.selected?.id));
    const present = new Set();
    for (const a of visible) {
      present.add(a.id);
      const sig = [a.zone,a.type,a.x,a.z,a.w,a.d,a.h].join('|');
      let item = registry.get(a.id);
      if (!item || item.sig !== sig) {
        if (item) { assetRoot.remove(item.g); item.g.traverse(o => { if (o.geometry && o.geometry !== boxGeometry && o.geometry !== cylinderGeometry) o.geometry.dispose(); }); }
        const model = assetMesh(a);
        item = { ...model, sig }; registry.set(a.id, item); assetRoot.add(item.g);
      }
      item.selection.visible = T.selected?.id === a.id;
      const fault = !!a.down || !!(flags.scenario !== 'normal' && (
        (flags.scenario === 'cooling' && /tower/i.test(a.id)) ||
        (flags.scenario === 'network' && a.id === 'core') ||
        (flags.scenario === 'power' && a.id === 'switchgear')
      ));
      if (fault) item.g.scale.y = 1 + Math.sin((T.simTime || 0) * 5) * .008;
      else item.g.scale.y = 1;
      if (item.moving) {
        const m = T.mechanical(a);
        item.moving.hinge.rotation.y = -(m.door || 0) * Math.PI * .53;
        item.moving.slide.position.z = -(m.slide || 0) * item.moving.reach;
        item.moving.cassette.position.z = -(m.cassette || 0) * item.moving.reach * .7;
        item.g.rotation.y = m.rear ? Math.PI : 0;
      }
    }
    for (const [id,item] of registry) if (!present.has(id)) { assetRoot.remove(item.g); item.g.traverse(o=>{if(o.geometry && o.geometry!==boxGeometry && o.geometry!==cylinderGeometry)o.geometry.dispose();}); registry.delete(id); }
    diagnostic.assets = registry.size;
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
  function personMesh(n) {
    const g = new THREE.Group();
    const uniform = n.role === 'guard' ? '#2d5681' : '#59b6c8';
    cube(g, 0, 1.04, 0, .54, .78, .34, uniform, .1, .75);
    cube(g, 0, 1.8, 0, .28, .14, .24, '#d7c29d', .02, .9);
    const head = new THREE.Mesh(new THREE.SphereGeometry(.21, 12, 8), material('#caa982',.05,.85));
    head.position.y = 2.04; g.add(head);
    cube(g, 0, 2.22, 0, .54, .13, .46, n.role === 'guard' ? '#28476a' : '#f0c46a', .11, .61);
    const limbs = [];
    for (const x of [-.21,.21]) {
      const arm = new THREE.Group(); arm.position.set(x,1.72,0); g.add(arm);
      cube(arm,0,-.52,0,.16,.56,.17,uniform,.05,.85); limbs.push(arm);
    }
    for (const x of [-.15,.15]) {
      const leg = new THREE.Group(); leg.position.set(x,1.04,0);g.add(leg);
      cube(leg,0,-.71,0,.2,.7,.22,'#324c65',.09,.8); limbs.push(leg);
    }
    g.userData.limbs=limbs; return g;
  }
  function refreshPeople(flags) {
    const all = T.npcs || [], visible = new Set();
    const staff = (T.facilityProfiles?.[T.facilityMode]?.staff || []);
    for (const n of all) {
      if (T.facilityMode !== 'aidc' && n.role === 'guard') continue;
      if (n.role !== 'guard' && !staff.includes(n.id)) continue;
      if (T.currentZone !== 'campus' && n.zone !== T.currentZone) continue;
      if (flags.inspect) continue;
      const pose = T.npcPose(n); if (!pose || !Number.isFinite(pose.x) || !Number.isFinite(pose.z)) continue;
      visible.add(n.id);
      let g = people.get(n.id);
      if (!g) { g=personMesh(n); people.set(n.id,g); peopleRoot.add(g); }
      g.position.set(pose.x,.04,pose.z);g.rotation.y = pose.heading || 0;
      const stride = pose.action === 'walk' ? Math.sin((T.simTime || 0) * 4.1) * .34 : 0;
      g.userData.limbs[0].rotation.x = stride;g.userData.limbs[1].rotation.x = -stride;
      g.userData.limbs[2].rotation.x = -stride;g.userData.limbs[3].rotation.x = stride;
    }
    for (const [id,g] of people) if(!visible.has(id)){peopleRoot.remove(g);people.delete(id);}
    diagnostic.npc = people.size;
  }
  function render(dt, view) {
    if (disposed) throw new Error('Three.js adapter was disposed');
    const rect = canvas.getBoundingClientRect();
    if (rect.width < 1 || rect.height < 1) return;
    const quality = window.LS3D_CONFIG?.quality || 'medium';
    const dpr = Math.min(window.devicePixelRatio || 1, quality === 'low' ? .8 : quality === 'high' ? 1.55 : 1.1);
    renderer.setPixelRatio(dpr);
    renderer.setSize(Math.floor(rect.width), Math.floor(rect.height), false);
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
    refreshAssets(flags);
    routeScene(flags);
    refreshPeople(flags);
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
      let obj=hit.object;
      while(obj && !obj.userData.assetId) obj=obj.parent;
      const id=obj?.userData.assetId, asset=T.assetMap.get(id);
      if(asset){T.pickAsset(asset);return id;}
    }
    return null;
  }
  function dispose() {
    if(disposed)return;disposed=true;
    clearGroup(routeRoot);clearGroup(zoneRoot);clearGroup(peopleRoot);clearGroup(assetRoot);
    for(const m of materialCache.values())m.dispose();
    materialCache.clear();boxGeometry.dispose();cylinderGeometry.dispose();
    renderer.dispose();
  }
  return { render, pick, dispose, diagnostic, renderer, scene, camera, THREE, get architecture(){return architecture.summary;}, get disposed(){return disposed;} };
}
