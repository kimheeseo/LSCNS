// LS3D STEP 04 — industrial architecture geometry for the real Three.js renderer.
// Builds spatial geometry: slabs, façades, doors, glass, ceiling structure and suspended trays.
// Illustrative coordinates (legacy scene units), NOT CAD/BIM or construction clearances.
export function describeArchitecturePart(kind,name,zoneId){
 if(kind!=='tray'&&kind!=='beam')return null;
 const component=name.includes('-hanger-')?'천장 행거':name.includes('-rung-')?'가로 지지대':name.includes('-rail-')?'측면 레일':name.endsWith('-base')?'바닥 받침':name.endsWith('-pick-surface')?'트레이 본체':'구조 부재';
 const power=name.includes('-power-tray'),fiber=name.includes('-fiber-tray');
 const group=kind==='tray'?name.replace(/-(?:base|rail-.*|rung-.*|hanger-.*|pick-surface)$/,''):name;
 const label=kind==='beam'?'천장 구조 보':power?'전력 케이블 트레이':fiber?'광케이블 트레이':name.includes('-cross-tray')?'교차 케이블 트레이':'설비 케이블 트레이';
 const description=kind==='beam'?'천장과 설비 지지 구조를 구성하는 보입니다. 전기 도체나 광케이블이 아닙니다.':power?'전력 케이블 트레이(Cable Tray)는 랙 상부에서 전력 케이블을 받치고 배선 경로를 정리하는 지지 설비입니다. 측면 레일·가로 지지대·바닥 받침과 천장 행거로 구성됩니다. 화면의 노란색은 전력 배선 경로를 구별하는 표시이며, 실제 케이블 자체를 뜻하지 않습니다. 전력은 트레이 위에 포설한 케이블을 통해 전달되며, 이 부품 자체가 전류를 공급하는 버스웨이는 아닙니다. 선정 시 케이블 중량에 따른 허용 하중, 지지 간격, 점유율, 접지·본딩 조건을 확인합니다. 교육용 개념 모델로 특정 LS전선 제품의 치수·정격을 나타내지는 않습니다.':fiber?'광케이블을 지지·정리하고 전력 배선과 경로를 구별하는 트레이입니다. 실제 포설 시 허용 굴곡반경과 충전율을 확인해야 합니다.':'구역 사이에서 케이블 경로를 연결하고 지지하는 트레이입니다. 실제 시공 치수와 내하중을 나타내는 모델은 아닙니다.';
 return {id:'structure:'+zoneId+':'+group,kind:kind==='tray'?'cable-tray':'structural-beam',label,description,component,zoneId};
}

export function createArchitectureKit(THREE, { cube }) {
  const doors = [];
  let counts = {}, lastMode = '', lastZone = '', lastFlags = {};
  const names = ['floor','wall','glass','door','beam','ceiling','roof','tray','frame','path','fence','detail'];
  function reset() { counts = Object.fromEntries(names.map(n => [n, 0])); doors.length = 0; }
  function part(parent, kind, name, x, y, z, w, h, d, color, metal = .33, rough = .48, transparent = false) {
    const mesh = cube(parent, x, y, z, w, h, d, color, metal, rough, transparent);
    mesh.name = 'arch:' + kind + ':' + name;
    mesh.userData.archKind = kind;
    mesh.userData.noPick = true;
    if(kind==='tray'||kind==='beam'){
      let owner=parent;while(owner&&!owner.userData?.zoneId)owner=owner.parent;
      if(owner){mesh.userData.inspectionPart=describeArchitecturePart(kind,name,owner.userData.zoneId);mesh.userData.noPick=false;}
    }
    mesh.castShadow = kind === 'wall' || kind === 'beam' || kind === 'door';
    counts[kind] = (counts[kind] || 0) + 1;
    return mesh;
  }
  function addSite(root, zones) {
    // Ground stack stays in the campus world only; server-room excludes this entire branch.
    const ground = new THREE.Group(); ground.name = 'arch:campus:ground'; root.add(ground);
    part(ground,'floor','site-foundation',0,-.68,0,142,.42,116,'#263f51',.12,.91);
    part(ground,'floor','asphalt',0,-.25,0,141,.13,114,'#3b5262',.11,.96);
    part(ground,'path','east-west-road',0,-.1,30,132,.055,6.3,'#273b4b',.06,.94);
    part(ground,'path','north-south-road',0,-.1,1,7.4,.055,101,'#2e4654',.06,.92);
    for (let x = -62; x <= 62; x += 9.5) {
      part(ground,'detail','road-stripe-'+x,x,-.03,30,3.9,.022,.11,'#d1b86d',.15,.78);
    }
    for (let z = -49; z <= 48; z += 10) {
      part(ground,'detail','road-stripe-z'+z,0,-.03,z,.10,.023,3.5,'#d1b86d',.15,.78);
    }
    for (const zone of zones) {
      if (zone.id === 'campus') continue;
      const [x,z]=zone.pos, [w,d]=zone.size;
      part(ground,'path',zone.id+'-apron',x,-.13,z,w+1.8,.065,d+1.8,'#617b88',.14,.88);
    }
    // Perimeter security boundary uses posts/rails with a real access opening on +z.
    for (let x=-68;x<=68;x+=8) {
      if (Math.abs(x)<7) continue;
      part(ground,'fence','perimeter-n-'+x,x,-.06,-54,.15,2.3,.16,'#71909c',.64,.37);
      part(ground,'fence','perimeter-s-'+x,x,-.06,54,.15,2.3,.16,'#71909c',.64,.37);
    }
    for(const z of [-54,54]) for(const x of [-37,37]) {
      part(ground,'frame','perimeter-marker-'+x+'-'+z,x,-.05,z,.32,2.4,.32,'#a6b9c1',.64,.34);
    }
    part(ground,'beam','gate-header',0,2.2,54,11,.46,.42,'#b8cbd0',.6,.3);
  }
  function floorAndAisles(g,z,w,d,mode) {
    const color=z.id==='hall'?'#4f6e7d':z.id==='gray'?'#526578':'#5e7a85';
    part(g,'floor','slab',0,-.24,0,w,.20,d,'#344d60',.23,.82);
    part(g,'floor','finished-surface',0,-.035,0,w-.35,.065,d-.35,color,.13,.74);
    // The regular raised-floor joints are meshes (not texture / floating SVG lines).
    const gridStep = z.id==='hall'?3.2:Math.max(3.4,w/9);
    for(let x=-w/2+gridStep; x<w/2-.2; x+=gridStep)
      part(g,'detail','tile-seam-x-'+x,x,.031,0,.028,.016,d-.7,'#8ba6b2',.17,.82);
    for(let zz=-d/2+gridStep; zz<d/2-.2; zz+=gridStep)
      part(g,'detail','tile-seam-z-'+zz,0,.031,zz,w-.7,.016,.029,'#8ba6b2',.17,.82);
    if(z.id==='hall'){
      for(let lane=-1;lane<=1;lane++){
        const xx=lane* Math.min(w*.22,9.8);
        part(g,'path','cold-hot-aisle-'+lane,xx,.04,0,2.1,.024,d*.78,lane===0?'#398da4':'#b07545',.07,.91);
        for(let zz=-d*.35;zz<=d*.35;zz+=4.5)
          part(g,'detail','aisle-stripe-'+lane+'-'+zz,xx,.068,zz,1.8,.015,.095,'#c7d9d9',.1,.83);
      }
    }
    if(mode==='server-room') {
      part(g,'path','server-room-cable-entry',-w*.31,.045,-d*.36,1.8,.038,4.2,'#d2b16f',.27,.68);
    }
  }
  function structuralColumn(g,name,x,z,height,color='#90aebc') {
    part(g,'frame',name+'-pillar',x,.03,z,.44,height,.44,color,.67,.35);
    part(g,'detail',name+'-shoe',x,.02,z,.72,.2,.72,'#455e70',.55,.48);
  }
  function glassGrid(g,frontName,w,d,y,height,cutaway,side=1) {
    const segments=Math.max(4,Math.min(9,Math.round(w/5.7)));
    const half=w/2,step=w/segments,front=side*d/2+side*.025;
    for(let i=0;i<=segments;i++) {
      const x=-half+i*step;
      part(g,'frame',frontName+'-mullion-'+i,x,y,front,.12,height,.17,'#9fc1ce',.7,.28);
    }
    part(g,'beam',frontName+'-header',0,y+height-.14,front,w,.25,.23,'#b4cad2',.75,.3);
    for(let i=0;i<segments;i++){
      const x=-half+(i+.5)*step;
      // Door portal stays a true void; glazing is deliberately thinner in cutaway.
      if(Math.abs(x)<step*.6)continue;
      if(!cutaway){
        part(g,'glass',frontName+'-pane-'+i,x,y+.23,front,step-.19,height-.42,.06,'#90c9de',.07,.16,true);
      }
      part(g,'frame',frontName+'-sill-'+i,x,y+.31,front,step-.18,.10,.21,'#618798',.6,.4);
    }
  }
  function hingedDoor(g,name,x,z,height,openToOutside) {
    const leafWidth=1.37;
    const leafHeight=Math.min(3.25,height-.28);
    const joints=[];
    for (const dir of [-1,1]) {
      const hinge=new THREE.Group();
      hinge.name='arch:door-hinge:'+name+(dir<0?':left':':right');
      hinge.userData.archKind='door';
      hinge.userData.noPick=true;
      hinge.position.set(x+dir*leafWidth, .08, z);
      g.add(hinge);
      // Pivot around outer door edge, with glass and handles as real meshes.
      part(hinge,'door',name+'-leaf-'+dir, -dir*leafWidth*.5,0,0,
        leafWidth,leafHeight,.095,'#557c91',.7,.3);
      part(hinge,'glass',name+'-leaf-glass-'+dir,-dir*leafWidth*.5,.42, .083,
        leafWidth*.78,leafHeight-.7,.035,'#a5dfef',.07,.12,true);
      part(hinge,'detail',name+'-handle-'+dir,-dir*leafWidth*.22,1.18,.15,
        .055,.52,.1,'#e7cf90',.77,.22);
      joints.push({hinge,dir});
    }
    part(g,'frame',name+'-lintel',x,leafHeight+.05,z,leafWidth*2.12,.29,.34,'#b5cbd0',.64,.27);
    part(g,'frame',name+'-access-panel',x+leafWidth*1.3,1.22,z+.20,.21,.38,.13,'#40ab9b',.45,.4);
    const rec={joints,phase:0,openToOutside,zone:lastZone};doors.push(rec);return rec;
  }
  function suspendedTray(g,name,x,y,z,length,alongX,tint) {
    const channel=alongX?{w:length,d:.8}:{w:.8,d:length};
    // A tight transparent pick surface spans the tray channel so a touch does
    // not have to land precisely on a 0.065-unit rail. It never hides geometry.
    const pickSurface=part(g,'tray',name+'-pick-surface',x,y,z,channel.w,.29,channel.d,tint);
    pickSurface.userData.inspectionPickSurface=true;
    pickSurface.castShadow=false;pickSurface.receiveShadow=false;
    if(pickSurface.material?.clone){
      pickSurface.material=pickSurface.material.clone();
      pickSurface.material.transparent=true;pickSurface.material.opacity=0;
      pickSurface.material.depthWrite=false;pickSurface.material.colorWrite=false;
    }
    // U-shaped suspended cable support: base, two rails, ladder rungs, hanging rods.
    part(g,'tray',name+'-base',x,y,z,channel.w,.08,channel.d,'#627e8a',.65,.34);
    if (alongX) {
      for(const dz of [-.38,.38])part(g,'tray',name+'-rail-'+dz,x,y,z+dz,length,.28,.065,tint,.62,.31);
      for(let i=0;i<=Math.ceil(length/2.4);i++) {
        const xx=x-length/2 + i*length/Math.ceil(length/2.4);
        part(g,'tray',name+'-rung-'+i,xx,y+.06,z,.09,.08,.72,'#c6d2d4',.7,.27);
      }
      for(const xx of [-length*.38,0,length*.38]){
        part(g,'tray',name+'-hanger-left-'+xx,x+xx,y+.22,z-.36,.07,.78,.07,'#b9c7c8',.66,.32);
        part(g,'tray',name+'-hanger-right-'+xx,x+xx,y+.22,z+.36,.07,.78,.07,'#b9c7c8',.66,.32);
      }
    } else {
      for(const dx of [-.38,.38])part(g,'tray',name+'-rail-'+dx,x+dx,y,z,.065,.28,length,tint,.62,.31);
      for(let i=0;i<=Math.ceil(length/2.4);i++) {
        const zz=z-length/2 + i*length/Math.ceil(length/2.4);
        part(g,'tray',name+'-rung-'+i,x,y+.06,zz,.72,.08,.09,'#c6d2d4',.7,.27);
      }
      for(const zz of [-length*.38,0,length*.38]){
        part(g,'tray',name+'-hanger-left-'+zz,x-.36,y+.22,z+zz,.07,.78,.07,'#b9c7c8',.66,.32);
        part(g,'tray',name+'-hanger-right-'+zz,x+.36,y+.22,z+zz,.07,.78,.07,'#b9c7c8',.66,.32);
      }
    }
  }
  function enclosedBuilding(root,z,mode,current,flags) {
    const [cx,cz]=z.pos, dims=mode==='server-room'?[26,24]:z.size;
    const [w,d]=dims, h=z.id==='hall'?9.5:z.id==='gray'?8.8:8.0;
    const local=new THREE.Group();local.name='arch:zone:'+z.id;local.userData.zoneId=z.id;
    local.position.set(cx,0,cz);root.add(local);
    const selected=current===z.id,section=!!flags.showSection,openView=section||selected;
    floorAndAisles(local,z,w,d,mode);
    const facade='#819eb2', mullion='#b0c6d0', trim='#304e64';
    const wallBase=openView?1.15:h*.93;
    const sideVisible=section?0:(selected?1.8:h);
    // Outer wall foundation survives section mode; vertical walls become transparent/open.
    for(const side of [-1,1]) {
      if(section)break;
      part(local,'wall','side-'+side,side*w/2,.12,0,.24,sideVisible,d,facade,.37,.46);
      part(local,'glass','side-glass-'+side,side*(w/2+.13),1.45,0,.065,
        Math.max(.2,sideVisible-1.65),d*.52,'#81c0d6',.04,.2,true);
    }
    if(!section){
      // Campus camera approaches Gray Zone and Network/MMR from +Z, so their
      // entrance façades remain on that visible side while equipment faces rotate.
      const frontSide=1, rearSide=-frontSide;
      part(local,'wall','rear-wall',0,.1,rearSide*d/2,w,selected?1.7:h,.26,facade,.28,.55);
      if(!openView)glassGrid(local,z.id+'-rear',w,d,2.3,h-2.5,false,rearSide);
      // A meaningful open doorway, not a solid front wall overlapping the entrance.
      const portal=3.3, wing=Math.max(1,w/2-portal/2);
      const frontZ=frontSide*d/2;
      for(const side of [-1,1]){
        part(local,'wall','front-wing-'+side,side*(portal/2+wing/2),.1,frontZ,wing,wallBase,.24,facade,.38,.46);
        structuralColumn(local,'front-jamb-'+side,side*portal/2,frontZ,h,mullion);
      }
      if(!openView)glassGrid(local,z.id+'-facade',w,d,3.32,h-3.45,false,frontSide);
      else part(local,'beam','front-portal-header',0,3.24,frontZ,4.1,.38,.5,trim,.64,.32);
      hingedDoor(local,z.id+'-entry',0,frontZ+frontSide*.19,h,frontSide>0);
    } else {
      // Cut-away: keep real frame + roof skeleton instead of deleting the whole building.
      for(const sx of [-1,1])for(const sz of [-1,1])
        structuralColumn(local,'cut-column-'+sx+'-'+sz,sx*w*.49,sz*d*.48,h,mullion);
    }
    // Structural beams over IT aisles (visibly above 6-unit-tall rack concept meshes).
    for(const sx of [-1,1]) {
      part(local,'beam','primary-beam-z-'+sx,sx*w*.41,h-.45,0,.27,.38,d-.5,'#c6cdd0',.68,.31);
      part(local,'beam','cross-beam-x-'+sx,0,h-.45,sx*d*.29,w-.5,.34,.30,'#b8cad0',.66,.3);
    }
    // Suspended ceiling is a set of real acoustic panels along the perimeter,
    // leaving the room center open to the overhead default camera.
    for(const sx of [-1,1]) {
      part(local,'ceiling','perimeter-ceiling-'+sx,sx*w*.375,h-.21,0,w*.24,.12,d*.85,'#a7c1c9',.3,.65);
      for(let j=-2;j<=2;j++) part(local,'ceiling','light-strip-'+sx+'-'+j,
        sx*w*.35,h-.34,j*d*.15,w*.16,.07,.12,'#bde5ed',.12,.25);
    }
    const trayY=h-1.62;
    suspendedTray(local,z.id+'-power-tray',-w*.29,trayY,0,d*.78,false,'#e4b867');
    suspendedTray(local,z.id+'-fiber-tray',w*.29,trayY,0,d*.78,false,'#4dccbc');
    suspendedTray(local,z.id+'-cross-tray',0,trayY-.31,-d*.22,w*.68,true,'#89b5d4');
    // Roof is optional and responds to the existing Roof/Section buttons.
    if(flags.showRoof&&!section) {
      part(local,'roof','roof-slab',0,h+.02,0,w+.7,.30,d+.7,'#779aab',.57,.33);
      part(local,'roof','clerestory',0,h+.34,0,w*.44,.13,d*.20,'#9ac4d2',.35,.34);
      for(let i=-2;i<=2;i++)part(local,'detail','roof-vent-'+i,i*w*.15,h+.47,-d*.29,w*.12,.14,.7,'#466b80',.67,.28);
    } else {
      // Partial ceiling can still be inspected under the open roof.
      part(local,'ceiling','overhead-grille',0,h-.3,-d*.42,w*.58,.09,d*.09,'#718f9d',.53,.35);
    }
    const entrySide=1;
    part(local,'path','entry-ramp',0,-.055,entrySide*d*.55,4.4,.09,3.5,'#87a6b0',.24,.73);
    part(local,'detail','entry-threshold',0,.05,entrySide*d*.503,3.45,.06,.24,'#f0cb79',.55,.31);
  }
  function outdoorYard(root,z,current,flags) {
    const [w,d]=z.size,[cx,cz]=z.pos;
    const isCooling=z.id==='cooling';
    const h=isCooling?7.7:6.6;
    const local=new THREE.Group();local.name='arch:zone:'+z.id;local.userData.zoneId=z.id;
    local.position.set(cx,0,cz);root.add(local);
    floorAndAisles(local,z,w,d,'outdoor');
    // Yard remains OPEN, physically distinct from an enclosed white-space data hall.
    for(const sx of [-1,1])for(let j=-2;j<=2;j++){
      const zz=j*d*.20;
      structuralColumn(local,'yard-fence-post-x'+sx+'-z'+j,sx*w*.49,zz,isCooling?3.0:2.6,'#7697a2');
    }
    for(const sz of [-1,1])for(let j=-3;j<=3;j++){
      const xx=j*w*.137;
      if(sz===1 && Math.abs(xx)<3.3)continue; // real access opening
      structuralColumn(local,'yard-fence-post-z'+sz+'-x'+j,xx,sz*d*.49,2.6,'#7697a2');
    }
    for(const sz of [-1,1]) {
      if(sz===1) {
        for(const sx of [-1,1])part(local,'fence','front-guardrail-'+sx,sx*w*.3,1.06,sz*d*.49,w*.37,.08,.12,'#a6b9c3',.67,.31);
      } else part(local,'fence','rear-guardrail',0,1.06,sz*d*.49,w,.08,.12,'#a6b9c3',.67,.31);
    }
    for(const sx of [-1,1]) {
      part(local,'fence','side-guardrail-'+sx,sx*w*.49,1.06,0,.12,.08,d,'#a6b9c3',.67,.31);
      for(let j=-2;j<=2;j++)part(local,'glass','open-mesh-panel-'+sx+'-'+j,
        sx*w*.49,.27,j*d*.20,.055,1.42,d*.19,'#7798a5',.05,.55,true);
    }
    // A real high-level cantilever canopy/gantry, NOT a fake solid building.
    for(const sx of [-1,1])for(const sz of [-1,1])
      structuralColumn(local,'gantry-'+sx+'-'+sz,sx*w*.41,sz*d*.40,h,'#b5c6ca');
    for(const sx of [-1,1])part(local,'beam','gantry-long-'+sx,sx*w*.41,h-.37,0,.27,.48,d*.86,'#a5beca',.64,.38);
    for(const sz of [-1,1])part(local,'beam','gantry-cross-'+sz,0,h-.37,sz*d*.4,w*.84,.37,.34,'#a5beca',.64,.38);
    suspendedTray(local,z.id+'-service-tray',0,h-1.12,-d*.32,w*.72,true,isCooling?'#54c8e8':'#e9bb69');
    if(flags.showRoof && !flags.showSection && isCooling)
      part(local,'roof','partial-cooling-canopy',0,h+.04,-d*.34,w*.88,.19,d*.24,'#83a5b3',.62,.3);
    part(local,'path','yard-safety-access',0,.04,d*.37,4.5,.025,3.8,isCooling?'#6dbad5':'#d5b16a',.14,.76);
  }
  function build(root, zones, mode, current, flags = {}) {
    reset();
    lastMode=mode;lastZone=current;lastFlags={showRoof:!!flags.showRoof,showSection:!!flags.showSection};
    const displayZones=zones.filter(z=>z.id!=='campus' && (current==='campus'||z.id===current));
    if(mode!=='server-room')addSite(root,displayZones);
    for(const z of displayZones) {
      if(z.id==='utility'||z.id==='cooling')outdoorYard(root,z,current,flags);
      else enclosedBuilding(root,z,mode,current,flags);
    }
    return {zones:displayZones.length+(mode==='server-room'?0:1),doors:doors.length,counts:{...counts}};
  }
  function tick(dt,current,flags={}) {
    for (const door of doors) {
      // Automatic opening when entering a section/zone; keeps the main input API unchanged.
      const target=(current!=='campus'||flags.showSection)?1:0;
      door.phase+=(target-door.phase)*(1-Math.exp(-Math.max(0,Math.min(.10,dt))*4.7));
      for(const {hinge,dir} of door.joints)
        hinge.rotation.y=dir * (door.openToOutside?1:-1) * door.phase * 1.17;
    }
  }
  return {build,tick,get summary(){return {mode:lastMode,zone:lastZone,flags:{...lastFlags},doors:doors.length,counts:{...counts}};}};
}
