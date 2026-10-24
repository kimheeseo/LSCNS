// STEP 05 / 20 — distinguishing 3D equipment + geometry LOD and real GPU instancing.
// Original asset ids/types, positions, dimensions and BOM values remain authoritative.
// All sizes/shapes are illustrative; manufacturer CAD fidelity is NOT claimed.
export function createEquipmentFactory(THREE, { cube, cylinder, material, boxGeometry }) {
  const diagnostics = {created:0, instancedBatches:0, instancedParts:0, kinds:{}};
  const tmp = new THREE.Object3D();
  function classify(a) {
    const t = String(a.type || '').toLowerCase(), id=String(a.id||'').toLowerCase();
    const s=t+' '+id;
    if(/site-sign|ls data center.*안내/.test(s))return 'visitor-center';
    if(/visitor-lobby|briefing|classroom/.test(s))return 'briefing-room';
    if(/visitor-loading|하역장|loading bay/.test(s))return 'loading-bay';
    if(/cooling tower|\btower-[ab]\b/.test(s))return 'cooling-tower';
    if(/transformer|xfmr/.test(s))return 'transformer';
    if(/switchgear/.test(s))return 'switchgear';
    if(/generator/.test(s))return 'generator';
    if(/(^|\s)ats(\s|$)|automatic transfer/.test(s))return 'ats';
    if(/fuel tank|\bfuel\b/.test(s))return 'fuel-tank';
    if(/chiller/.test(s))return 'chiller';
    if(/water tank|water-storage|\bwater\b/.test(s))return 'water-tank';
    if(/pump|heat exchanger/.test(s))return 'pump-hx';
    if(/\bcdu\b|coolant distribution/.test(s))return 'cdu';
    if(/\bups\b/.test(s))return 'ups';
    if(/battery/.test(s))return 'battery-bank';
    if(/\bpdu\b/.test(s))return 'pdu';
    if(/busway|cable.tray/.test(s))return 'busway';
    if(/air handling|crah|crac/.test(s))return 'air-handler';
    if(/(optical entry|\bodf\b|cross-connect)/.test(s))return 'odf';
    if(/(spine|leaf|core|carrier|mmr)/.test(s))return 'network';
    if(/x-ray|xray|baggage scanner/.test(s))return 'xray';
    if(/metal.detector|walk-through detector/.test(s))return 'metal-detector';
    if(/security gate|출입 보안|(^|\s)gate(\s|$)/.test(s))return 'access-gate';
    if(/customer cage|colocation cage/.test(s))return 'customer-cage';
    if(/fire suppression/.test(s))return 'fire-suppression';
    if(/noc monitoring|security|desk|reception/.test(s))return 'control-desk';
    if(/maintenance/.test(s))return 'maintenance';
    return 'generic';
  }
  function part(root, name, x,y,z,w,h,d,color,metal=.38,rough=.5){
    const m=cube(root,x,y,z,w,h,d,color,metal,rough);
    m.name='equipment-part:'+name;
    return m;
  }
  function round(root,name,x,y,z,r,h,color){
    const m=cylinder(root,x,y,z,r,h,color);m.name='equipment-part:'+name;return m;
  }
  function batch(root,name,positions,w,h,d,color){
    if(!positions.length)return;
    const b=new THREE.InstancedMesh(boxGeometry,material(color,.68,.38),positions.length);
    b.name='equipment-instance:'+name;
    for(let i=0;i<positions.length;i++){
      const p=positions[i];
      tmp.position.set(p[0],p[1]+h/2,p[2]);tmp.scale.set(w,h,d);
      tmp.rotation.set(0,0,0);tmp.updateMatrix();
      b.setMatrixAt(i,tmp.matrix);
    }
    b.instanceMatrix.needsUpdate=true;
    b.castShadow=false;b.receiveShadow=true;
    b.frustumCulled=false; // per-batch bounding-sphere recompute avoided; per asset LOD gates visibility
    root.add(b);
    diagnostics.instancedBatches++;
    diagnostics.instancedParts+=positions.length;
  }
  function panel(root,name,w,h,d,color='#567a91'){
    part(root,name+'-case',0,.12,0,w*.92,h*.93,d*.9,color,.62,.37);
    part(root,name+'-door',0,h*.16,-d*.463,w*.8,h*.56,.06,'#1b374c',.41,.37);
    part(root,name+'-display',0,h*.61,-d*.52,w*.37,.24,.05,'#4bd6cb',.15,.25);
    batch(root,name+'-vents',Array.from({length:9},(_,i)=>
      [-w*.33+i*w*.083,h*.28,-d*.512]),w*.055,.045,.04,'#a6cbd7');
  }
  function electricalBus(root,w,d,h){
    part(root,'busway-conductor',0,h*.65,0,w*.94,.33,d*.34,'#b78e4a',.78,.26);
    part(root,'busway-shell',0,h*.46,0,w*.98,.33,d*.48,'#9eacb8',.52,.39);
    batch(root,'busway-hangers',Array.from({length:5},(_,i)=>
      [-w*.4+i*w*.20,h*.45,-d*.28]),.15,h*.36,.13,'#6f879a');
  }
  function detailed(root,a,kind) {
    const {w,d,h}=a;
    const front=-d*.48;
    switch(kind){
      case 'transformer': {
        part(root,'oil-filled-main-tank',0,.34,0,w*.86,h*.66,d*.82,'#718b9a',.61,.42);
        part(root,'protective-base',0,.05,0,w*.95,.28,d*.9,'#334c5e');
        const fins=Array.from({length:11},(_,i)=>[-w*.35+i*w*.07,h*.31,d*.447]);
        batch(root,'transformer-radiator-fins',fins,w*.038,h*.56,.10,'#b4c5c9');
        for(let i=-1;i<=1;i++){
          round(root,'porcelain-bushing-'+i,i*w*.23,h*.72,0,.18,h*.2,'#dedac4');
          round(root,'bushing-terminal-'+i,i*w*.23,h*.89,0,.09,.16,'#bd9864');
        }
        part(root,'transformer-top-cover',0,h*.74,0,w*.90,.19,d*.88,'#a9bdc4',.63,.33);
        break;
      }
      case 'switchgear': {
        part(root,'metalclad-main',0,.08,0,w*.98,h*.92,d*.86,'#748a96',.65,.38);
        const cubicles=5;
        for(let i=0;i<cubicles;i++){
          const x=-w*.39+i*w*.195;
          part(root,'switchgear-bay-'+i,x,h*.09,front,w*.18,h*.78,.065,'#9fb2bc',.59,.39);
          part(root,'breaker-window-'+i,x,h*.53,front-.055,w*.12,.34,.025,'#1b3e57',.14,.27);
          part(root,'breaker-lamp-'+i,x,h*.77,front-.073,.065,.095,.035,i%2?'#53d7aa':'#f3cd76',.1,.34);
        }
        break;
      }
      case 'generator': {
        part(root,'genset-skid',0,.12,0,w*.98,.33,d*.88,'#384958',.66,.46);
        part(root,'diesel-generator-housing',0,.55,0,w*.84,h*.61,d*.78,'#acb08e',.46,.53);
        part(root,'radiator-hood',-w*.35,.8,0,w*.16,h*.37,d*.73,'#3d6675',.52,.38);
        batch(root,'radiator-grilles',Array.from({length:11},(_,i)=>[
          -w*.44,.94+i*(h*.32/11),d*.32]),w*.06,.035,.042,'#263f53');
        round(root,'exhaust-pipe',w*.28,h*.65,d*.12,.19,h*.22,'#899ca6');
        part(root,'exhaust-weather-cap',w*.28,h*.89,d*.12,.55,.09,.55,'#d3b878');
        break;
      }
      case 'ups': {
        // Online UPS: distinct cabinet face, bypass column and battery/DC status bank.
        part(root,'ups-plinth',0,.03,0,w*.98,.18,d*.93,'#263948',.66,.38);
        part(root,'ups-main-cabinet',-w*.11,.15,0,w*.68,h*.78,d*.85,'#738897',.59,.35);
        part(root,'ups-door',-w*.11,h*.18,front,w*.57,h*.57,.065,'#193247',.35,.29);
        part(root,'ups-hmi-display',-w*.11,h*.61,front-.05,w*.29,.22,.025,'#58e1cd',.12,.18);
        part(root,'ups-bypass-column',w*.31,.14,0,w*.19,h*.76,d*.72,'#405d6e',.62,.38);
        part(root,'ups-input-output-strip',w*.31,h*.56,front-.38,w*.13,h*.15,.025,'#e4c66f',.2,.28);
        batch(root,'ups-status-leds',Array.from({length:8},(_,i)=>[
          -w*.34+(i%4)*w*.12,h*.82+Math.floor(i/4)*h*.07,front-.06]),w*.045,.04,.02,
          '#65e6b5');
        batch(root,'ups-vent-slots',Array.from({length:12},(_,i)=>[
          -w*.35+(i%6)*w*.085,h*.28+Math.floor(i/6)*h*.09,front-.07]),w*.05,.035,.025,
          '#9fb9c3');
        break;
      }
      case 'ats':case 'pdu':case 'cdu': {
        panel(root,kind,w,h,d,kind==='cdu'?'#4d8799':'#677e92');
        if(kind==='cdu'){
          for(const x of [-w*.25,w*.25]){
            round(root,'coolant-manifold-'+x,x,.03,d*.54,.14,h*.56,'#6abed0');
            round(root,'coolant-port-'+x,x,h*.66,d*.55,.19,.14,'#e3b76a');
          }
          batch(root,'cdu-indicators',Array.from({length:6},(_,i)=>[
            -w*.32+i*w*.13,h*.11,front-.08]),w*.065,.055,.045,'#5ee4d1');
        }
        break;
      }
      case 'battery-bank': {
        part(root,'battery-cabinet',0,.12,0,w*.97,h*.92,d*.85,'#718295',.51,.52);
        for(let i=0;i<4;i++) {
          const x=-w*.36+i*w*.245;
          part(root,'battery-drawer-'+i,x,h*.19,front,w*.20,h*.61,.065,'#c3cbd0',.46,.39);
        }
        batch(root,'battery-leds',Array.from({length:8},(_,i)=>[
          -w*.39+i*w*.11,h*.77,front-.08]),w*.055,.055,.028,'#83edbc');
        break;
      }
      case 'fuel-tank':case 'water-tank': {
        round(root,kind+'-cyl-tank',0,.18,0,Math.min(w,d)*.39,h*.85,kind==='fuel-tank'?'#8f9b7c':'#83b8c9');
        part(root,'tank-ring-platform',0,h*.52,0,w*.83,.15,d*.83,'#5d7685',.61,.44);
        round(root,'tank-vent',w*.18,h*.81,0,.13,.30,'#d9c7a3');
        break;
      }
      case 'cooling-tower':{
        part(root,'concrete-basin',0,.08,0,w*.98,h*.17,d*.98,'#365e73',.37,.55);
        part(root,'tower-filled-media',0,h*.17,0,w*.90,h*.53,d*.86,'#83bac9',.28,.59);
        batch(root,'louver-blades',Array.from({length:14},(_,i)=>[
          -w*.39+i*w*.06,h*.28,front]),w*.036,h*.38,.13,'#d5e8e7');
        round(root,'fan-housing',0,h*.72,0,Math.min(w,d)*.34,.30,'#427185');
        const vanes=[];
        for(let i=0;i<8;i++){
          const angle=i*Math.PI/4,rr=Math.min(w,d)*.19;
          vanes.push([Math.sin(angle)*rr,h*.88,Math.cos(angle)*rr]);
        }
        batch(root,'fan-blades',vanes,w*.13,.085,d*.15,'#e0f1f0');
        part(root,'fan-protective-grille',0,h*.95,0,w*.75,.07,d*.75,'#2c5367',.49,.53);
        break;
      }
      case 'chiller':{
        part(root,'chiller-chassis',0,.14,0,w*.95,h*.67,d*.85,'#699eb2',.52,.4);
        for(let i=-1;i<=1;i++)
          round(root,'compressor-can-'+i,i*w*.25,h*.69,0,Math.min(w,d)*.13,h*.19,'#c4d5da');
        batch(root,'condenser-coil-fins',Array.from({length:14},(_,i)=>[
          -w*.40+i*w*.062,h*.30,front]),w*.04,h*.34,.065,'#a1d2d7');
        for(const x of [-w*.28,w*.28])
          round(root,'fluid-pipe-'+x,x,h*.61,d*.49,.13,.29,'#d8b677');
        break;
      }
      case 'pump-hx':{
        part(root,'heat-exchanger-frame',-w*.18,.34,0,w*.49,h*.72,d*.65,'#84a3b1');
        batch(root,'heat-exchanger-plates',Array.from({length:13},(_,i)=>[
          -w*.35+i*w*.031,h*.56,d*.35]),w*.018,h*.43,.09,'#d1d9dd');
        for(const x of [w*.25,w*.40]){
          round(root,'circulation-pump-'+x,x,.20,-d*.14,Math.min(w,d)*.16,h*.45,'#4a9cbe');
          part(root,'pump-motor',x,h*.56,-d*.14,.7,.45,.7,'#4b7085');
        }
        break;
      }
      case 'air-handler':{
        part(root,'crac-chassis',0,.12,0,w*.92,h*.91,d*.88,'#92a8ae');
        part(root,'air-return-vent',0,.46,front,w*.82,h*.48,.08,'#23445b',.25,.72);
        batch(root,'louver-array',Array.from({length:14},(_,i)=>[
          0,.55+i*(h*.40/14),front-.06]),w*.74,.047,.045,'#a9c5cd');
        part(root,'supply-vent',0,h*.76,front,w*.80,.18,.10,'#4b8395');
        break;
      }
      case 'odf':case 'network':{
        part(root,'network-cabinet',0,.06,0,w*.9,h*.95,d*.87,'#3b5673',.65,.37);
        for(let i=0;i<5;i++) {
          const y=.72+i*(h*.69/5);
          part(root,'optical-switch-tray-'+i,0,y,front,w*.78,.25,.09,
            kind==='odf'?'#af994f':'#315f7d',.65,.4);
        }
        const led=[];
        for(let i=0;i<36;i++)led.push([-w*.34+(i%12)*w*.063,
          .86+Math.floor(i/12)*h*.23,front-.10]);
        batch(root,kind+'-ports',led,w*.034,.045,.025,'#70ecce');
        part(root,'cable-management-duct',w*.43,.25,front,w*.07,h*.67,.15,'#bac9d1',.5,.39);
        break;
      }
      case 'xray':{
        part(root,'baggage-scanner',0,.68,0,w*.54,h*.53,d*.82,'#637d91');
        for(const x of [-w*.32,w*.32]){
          part(root,'conveyor-wing-'+x,x,.66,0,w*.34,.22,d*.60,'#98a4aa',.4,.55);
          batch(root,'xray-rollers-'+x,Array.from({length:7},(_,i)=>[
            x-w*.13+i*w*.043,.75,0]),w*.023,.10,d*.57,'#4b6675');
        }
        part(root,'xray-shield-curtain',0,.86,d*.40,w*.45,h*.36,.13,'#182e43');
        break;
      }
      case 'metal-detector':case 'access-gate': {
        for (const xx of [-w*.34,w*.34]) part(root,'gate-post-'+xx,xx,.06,0,w*.13,h*.91,d*.42,'#adbfc9');
        part(root,'gate-lintel',0,h*.86,0,w*.88,.22,d*.44,'#74b5c4');
        part(root,'card-scanner',w*.41,h*.33,d*.22,.28,.44,.16,'#43d6b8');
        break;
      }
      case 'customer-cage':{
        for(const xx of [-w*.46,w*.46])for(const zz of [-d*.45,d*.45])
          part(root,'mesh-cage-post',xx,.06,zz,.16,h*.94,.16,'#889eac');
        for(const zz of [-d*.45,d*.45])
          part(root,'cage-rail-z',0,h*.78,zz,w*.92,.07,.1,'#7997a9');
        for(const xx of [-w*.46,w*.46])
          part(root,'cage-rail-x',xx,h*.78,0,.1,.07,d*.90,'#7997a9');
        break;
      }
      case 'fire-suppression':{
        for(let i=0;i<4;i++){
          const x=-w*.32+i*w*.21;
          round(root,'suppression-cylinder-'+i,x,.13,0,Math.min(w,d)*.12,h*.69,'#be6577');
          round(root,'valve-'+i,x,h*.81,0,.08,.20,'#cbb991');
        }
        part(root,'pipe-header',0,h*.86,0,w*.78,.12,.12,'#b8b9c4');
        break;
      }
      case 'control-desk':case 'maintenance':{
        part(root,'counter',0,.54,0,w*.86,.18,d*.72,'#7392a4');
        part(root,'desk-front',0,.06,-d*.28,w*.80,.66,.15,'#325c73');
        if(kind==='control-desk') for(const x of [-w*.24,0,w*.24]){
          part(root,'monitor-'+x,x,.87,-d*.12,w*.18,.58,.12,'#1d384e');
          part(root,'display-'+x,x,.94,-d*.20,w*.16,.37,.03,'#49b4ce',.1,.14);
        }
        break;
      }
      case 'visitor-center': {
        part(root,'visitor-center-building',0,.06,0,w*.96,h*.83,d*.90,'#71808b',.40,.50);
        part(root,'visitor-center-glass',0,h*.43,front,w*.72,h*.42,.055,'#254c62',.12,.23);
        part(root,'visitor-center-ls-plaque',0,h*.73,front-.06,w*.34,.22,.03,'#d9363e',.24,.29);
        part(root,'visitor-center-canopy',0,h*.84,front,w*.88,.12,d*.34,'#d8e3e8',.45,.35);
        break;
      }
      case 'briefing-room': {
        part(root,'briefing-floor',0,.04,0,w*.94,.12,d*.92,'#536f7f',.28,.60);
        part(root,'briefing-backwall',0,h*.28,d*.38,w*.90,h*.56,.08,'#e0e7e4',.12,.52);
        part(root,'briefing-screen',0,h*.57,d*.32,w*.43,h*.24,.035,'#2f6078',.1,.20);
        // Classroom desks and chairs face the briefing screen.
        for(const zz of [-d*.20,d*.03])for(const xx of [-w*.27,0,w*.27]){
          part(root,'briefing-desk',xx,.31,zz,w*.17,.08,d*.15,'#ad8f66',.25,.48);
          part(root,'briefing-chair',xx,.08,zz-d*.14,w*.12,.23,d*.10,'#364b58',.34,.48);
        }
        // Red-shirt instructor, intentionally simplified educational figure.
        round(root,'briefing-instructor-head',-w*.22,h*.66,d*.19,.11,.20,'#e1b28d');
        part(root,'briefing-instructor-red-shirt',-w*.22,h*.38,d*.19,.24,h*.27,.15,'#d74248',.20,.42);
        part(root,'briefing-podium',-w*.04,.18,d*.23,w*.15,h*.37,d*.14,'#4b7180',.35,.44);
        break;
      }
      case 'loading-bay': {
        part(root,'loading-bay-floor',0,.04,0,w*.95,.12,d*.94,'#536b78',.30,.58);
        part(root,'loading-truck-body',-w*.19,.18,0,w*.55,h*.48,d*.46,'#edf1ef',.30,.42);
        part(root,'loading-truck-cab',w*.19,.19,-d*.02,w*.20,h*.34,d*.43,'#b8c9d0',.34,.39);
        part(root,'loading-truck-open-tail',-w*.48,h*.28,d*.28,w*.09,h*.39,.035,'#d7e3e6',.30,.42);
        for(const xx of [-w*.33,w*.20])for(const zz of [-d*.22,d*.22])
          round(root,'loading-truck-wheel',xx,.08,zz,.12,.14,'#202d35');
        part(root,'loading-ls-logo',-w*.17,h*.49,-d*.25,w*.22,.13,.02,'#d9363e',.20,.30);
        // Forklift + branded delivery cartons beside open tailgate.
        part(root,'loading-forklift-body',w*.29,.10,d*.24,w*.22,h*.20,d*.18,'#e5a94d',.42,.36);
        part(root,'loading-forklift-mast',w*.39,.14,d*.24,.05,h*.46,.05,'#3e4f58',.52,.38);
        part(root,'loading-forklift-forks',w*.43,.06,d*.24,w*.22,.04,.09,'#697882',.50,.35);
        for(let i=0;i<3;i++)part(root,'loading-ls-box-'+i,-w*.42+i*w*.12,.12,d*.27,w*.10,h*.18,d*.13,'#c99b54',.18,.56);
        break;
      }
      case 'busway':electricalBus(root,w,d,h);break;
      default:{
        panel(root,'generic',w,h,d);
      }
    }
  }
  function build(a) {
    const kind=classify(a), w=Math.max(.3,+a.w||3), d=Math.max(.3,+a.d||3), h=Math.max(.4,+a.h||3);
    const dims={...a,w,d,h};
    const lod=new THREE.LOD();lod.name='equipment-lod:'+kind;
    // High detail for selected/near equipment; repeat subcomponents share instanced draw calls.
    const high=new THREE.Group();high.name='equipment-high:'+kind;
    detailed(high,dims,kind);
    const mid=new THREE.Group();mid.name='equipment-medium:'+kind;
    const shellColor=kind.includes('cool')||kind==='cdu'||kind==='chiller'?'#679eaf':
      /transformer|switchgear|generator|ups|battery|pdu|ats/.test(kind)?'#879aa3':'#507589';
    part(mid,'mid-shell',0,.08,0,w*.88,h*.84,d*.84,shellColor,.43,.52);
    part(mid,'mid-door',0,h*.24,-d*.44,w*.7,h*.43,.08,'#2e5269');
    const low=new THREE.Group();low.name='equipment-low:'+kind;
    part(low,'low-silhouette',0,.04,0,w*.89,h*.87,d*.85,shellColor,.22,.78);
    lod.addLevel(high,0);
    lod.addLevel(mid,65);
    lod.addLevel(low,155);
    diagnostics.created++;diagnostics.kinds[kind]=(diagnostics.kinds[kind]||0)+1;
    return {lod,kind};
  }
  return {build,classify,diagnostics};
}
