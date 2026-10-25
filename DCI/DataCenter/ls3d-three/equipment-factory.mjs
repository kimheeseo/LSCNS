// STEP 05 / 20 — distinguishing 3D equipment + geometry LOD and real GPU instancing.
// Original asset ids/types, positions, dimensions and BOM values remain authoritative.
// All sizes/shapes are illustrative; manufacturer CAD fidelity is NOT claimed.
export function createEquipmentFactory(THREE, { cube, cylinder, material, boxGeometry }) {
  const diagnostics = {created:0, instancedBatches:0, instancedParts:0, kinds:{}};
  const tmp = new THREE.Object3D();
  function classify(a) {
    const t = String(a.type || '').toLowerCase(), id=String(a.id||'').toLowerCase();
    const s=t+' '+id;
    if(/visitor-parking|visitor parking|pre-registration parking/.test(s))return 'visitor-parking';
    if(/site-sign|ls data center.*안내/.test(s))return 'visitor-center';
    if(/visitor-lobby|briefing|classroom/.test(s))return 'briefing-room';
    if(/visitor-loading|하역장|loading bay/.test(s))return 'loading-bay';
    if(/visitor-locker|camera.*usb|laptop.*register/.test(s))return 'identity-camera';
    if(/visitor-mantrap|mantrap|iris|fingerprint/.test(s))return 'biometric-mantrap';
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
  function decal(root,name,text,x,y,z,w,h,background='#153344',foreground='#f2ffff'){
    if(typeof document==='undefined')return null;
    const canvas=document.createElement('canvas');canvas.width=512;canvas.height=192;
    const ctx=canvas.getContext('2d');if(!ctx)return null;
    ctx.fillStyle=background;ctx.fillRect(0,0,canvas.width,canvas.height);
    ctx.strokeStyle='#d8f7f2';ctx.lineWidth=9;ctx.strokeRect(5,5,502,182);
    ctx.fillStyle=foreground;ctx.textAlign='center';ctx.textBaseline='middle';
    let size=76;ctx.font='700 '+size+'px system-ui,sans-serif';
    while(ctx.measureText(text).width>450&&size>28){size-=2;ctx.font='700 '+size+'px system-ui,sans-serif';}
    ctx.fillText(text,256,98);
    const texture=new THREE.CanvasTexture(canvas);texture.colorSpace=THREE.SRGBColorSpace;
    const mesh=new THREE.Mesh(new THREE.PlaneGeometry(w,h),new THREE.MeshBasicMaterial({map:texture,transparent:true,depthWrite:false,side:THREE.DoubleSide}));
    mesh.name='equipment-decal:'+name;mesh.position.set(x,y,z);root.add(mesh);return mesh;
  }
  function kepcoLogoPanel(root,name,x,y,z,w,h,rotationY=0){
    if(typeof document==='undefined')return null;
    const canvas=document.createElement('canvas');canvas.width=900;canvas.height=420;
    const ctx=canvas.getContext('2d');if(!ctx)return null;
    ctx.fillStyle='#ffffff';ctx.fillRect(0,0,900,420);
    const cx=205,cy=172,r=132;
    ctx.fillStyle='#ef1d27';ctx.beginPath();ctx.arc(cx,cy,r,0,Math.PI*2);ctx.fill();
    // Red KEPCO sun-disc with diagonal white energy bands.
    ctx.save();ctx.beginPath();ctx.arc(cx,cy,r,0,Math.PI*2);ctx.clip();
    ctx.strokeStyle='#ffffff';ctx.lineWidth=22;
    for(let o=-260;o<=260;o+=45){ctx.beginPath();ctx.moveTo(cx-210,cy+o);ctx.lineTo(cx+210,cy+o-180);ctx.stroke();}
    ctx.restore();
    ctx.fillStyle='#164fa5';ctx.font='800 120px Arial,Helvetica,sans-serif';ctx.textAlign='center';ctx.textBaseline='middle';
    ctx.fillText('kepco',560,300);
    const texture=new THREE.CanvasTexture(canvas);texture.colorSpace=THREE.SRGBColorSpace;
    const mesh=new THREE.Mesh(new THREE.PlaneGeometry(w,h),new THREE.MeshBasicMaterial({map:texture,side:THREE.DoubleSide,transparent:true,depthWrite:false}));
    mesh.name='equipment-kepco-logo:'+name;mesh.position.set(x,y,z);mesh.rotation.y=rotationY;root.add(mesh);return mesh;
  }
  function schneiderLogoPanel(root,name,x,y,z,w,h,rotationY=0){
    if(typeof document==='undefined')return null;
    const canvas=document.createElement('canvas');canvas.width=1000;canvas.height=390;
    const ctx=canvas.getContext('2d');if(!ctx)return null;
    ctx.fillStyle='#ffffff';ctx.fillRect(0,0,1000,390);
    ctx.fillStyle='#009b4c';ctx.textAlign='center';ctx.textBaseline='middle';
    ctx.font='800 138px Arial,Helvetica,sans-serif';ctx.fillText('Schneider',530,130);
    ctx.beginPath();ctx.arc(320,270,42,0,Math.PI*2);ctx.strokeStyle='#009b4c';ctx.lineWidth=16;ctx.stroke();
    ctx.font='800 84px Arial,Helvetica,sans-serif';ctx.fillText('Electric',590,278);
    ctx.strokeStyle='#ffffff';ctx.lineWidth=12;ctx.beginPath();ctx.moveTo(292,270);ctx.lineTo(348,270);ctx.moveTo(320,242);ctx.lineTo(320,298);ctx.stroke();
    const texture=new THREE.CanvasTexture(canvas);texture.colorSpace=THREE.SRGBColorSpace;
    const mesh=new THREE.Mesh(new THREE.PlaneGeometry(w,h),new THREE.MeshBasicMaterial({map:texture,side:THREE.DoubleSide,transparent:true,depthWrite:false}));
    mesh.name='equipment-schneider-logo:'+name;mesh.position.set(x,y,z);mesh.rotation.y=rotationY;root.add(mesh);return mesh;
  }
  function statusPanel(root,name,title,value,x,y,z,w,h,rotationY=0){
    if(typeof document==='undefined')return null;
    const canvas=document.createElement('canvas');canvas.width=640;canvas.height=300;
    const ctx=canvas.getContext('2d');if(!ctx)return null;
    ctx.fillStyle='#071b26';ctx.fillRect(0,0,640,300);
    ctx.strokeStyle='#62e8d9';ctx.lineWidth=12;ctx.strokeRect(8,8,624,284);
    ctx.fillStyle='#b9f9f1';ctx.textAlign='center';
    let titleSize=68;ctx.font='700 '+titleSize+'px system-ui,sans-serif';
    while(ctx.measureText(title).width>570&&titleSize>24){titleSize-=2;ctx.font='700 '+titleSize+'px system-ui,sans-serif';}
    ctx.fillText(title,320,100);ctx.fillStyle='#67f6c9';
    let valueSize=82;ctx.font='700 '+valueSize+'px ui-monospace,monospace';
    while(ctx.measureText(value).width>570&&valueSize>24){valueSize-=2;ctx.font='700 '+valueSize+'px ui-monospace,monospace';}
    ctx.fillText(value,320,220);
    const texture=new THREE.CanvasTexture(canvas);texture.colorSpace=THREE.SRGBColorSpace;
    const mesh=new THREE.Mesh(new THREE.PlaneGeometry(w,h),new THREE.MeshBasicMaterial({map:texture,side:THREE.DoubleSide,transparent:true,depthWrite:false}));
    mesh.name='equipment-status:'+name;mesh.position.set(x,y,z);mesh.rotation.y=rotationY;root.add(mesh);return mesh;
  }
  function batteryBrandPanel(root,name,x,y,z,w,h,rotationY=0){
    if(typeof document==='undefined')return null;
    // Half-size centred Samsung SDI logo plate: smaller overall mark, but legible wordmark.
    const canvas=document.createElement('canvas');canvas.width=1200;canvas.height=300;
    const ctx=canvas.getContext('2d');if(!ctx)return null;
    ctx.fillStyle='#ffffff';ctx.fillRect(0,0,1200,300);
    ctx.strokeStyle='#d6dde3';ctx.lineWidth=8;ctx.strokeRect(4,4,1192,292);
    ctx.fillStyle='#1d32a5';ctx.font='800 102px Arial,Helvetica,sans-serif';
    ctx.letterSpacing='1px';ctx.textAlign='center';ctx.textBaseline='middle';
    ctx.fillText('SAMSUNG SDI',600,158);
    const texture=new THREE.CanvasTexture(canvas);texture.colorSpace=THREE.SRGBColorSpace;
    const mesh=new THREE.Mesh(new THREE.PlaneGeometry(w,h),new THREE.MeshBasicMaterial({map:texture,side:THREE.DoubleSide,transparent:true,depthWrite:false}));
    mesh.name='equipment-brand:'+name;mesh.position.set(x,y,z);mesh.rotation.y=rotationY;root.add(mesh);return mesh;
  }
  function shippingBrandPanel(root,name,brand,x,y,z,w,h){
    if(typeof document==='undefined')return null;
    const canvas=document.createElement('canvas');canvas.width=900;canvas.height=360;
    const ctx=canvas.getContext('2d');if(!ctx)return null;
    ctx.fillStyle='#d3a35b';ctx.fillRect(0,0,900,360);
    ctx.fillStyle='rgba(108,70,23,.16)';for(let x=-360;x<980;x+=52){ctx.save();ctx.translate(x,0);ctx.rotate(-.23);ctx.fillRect(0,0,21,650);ctx.restore();}
    ctx.strokeStyle='#6e451d';ctx.lineWidth=14;ctx.strokeRect(7,7,886,346);
    ctx.fillStyle='#5c3817';ctx.font='700 30px Arial,Helvetica,sans-serif';ctx.textAlign='center';ctx.textBaseline='middle';ctx.fillText('DATA CENTER EQUIPMENT  ·  HANDLE WITH CARE',450,38);
    ctx.font='900 44px Arial,Helvetica,sans-serif';ctx.fillText('↑   ↑',750,310);
    // Fill the carton label so each company's emblem remains legible from the tour camera.
    const key=String(brand).toLowerCase();ctx.textAlign='center';ctx.textBaseline='middle';
    if(key==='cisco'){
      ctx.fillStyle='#1695d2';for(let i=0;i<8;i++){const bh=58+(i%4)*34;ctx.fillRect(65+i*31,150-bh/2,19,bh);}
      ctx.font='800 128px Arial,Helvetica,sans-serif';ctx.fillText('Cisco',575,184);
    }else if(key==='nvidia'){
      ctx.fillStyle='#76b900';ctx.fillRect(45,55,250,250);ctx.strokeStyle='#fff';ctx.lineWidth=19;
      ctx.beginPath();ctx.ellipse(170,180,80,60,0,0,Math.PI*2);ctx.stroke();
      ctx.fillStyle='#fff';ctx.beginPath();ctx.arc(170,180,18,0,Math.PI*2);ctx.fill();
      ctx.fillStyle='#76b900';ctx.font='800 112px Arial,Helvetica,sans-serif';ctx.fillText('NVIDIA',595,182);
    }else{
      // LS Cable & System CI: blue LS wordmark with red arrow accent and gray English logotype.
      ctx.fillStyle='#164fa5';ctx.font='900 172px Arial,Helvetica,sans-serif';ctx.fillText('LS',175,196);
      ctx.fillStyle='#e8232d';ctx.beginPath();ctx.moveTo(230,49);ctx.lineTo(281,49);ctx.lineTo(256,82);ctx.closePath();ctx.fill();
      ctx.fillStyle='#68767e';ctx.font='700 72px Arial,Helvetica,sans-serif';ctx.fillText('Cable & System',590,190);
    }
    const texture=new THREE.CanvasTexture(canvas);texture.colorSpace=THREE.SRGBColorSpace;
    const mesh=new THREE.Mesh(new THREE.PlaneGeometry(w,h),new THREE.MeshBasicMaterial({map:texture,side:THREE.DoubleSide,transparent:true,depthWrite:false}));
    mesh.name='equipment-shipping-brand:'+name;mesh.position.set(x,y,z);root.add(mesh);return mesh;
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
  function nocOperationsPanel(root,name,mode,x,y,z,w,h,rotationY=0){
    if(typeof document==='undefined')return null;
    const canvas=document.createElement('canvas');canvas.width=700;canvas.height=400;
    const ctx=canvas.getContext('2d');if(!ctx)return null;
    ctx.fillStyle='#07131d';ctx.fillRect(0,0,700,400);
    ctx.strokeStyle='#4c7891';ctx.lineWidth=8;ctx.strokeRect(5,5,690,390);
    ctx.font='800 45px Arial,Helvetica,sans-serif';ctx.textAlign='left';ctx.textBaseline='middle';
    if(mode==='cctv'){
      ctx.fillStyle='#8eeaf0';ctx.fillText('CCTV · LIVE',36,42);
      for(let r=0;r<2;r++)for(let c=0;c<2;c++){
        const px=35+c*330,py=82+r*145;ctx.fillStyle='#163246';ctx.fillRect(px,py,300,115);
        ctx.strokeStyle='#6fd9e4';ctx.lineWidth=4;ctx.strokeRect(px,py,300,115);
        ctx.fillStyle='#4b9dc2';ctx.fillRect(px+15,py+25,120,54);
        ctx.fillStyle='#9fd7d2';ctx.fillRect(px+155,py+33,105,16);
        ctx.fillStyle='#52e9a3';ctx.beginPath();ctx.arc(px+274,py+22,7,0,Math.PI*2);ctx.fill();
      }
    }else{
      ctx.fillStyle='#72f0bf';ctx.fillText('BMS · LIVE',36,42);
      ctx.strokeStyle='#24485b';ctx.lineWidth=2;
      for(let row=0;row<4;row++){ctx.beginPath();ctx.moveTo(35,86+row*60);ctx.lineTo(665,86+row*60);ctx.stroke();}
      ctx.strokeStyle='#5af0bc';ctx.lineWidth=7;ctx.beginPath();ctx.moveTo(45,270);ctx.lineTo(150,225);ctx.lineTo(250,250);ctx.lineTo(360,150);ctx.lineTo(470,195);ctx.lineTo(650,110);ctx.stroke();
      ctx.fillStyle='#e6bf56';ctx.fillRect(50,300,80,54);ctx.fillRect(160,275,80,79);ctx.fillRect(270,230,80,124);
      ctx.fillStyle='#73cff0';ctx.font='700 31px ui-monospace,monospace';ctx.fillText('PUE 1.34   IT 1.219MW   23.4°C',38,382);
    }
    const texture=new THREE.CanvasTexture(canvas);texture.colorSpace=THREE.SRGBColorSpace;
    const mesh=new THREE.Mesh(new THREE.PlaneGeometry(w,h),new THREE.MeshBasicMaterial({map:texture,side:THREE.DoubleSide,transparent:true,depthWrite:false}));
    mesh.name='equipment-noc-display:'+name;mesh.position.set(x,y,z);mesh.rotation.y=rotationY;root.add(mesh);return mesh;
  }
  function detailed(root,a,kind) {
    const {w,d,h}=a;
    const front=-d*.48;
    switch(kind){
      case 'transformer': {
        part(root,'oil-filled-main-tank',0,.34,0,w*.86,h*.66,d*.82,'#718b9a',.61,.42);
        part(root,'protective-base',0,.05,0,w*.95,.28,d*.9,'#334c5e');
        // Twin radiator banks make the transformer recognisable from front and rear.
        for(const side of [-1,1]){
          const face=side*d*.447;
          const fins=Array.from({length:11},(_,i)=>[-w*.35+i*w*.07,h*.31,face]);
          batch(root,'transformer-radiator-fins-'+side,fins,w*.038,h*.56,.10,'#b4c5c9');
          part(root,'transformer-radiator-header-top-'+side,0,h*.62,face,w*.78,.065,.085,'#899fa9',.52,.36);
          part(root,'transformer-radiator-header-bottom-'+side,0,h*.12,face,w*.78,.065,.085,'#899fa9',.52,.36);
        }
        for(let i=-1;i<=1;i++){
          round(root,'porcelain-bushing-'+i,i*w*.23,h*.72,0,.18,h*.2,'#dedac4');
          round(root,'bushing-terminal-'+i,i*w*.23,h*.89,0,.09,.16,'#bd9864');
        }
        part(root,'transformer-top-cover',0,h*.74,0,w*.90,.19,d*.88,'#a9bdc4',.63,.33);
        for(const side of [-1,1])decal(root,'transformer-rating-'+side,'MAIN TRANSFORMER',0,h*.42,side*d*.505,w*.42,h*.085,'#3e5a67','#edf8f8',side<0?Math.PI:0);
        break;
      }
      case 'switchgear': {
        part(root,'metalclad-main',0,.08,0,w*.98,h*.92,d*.86,'#748a96',.65,.38);
        part(root,'switchgear-plinth',0,.035,0,w*1.02,.13,d*.92,'#2d4656',.62,.36);
        // Service face is intentionally on +Z to align the Utility Yard approach.
        const operatingFace=-front, cubicles=5;
        for(let i=0;i<cubicles;i++){
          const x=-w*.39+i*w*.195;
          part(root,'switchgear-bay-'+i,x,h*.09,operatingFace,w*.18,h*.78,.065,'#9fb2bc',.59,.39);
          part(root,'breaker-window-'+i,x,h*.43,operatingFace+.055,w*.12,.25,.025,'#102c40',.14,.27);
          part(root,'breaker-handle-'+i,x,h*.49,operatingFace+.083,w*.036,.115,.018,i===0?'#e4b54b':'#e7ecee',.25,.30);
          part(root,'breaker-lamp-'+i,x,h*.72,operatingFace+.073,.065,.075,.035,i%2?'#53d7aa':'#f3cd76',.1,.34);
          part(root,'feeder-tag-'+i,x,h*.25,operatingFace+.078,w*.105,.055,.018,'#365166',.14,.30);
        }
        // Main busbars, metering relay and five individual outgoing feeder bays.
        for(const [name,y,color] of [['r',h*.83,'#d65d59'],['s',h*.78,'#e2c257'],['t',h*.73,'#4c9cba']])
          part(root,'switchgear-main-bus-'+name,0,y,operatingFace+.078,w*.73,.032,.028,color,.45,.25);
        statusPanel(root,'switchgear-meter','LV SWITCHGEAR','NORMAL',0,h*.87,operatingFace+.091,w*.30,h*.105,0);
        decal(root,'switchgear-purpose','PROTECTION · DISTRIBUTION',0,h*.12,operatingFace+.087,w*.48,h*.070,'#244357','#e5fbfa');
        break;
      }
      case 'generator': {
        part(root,'genset-skid',0,.12,0,w*.98,.33,d*.88,'#384958',.66,.46);
        part(root,'diesel-generator-housing',0,.55,0,w*.84,h*.61,d*.78,'#7f8788',.46,.48);
        // Outdoor container genset: elevated radiator tower, exhaust stack and service ducting.
        part(root,'generator-radiator-tower',-w*.31,h*.16,0,w*.23,h*.78,d*.72,'#68747a',.44,.42);
        part(root,'generator-radiator-head',-w*.31,h*.93,0,w*.29,h*.20,d*.78,'#768288',.42,.38);
        round(root,'generator-exhaust-stack',w*.12,h*.58,d*.10,Math.min(w,d)*.10,h*.56,'#b6bec0');
        round(root,'generator-exhaust-cap',w*.12,h*1.13,d*.10,Math.min(w,d)*.15,.10,'#273f4c');
        part(root,'generator-exhaust-duct',w*.28,h*.58,d*.10,w*.26,.16,d*.20,'#aeb8ba',.44,.34);
        for(const side of [-1,1])batch(root,'generator-louvers-'+side,Array.from({length:10},(_,i)=>[
          w*.08,h*.24+i*h*.052,side*d*.405]),w*.54,.022,.018,'#263e49');
        part(root,'generator-service-door',w*.08,.54,front,w*.42,h*.43,.04,'#566c65',.42,.40);
        part(root,'generator-control-panel',w*.13,h*.63,front-.052,w*.15,.16,.016,'#58d8bd',.08,.20);
        part(root,'radiator-hood',-w*.35,.8,0,w*.16,h*.37,d*.73,'#3d6675',.52,.38);
        batch(root,'radiator-grilles',Array.from({length:11},(_,i)=>[
          -w*.44,.94+i*(h*.32/11),d*.32]),w*.06,.035,.042,'#263f53');
        batch(root,'generator-louvers',Array.from({length:9},(_,i)=>[
          w*.10,h*.25+i*h*.055,front-.06]),w*.50,.022,.018,'#334e58');
        round(root,'exhaust-pipe',w*.28,h*.65,d*.12,.19,h*.22,'#899ca6');
        part(root,'exhaust-weather-cap',w*.28,h*.89,d*.12,.55,.09,.55,'#d3b878');
        part(root,'generator-cable-termination',w*.37,.18,-d*.33,w*.14,h*.32,.10,'#293d47',.40,.34);
        // KEPCO grid mark on both service faces, matching the supplied emblem rather than plain text.
        // Compact utility mark: retain identification without dominating the generator louvre panel.
        kepcoLogoPanel(root,'generator-grid-front',-w*.10,h*.64,front-.082,w*.215,h*.12,Math.PI);
        kepcoLogoPanel(root,'generator-grid-rear',-w*.10,h*.64,-front+.082,w*.215,h*.12,0);
        break;
      }
      case 'ups': {
        // White-and-black modular UPS cabinet inspired by the supplied Schneider Galaxy-style reference:
        // black fan bays, white power modules, a green control spine and central HMI.
        part(root,'ups-plinth',0,.03,0,w*.98,.16,d*.93,'#151a1d',.66,.38);
        part(root,'ups-white-chassis',0,.14,0,w*.94,h*.80,d*.86,'#e5e8e7',.50,.30);
        part(root,'ups-left-fan-bay',-w*.36,h*.14,front,w*.17,h*.66,.070,'#171d20',.42,.28);
        part(root,'ups-right-fan-bay',w*.34,h*.14,front,w*.19,h*.66,.070,'#171d20',.42,.28);
        part(root,'ups-main-white-door',-w*.10,h*.14,front-.022,w*.33,h*.67,.043,'#f7f8f7',.36,.27);
        part(root,'ups-control-spine',w*.12,h*.14,front-.045,w*.15,h*.67,.050,'#d9ece0',.25,.22);
        part(root,'ups-green-light-column',w*.12,h*.20,front-.074,w*.055,h*.54,.016,'#a6edbf',.05,.12);
        part(root,'ups-hmi-bezel',w*.12,h*.48,front-.085,w*.10,h*.16,.020,'#2b3840',.16,.20);
        part(root,'ups-hmi-display',w*.12,h*.51,front-.098,w*.070,h*.085,.010,'#6ed0df',.05,.14);
        part(root,'ups-status-led',w*.12,h*.43,front-.100,w*.042,.022,.012,'#45d98c',.04,.20);
        // Circular fan details make the dark intake bays legible from the inspection view.
        for(const xx of [-w*.36,w*.34])for(let row=0;row<4;row++)
          round(root,'ups-intake-fan-'+xx+'-'+row,xx,h*.22+row*h*.145,front-.087,Math.min(w,h)*.040,.022,'#48565a');
        batch(root,'ups-vent-slots',Array.from({length:24},(_,i)=>[
          (i%2?-w*.36:w*.34),h*.18+Math.floor(i/2)*h*.056,front-.102]),w*.095,.014,.012,'#3c494c');
        schneiderLogoPanel(root,'ups-front',-w*.08,h*.84,front-.090,w*.58,h*.18,0);
        // White bypass section with a smaller operational HMI sits to the right, as on a modular UPS lineup.
        part(root,'ups-bypass-column',w*.43,h*.14,0,w*.10,h*.70,d*.76,'#f1f3f2',.50,.30);
        part(root,'ups-bypass-hmi',w*.43,h*.55,front-.39,w*.068,h*.10,.018,'#6ed0df',.06,.16);
        part(root,'ups-input-output-strip',w*.43,h*.38,front-.395,w*.068,h*.032,.016,'#54d98d',.05,.20);
        for(const x of [-w*.40,w*.40])for(const z of [-d*.30,d*.30])
          round(root,'ups-foot-'+x+'-'+z,x,.02,z,.09,.14,'#242a2d');
        break;
      }
      case 'ats': {
        // ATS is intentionally unlike a distribution cubicle: two source sections,
        // a centre transfer mechanism, and a visible AUTO-transfer controller.
        part(root,'ats-plinth',0,.035,0,w*1.02,.13,d*.92,'#2c4553',.58,.38);
        part(root,'ats-cabinet',0,.12,0,w*.96,h*.86,d*.84,'#65798a',.57,.37);
        const operatingFace=-front;
        for(const [side,source,color] of [[-1,'UTILITY','#4f9fc0'],[1,'GENERATOR','#d39b54']]){
          const x=side*w*.27;
          part(root,'ats-source-section-'+source,x,h*.14,operatingFace,w*.25,h*.62,.070,'#203e52',.38,.30);
          part(root,'ats-source-inlet-'+source,x,h*.67,operatingFace+.06,w*.17,.16,.026,color,.18,.28);
          part(root,'ats-source-bus-'+source,x,h*.45,operatingFace+.083,w*.16,.035,.024,color,.35,.24);
          decal(root,'ats-source-label-'+source,source,x,h*.30,operatingFace+.087,w*.17,h*.065,'#193442','#efffff');
        }
        part(root,'ats-transfer-section',0,h*.15,operatingFace,w*.22,h*.59,.077,'#9aaab0',.55,.35);
        part(root,'ats-transfer-window',0,h*.49,operatingFace+.07,w*.13,.24,.026,'#122f42',.15,.25);
        const handle=part(root,'ats-transfer-handle',0,h*.51,operatingFace+.096,w*.055,.18,.018,'#e5b94d',.30,.25);
        handle.rotation.z=-.62;
        part(root,'ats-link-utility',-w*.14,h*.54,operatingFace+.091,w*.15,.032,.020,'#4f9fc0',.35,.24);
        part(root,'ats-link-generator',w*.14,h*.54,operatingFace+.091,w*.15,.032,.020,'#d39b54',.35,.24);
        statusPanel(root,'ats-controller','AUTO TRANSFER','UTILITY ⇄ GEN',0,h*.80,operatingFace+.10,w*.38,h*.105,0);
        part(root,'ats-normal-led',-w*.055,h*.69,operatingFace+.106,.040,.040,.016,'#53d7aa',.10,.20);
        part(root,'ats-alert-led',w*.055,h*.69,operatingFace+.106,.040,.040,.016,'#f0c35e',.10,.20);
        decal(root,'ats-purpose','AUTOMATIC TRANSFER SWITCH',0,h*.12,operatingFace+.10,w*.58,h*.070,'#294757','#efffff');
        break;
      }
      case 'pdu': {
        panel(root,kind,w,h,d,'#677e92');
        break;
      }
      case 'cdu': {
        // Reverse the CDU service orientation: operating cabinet face on +Z,
        // coolant manifolds and ports on the opposite (-Z) side.
        const serviceFace=-front, manifoldFace=front;
        part(root,'cdu-case',0,.12,0,w*.92,h*.93,d*.90,'#4d8799',.62,.37);
        part(root,'cdu-door',0,h*.16,serviceFace,w*.80,h*.56,.06,'#1b374c',.41,.37);
        part(root,'cdu-display',0,h*.61,serviceFace+.057,w*.37,.24,.05,'#4bd6cb',.15,.25);
        batch(root,'cdu-vents',Array.from({length:9},(_,i)=>
          [-w*.33+i*w*.083,h*.28,serviceFace+.068]),w*.055,.045,.04,'#a6cbd7');
        for(const x of [-w*.25,w*.25]){
          round(root,'coolant-manifold-'+x,x,.03,manifoldFace-.06,.14,h*.56,'#6abed0');
          round(root,'coolant-port-'+x,x,h*.66,manifoldFace-.07,.19,.14,'#e3b76a');
        }
        batch(root,'cdu-indicators',Array.from({length:6},(_,i)=>[
          -w*.32+i*w*.13,h*.11,serviceFace+.08]),w*.065,.055,.045,'#5ee4d1');
        decal(root,'cdu-service-label','COOLANT DISTRIBUTION UNIT',0,h*.82,serviceFace+.09,w*.56,h*.075,'#164455','#e8fffc');
        break;
      }
      case 'battery-bank': {
        // UPS battery strings: rack bays, individual front modules, DC busbars and cable exits.
        part(root,'battery-plinth',0,.035,0,w*.99,.14,d*.92,'#263944',.60,.38);
        part(root,'battery-cabinet',0,.14,0,w*.97,h*.88,d*.85,'#465a68',.51,.43);
        for(let bay=0;bay<3;bay++) {
          const x=-w*.30+bay*w*.30;
          part(root,'battery-bay-frame-'+bay,x,.17,front,w*.24,h*.66,.055,'#192e3b',.40,.34);
          for(let row=0;row<4;row++){
            const y=h*.24+row*h*.145;
            part(root,'battery-module-'+bay+'-'+row,x,y,front-.045,w*.19,h*.10,.028,row%2?'#7d909c':'#99a7ad',.35,.35);
            part(root,'battery-module-terminal-'+bay+'-'+row,x+w*.068,y+h*.032,front-.069,.030,.035,.012,row%2?'#4f81b4':'#cf5960',.08,.28);
          }
        }
        part(root,'battery-dc-bus-red',-w*.36,h*.77,front-.08,w*.72,.045,.03,'#d85058',.30,.26);
        part(root,'battery-dc-bus-blue',-w*.36,h*.70,front-.08,w*.72,.045,.03,'#4389bb',.30,.26);
        // Correct 50%-scale mark centred on both faces; orient each panel toward its viewing side.
        for(const side of [-1,1])batteryBrandPanel(root,'samsung-sdi-'+side,0,h*.49,side*d*.451,w*.41,h*.115,side<0?Math.PI:0);
        for(const x of [-w*.38,w*.38]){
          part(root,'battery-cable-exit-'+x,x,h*.34,d*.42,.065,h*.46,.065,'#26343e',.46,.36);
          part(root,'battery-cable-red-'+x,x+w*.025,h*.34,d*.46,.018,h*.42,.018,'#d85058',.20,.26);
        }
        batch(root,'battery-leds',Array.from({length:9},(_,i)=>[
          -w*.38+i*w*.095,h*.86,front-.08]),w*.040,.040,.024,'#83edbc');
        break;
      }
      case 'fuel-tank': {
        round(root,'fuel-cyl-tank',0,.18,0,Math.min(w,d)*.39,h*.85,'#8f9b7c');
        part(root,'tank-ring-platform',0,h*.52,0,w*.83,.15,d*.83,'#5d7685',.61,.44);
        round(root,'tank-vent',w*.18,h*.81,0,.13,.30,'#d9c7a3');
        // Dedicated vent stack and gentle exhaust plume for the fuel-system exhibit.
        const stackX=w*.28, stackZ=d*.02, stackRadius=Math.min(w,d)*.058;
        round(root,'fuel-exhaust-stack',stackX,h*.90,stackZ,stackRadius,h*.95,'#59666d');
        round(root,'fuel-exhaust-collar',stackX,h*.52,stackZ,stackRadius*1.34,.075,'#343f46');
        round(root,'fuel-exhaust-cap',stackX,h*1.41,stackZ,stackRadius*1.38,.055,'#303a40');
        for(let i=0;i<4;i++){
          const phase=i*.82;
          const smoke=new THREE.Mesh(
            new THREE.SphereGeometry(h*(.075+i*.012),12,10),
            new THREE.MeshStandardMaterial({color:'#c5ced1',transparent:true,opacity:.38-(i*.035),roughness:1,depthWrite:false})
          );
          smoke.name='equipment-part:fuel-exhaust-smoke-'+i;
          smoke.position.set(stackX+Math.sin(phase)*.035,h*(1.49+i*.10),stackZ+Math.cos(phase)*.035);
          smoke.scale.set(1.18,.82,1.08);
          smoke.userData.fuelSmoke={baseX:smoke.position.x,baseY:smoke.position.y,baseZ:smoke.position.z,phase,drift:h*.045,rise:h*.28,scale:1.18,opacity:.38-(i*.035)};
          root.add(smoke);
        }
        decal(root,'fuel-tank-label','FUEL TANK',0,h*.38,front-.035,w*.43,h*.12,'#53656a','#ffffff');
        break;
      }
      case 'water-tank': {
        // Chilled-water buffer tank: insulated vessel, level indication, service ladder and pipework.
        const radius=Math.min(w,d)*.38;
        part(root,'water-tank-foundation',0,.04,0,w*.84,.12,d*.84,'#365a6c',.54,.40);
        round(root,'water-tank-insulated-shell',0,.19,0,radius,h*.88,'#74b7cc');
        for(const [name,y] of [['lower',h*.16],['mid',h*.49],['upper',h*.78]])
          part(root,'water-tank-band-'+name,0,y,0,w*.75,.050,d*.75,'#447f97',.48,.34);
        round(root,'water-tank-top-dome',0,h*.91,0,radius*.56,.14,'#abd9e2');
        round(root,'water-tank-manhole',0,h*.99,0,radius*.23,.075,'#4b7183');
        // Front service ladder and compact sight-level column.
        const ladderX=w*.31;
        for(const rail of [-1,1])part(root,'water-ladder-rail-'+rail,ladderX+rail*w*.07,h*.42,front-.07,.030,h*.70,.030,'#d8e7e9',.58,.27);
        for(let i=0;i<7;i++)part(root,'water-ladder-rung-'+i,ladderX,h*.16+i*h*.095,front-.085,w*.18,.025,.022,'#d8e7e9',.58,.27);
        part(root,'water-level-gauge',-w*.28,h*.41,front-.08,.055,h*.54,.025,'#102f43',.18,.24);
        part(root,'water-level-blue',-w*.28,h*.28,front-.095,.030,h*.28,.018,'#51c8ec',.08,.18);
        // Blue supply and darker return risers identify the chilled-water circuit.
        round(root,'chilled-water-supply',-w*.43,h*.31,d*.15,.10,h*.60,'#48bfe1');
        round(root,'chilled-water-return',w*.43,h*.27,-d*.15,.10,h*.52,'#397eaa');
        part(root,'water-pipe-bridge',0,h*.68,d*.15,w*.72,.075,.12,'#4ab9d8',.47,.28);
        statusPanel(root,'water-tank-level','CHILLED WATER','LEVEL 78%',0,h*.56,front-.10,w*.43,h*.115,Math.PI);
        decal(root,'water-tank-label','CHILLED WATER TANK',0,h*.36,front-.115,w*.46,h*.070,'#285a72','#effdff');
        break;
      }
      case 'cooling-tower':{
        part(root,'concrete-basin',0,.08,0,w*.98,h*.17,d*.98,'#365e73',.37,.55);
        part(root,'tower-filled-media',0,h*.17,0,w*.90,h*.53,d*.86,'#83bac9',.28,.59);
        // Mirror the air-intake louvres on front and rear, so the tower
        // remains recognisable from either side of the visitor route.
        for(const side of [-1,1]){
          const face=side*d*.48;
          batch(root,'louver-blades-'+side,Array.from({length:14},(_,i)=>[
            -w*.39+i*w*.06,h*.28,face]),w*.036,h*.38,.13,'#d5e8e7');
          batch(root,'tower-drift-eliminators-'+side,Array.from({length:8},(_,i)=>[
            -w*.31+i*w*.09,h*.66,face+side*.06]),w*.06,.15,.025,'#d6e8eb');
        }
        round(root,'fan-housing',0,h*.72,0,Math.min(w,d)*.34,.30,'#427185');
        const vanes=[];
        for(let i=0;i<8;i++){
          const angle=i*Math.PI/4,rr=Math.min(w,d)*.19;
          vanes.push([Math.sin(angle)*rr,h*.88,Math.cos(angle)*rr]);
        }
        batch(root,'fan-blades',vanes,w*.13,.085,d*.15,'#e0f1f0');
        part(root,'fan-protective-grille',0,h*.95,0,w*.75,.07,d*.75,'#2c5367',.49,.53);
        part(root,'tower-hot-water-inlet',w*.39,h*.52,d*.20,.14,.22,.14,'#cc9d53');
        part(root,'tower-cold-water-outlet',-w*.39,.23,-d*.20,.14,.22,.14,'#5bb9d5');
        break;
      }
      case 'chiller':{
        part(root,'chiller-chassis',0,.14,0,w*.95,h*.67,d*.85,'#699eb2',.52,.4);
        for(let i=-1;i<=1;i++)
          round(root,'compressor-can-'+i,i*w*.25,h*.69,0,Math.min(w,d)*.13,h*.19,'#c4d5da');
        // The front is a condenser/louvre service face only: opaque backing keeps
        // the rear HMI from showing through its fin gaps.
        part(root,'chiller-front-coil-back',0,h*.25,front+.055,w*.86,h*.45,.035,'#42788e',.38,.44);
        batch(root,'condenser-coil-fins',Array.from({length:14},(_,i)=>[
          -w*.40+i*w*.062,h*.30,front]),w*.04,h*.34,.065,'#a1d2d7');
        for(const x of [-w*.28,w*.28])
          round(root,'fluid-pipe-'+x,x,h*.61,d*.49,.13,.29,'#d8b677');
        // Rear-only controller and operating temperature display.
        const rear=d*.48;
        part(root,'chiller-operator-hmi-rear',0,h*.42,rear,w*.30,.20,.040,'#061a25',.16,.18);
        statusPanel(root,'chiller-runtime-rear','CHILLER','7.2C',0,h*.52,rear+.026,w*.25,h*.115,0);
        batch(root,'chiller-top-fans',[-w*.22,w*.22].map(x=>[x,h*.90,0]),w*.25,.05,d*.25,'#314c5a');
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
        decal(root,'fire-suppression-label','FIRE SUPPRESSION',0,h*.52,d*.16,w*.64,h*.14,'#742438','#ffffff');
        break;
      }
      case 'control-desk':{
        const deskId=String(a.id||'').toLowerCase(),isNoc=/noc/.test(deskId),isSecurity=/security/.test(deskId);
        if(isNoc){
          // NOC: operators and their consoles are in front; CCTV/BMS displays are on the rear wall.
          part(root,'noc-floor',0,.04,0,w*.96,.10,d*.94,'#263b48',.26,.52);
          part(root,'noc-console-main',0,.20,d*.14,w*.84,h*.16,d*.25,'#a98255',.27,.42);
          part(root,'noc-console-front',0,.09,d*.27,w*.86,h*.31,.09,'#314f62',.40,.34);
          for(const x of [-w*.20,w*.20]){
            part(root,'noc-seat-'+x,x,.08,d*.14,w*.17,h*.26,d*.17,'#17242c',.30,.38);
            round(root,'noc-operator-head-'+x,x,h*.53,d*.26,Math.min(w,d)*.044,.19,'#d9a884');
            part(root,'noc-operator-body-'+x,x,h*.30,d*.26,w*.12,h*.26,d*.10,'#405d7c',.24,.42);
            part(root,'noc-operator-arm-'+x,x+(x<0?w*.09:-w*.09),h*.39,d*.20,.032,h*.08,.14,'#d9a884',.12,.38);
            part(root,'noc-console-monitor-'+x,x,h*.47,d*.02,w*.17,h*.20,.035,'#142a38',.16,.24);
            part(root,'noc-console-monitor-live-'+x,x,h*.51,d*.0,w*.14,h*.13,.012,'#52b8d1',.05,.16);
          }
          nocOperationsPanel(root,'cctv-wall','cctv',-w*.24,h*.65,-d*.455,w*.34,h*.36,0);
          nocOperationsPanel(root,'bms-wall','bms',w*.16,h*.65,-d*.455,w*.42,h*.36,0);
          decal(root,'noc-room-sign','NOC · CCTV / BMS',0,h*.95,-d*.425,w*.58,h*.085,'#102935','#a5fff0');
        }else if(isSecurity){
          // Security operations: a staffed incident-intake desk, not an abstract row of monitors.
          part(root,'security-floor',0,.04,0,w*.96,.10,d*.94,'#344a55',.26,.52);
          part(root,'security-reception-table',0,.18,-d*.10,w*.78,h*.14,d*.29,'#90765a',.25,.44);
          part(root,'security-reception-front',0,.08,-d*.25,w*.80,h*.30,.10,'#345d6d',.36,.33);
          part(root,'security-operator-chair',0,.08,d*.03,w*.18,h*.26,d*.18,'#1c252a',.28,.40);
          round(root,'security-operator-head',0,h*.51,d*.01,Math.min(w,d)*.047,.20,'#d8aa86');
          part(root,'security-operator-uniform',0,h*.29,d*.01,w*.14,h*.27,d*.11,'#293f59',.25,.40);
          part(root,'security-operator-arm',w*.09,h*.37,-d*.07,.035,h*.08,.14,'#d8aa86',.12,.36);
          part(root,'security-incident-monitor',w*.16,h*.47,-d*.08,w*.20,h*.20,.04,'#162b37',.18,.23);
          part(root,'security-incident-display',w*.16,h*.51,-d*.11,w*.16,h*.12,.012,'#e0b45b',.05,.18);
          part(root,'security-desk-phone',-w*.18,h*.34,-d*.13,w*.11,.045,d*.09,'#20282d',.34,.30);
          part(root,'security-incident-form',-w*.06,h*.35,-d*.16,w*.17,.020,d*.12,'#edf2ef',.05,.45);
          statusPanel(root,'security-intake-wall','SECURITY DESK','INCIDENT INTAKE',-w*.15,h*.67,-d*.32,w*.58,h*.25,0);
          part(root,'security-event-lamp',-w*.27,h*.37,-d*.12,.045,.045,.045,'#f1b847',.05,.18);
        }else{
          part(root,'counter',0,.54,0,w*.86,.18,d*.72,'#7392a4');
          part(root,'desk-front',0,.06,-d*.28,w*.80,.66,.15,'#325c73');
          for(const x of [-w*.24,0,w*.24]){
            part(root,'monitor-'+x,x,.87,-d*.12,w*.18,.58,.12,'#1d384e');
            part(root,'display-'+x,x,.94,-d*.20,w*.16,.37,.03,'#49b4ce',.1,.14);
          }
        }
        break;
      }
      case 'maintenance':{
        part(root,'counter',0,.54,0,w*.86,.18,d*.72,'#7392a4');
        part(root,'desk-front',0,.06,-d*.28,w*.80,.66,.15,'#325c73');
        break;
      }
      case 'identity-camera': {
        // Human-height standing airport identity camera: lens, screen, document reader and locker bay.
        part(root,'identity-camera-floor-base',0,.035,0,w*.62,.10,d*.55,'#263d4a',.52,.40);
        part(root,'identity-camera-upright',0,.12,0,w*.15,h*.72,d*.15,'#607987',.44,.36);
        part(root,'identity-camera-head',0,h*.72,-d*.08,w*.55,h*.20,d*.15,'#172d39',.28,.24);
        part(root,'identity-camera-display',0,h*.74,-d*.17,w*.42,h*.12,.018,'#5bd5d2',.08,.15);
        round(root,'identity-camera-lens',0,h*.88,-d*.18,Math.min(w,d)*.082,.09,'#111f2a');
        round(root,'identity-camera-lens-ring',0,h*.88,-d*.234,Math.min(w,d)*.108,.024,'#81ece0');
        part(root,'identity-camera-privacy-visor',0,h*.97,-d*.04,w*.40,.07,d*.26,'#405968',.43,.38);
        part(root,'identity-camera-document-reader',w*.25,h*.32,-d*.19,w*.20,.12,d*.26,'#d6dfe1',.27,.34);
        part(root,'identity-camera-reader-slot',w*.25,h*.38,-d*.335,w*.12,.025,.014,'#467086',.08,.20);
        part(root,'identity-camera-green-status',-w*.26,h*.42,-d*.19,w*.10,.10,.024,'#55e1a2',.10,.25);
        part(root,'identity-camera-locker-bay',w*.36,.12,d*.06,w*.22,h*.38,d*.46,'#304b5b',.42,.34);
        batch(root,'identity-camera-locker-slots',Array.from({length:3},(_,i)=>[
          w*.36,h*.20+i*h*.105,d*.30]),w*.16,.024,.02,'#87a7b1');
        break;
      }
      case 'biometric-mantrap': {
        // Two-door interlock with iris camera, fingerprint pad and card reader status pair.
        part(root,'mantrap-floor',0,.03,0,w*.96,.10,d*.95,'#485f6b',.25,.58);
        for(const xx of [-w*.40,w*.40])
          part(root,'mantrap-frame-'+xx,xx,.05,0,w*.07,h*.92,d*.82,'#8da9b3',.52,.34);
        part(root,'mantrap-outer-lintel',0,h*.83,-d*.35,w*.88,.15,.10,'#7ca7b3',.48,.35);
        part(root,'mantrap-inner-lintel',0,h*.83,d*.35,w*.88,.15,.10,'#7ca7b3',.48,.35);
        // Authentication-complete state: both glass leaves are fully swung 90° open on their outer hinges.
        // This leaves a clear centre passage instead of presenting a closed glass wall.
        const openDoors=[
          {name:'left',hinge:-w*.40,offset:w*.16,angle:-Math.PI/2},
          {name:'right',hinge:w*.40,offset:-w*.16,angle:Math.PI/2}
        ];
        for(const door of openDoors){
          const leaf=new THREE.Group();leaf.name='equipment-mantrap-open-door:'+door.name;
          leaf.position.set(door.hinge,h*.10,-d*.35);leaf.rotation.y=door.angle;root.add(leaf);
          part(leaf,'mantrap-open-glass-'+door.name,door.offset,0,0,w*.32,h*.72,.035,'#6fb2c3',.12,.20);
          part(leaf,'mantrap-open-door-edge-'+door.name,door.offset,0,-.025,w*.34,h*.025,.050,'#c2d4d8',.42,.28);
          part(leaf,'mantrap-open-handle-'+door.name,door.offset+(door.name==='left'?w*.08:-w*.08),h*.33,-.05,.035,h*.14,.025,'#3b5664',.45,.25);
        }
        part(root,'mantrap-entry-open-light',0,.11,-d*.405,w*.55,.035,.025,'#57e3a2',.06,.20);
        part(root,'mantrap-iris-pedestal',-w*.28,.12,-d*.12,w*.13,h*.58,d*.13,'#2c4655',.44,.34);
        round(root,'mantrap-iris-camera',-w*.28,h*.68,-d*.20,Math.min(w,d)*.075,.10,'#132a3a');
        round(root,'mantrap-iris-ring',-w*.28,h*.68,-d*.255,Math.min(w,d)*.10,.022,'#5be7db');
        part(root,'mantrap-fingerprint-pad',-w*.28,h*.37,-d*.23,w*.14,.10,.035,'#8bcfd1',.12,.22);
        part(root,'mantrap-card-reader',w*.27,h*.32,-d*.30,w*.15,h*.24,.09,'#263f4f',.35,.29);
        part(root,'mantrap-card-slot',w*.27,h*.43,-d*.355,w*.08,.024,.014,'#d8e9e8',.08,.22);
        part(root,'mantrap-card-led-red',w*.27,h*.53,-d*.356,w*.07,.055,.014,'#e94750',.08,.20);
        part(root,'mantrap-card-led-green',w*.27,h*.53,-d*.372,w*.07,.055,.014,'#56e3a0',.08,.20);
        break;
      }
      case 'visitor-parking': {
        // Stage 1: a separate visitor-parking scene for pre-registration work.
        part(root,'parking-asphalt',0,.03,0,w*.96,.10,d*.95,'#3d5665',.20,.68);
        for(const x of [-w*.22,w*.22]){
          part(root,'parking-bay-line-'+x,x,.105,0,.045,.015,d*.64,'#e5edf0',.08,.34);
          part(root,'parking-bay-stop-'+x,x,.11,-d*.27,w*.20,.018,.05,'#f2dc68',.08,.30);
        }
        for(const [i,x] of [-w*.22,w*.22].entries()){
          const carColor=['#557f99','#d8e0e2','#6f8994'][i];
          part(root,'visitor-car-body-'+i,x,.15,-d*.12,w*.17,h*.14,d*.32,carColor,.42,.35);
          part(root,'visitor-car-roof-'+i,x,.15+h*.14,-d*.12,w*.12,h*.07,d*.19,'#c8dde4',.25,.28);
          for(const xx of [-w*.058,w*.058])for(const zz of [-d*.115,d*.115])
            round(root,'visitor-car-wheel-'+i+'-'+xx+'-'+zz,x+xx,.07,-d*.12+zz,Math.min(w,d)*.025,.10,'#15232c');
        }
        part(root,'parking-checkin-kiosk',w*.36,.12,d*.19,w*.16,h*.50,d*.18,'#355867',.42,.36);
        part(root,'parking-kiosk-screen',w*.36,h*.48,d*.285,w*.12,h*.14,.018,'#61e0d2',.06,.18);
        part(root,'parking-kiosk-card-slot',w*.36,h*.31,d*.285,w*.07,.026,.016,'#dde9ea',.08,.20);
        part(root,'parking-barrier-post',-w*.40,.08,d*.30,.10,h*.34,.10,'#d9e1df',.40,.34);
        part(root,'parking-barrier-arm',-w*.22,h*.42,d*.30,w*.34,.055,.055,'#f4d56d',.18,.27);
        statusPanel(root,'parking-registration','PRE-REGISTRATION','ID / VISITOR LIST',0,h*.74,front-.06,w*.54,h*.19,0);
        decal(root,'parking-name','VISITOR PARKING',0,h*.94,front-.07,w*.48,h*.09,'#163845','#bff8ef');
        break;
      }
      case 'visitor-center': {
        part(root,'visitor-center-building',0,.06,0,w*.96,h*.83,d*.90,'#71808b',.40,.50);
        part(root,'visitor-center-glass',0,h*.43,front,w*.72,h*.42,.055,'#254c62',.12,.23);
        part(root,'visitor-center-ls-plaque',0,h*.73,front-.06,w*.34,.22,.03,'#d9363e',.24,.29);
        // LS mark is fixed to the façade, rather than only appearing as a floating scene label.
        part(root,'visitor-center-logo-l-vert',-w*.075,h*.70,front-.085,w*.025,.12,.014,'#ffffff',.05,.30);
        part(root,'visitor-center-logo-l-base',-w*.045,h*.65,front-.085,w*.085,.025,.014,'#ffffff',.05,.30);
        part(root,'visitor-center-logo-s-top',w*.07,h*.75,front-.085,w*.085,.025,.014,'#ffffff',.05,.30);
        part(root,'visitor-center-logo-s-mid',w*.07,h*.70,front-.085,w*.085,.025,.014,'#ffffff',.05,.30);
        part(root,'visitor-center-logo-s-base',w*.07,h*.65,front-.085,w*.085,.025,.014,'#ffffff',.05,.30);
        part(root,'visitor-center-logo-s-left',w*.03,h*.725,front-.085,w*.025,.075,.014,'#ffffff',.05,.30);
        part(root,'visitor-center-logo-s-right',w*.11,h*.675,front-.085,w*.025,.075,.014,'#ffffff',.05,.30);
        part(root,'visitor-center-canopy',0,h*.84,front,w*.88,.12,d*.34,'#d8e3e8',.45,.35);
        break;
      }
      case 'briefing-room': {
        part(root,'briefing-floor',0,.04,0,w*.94,.12,d*.92,'#536f7f',.28,.60);
        // Front is intentionally open so the classroom scene is visible during the tour.
        part(root,'briefing-backwall',0,h*.28,d*.38,w*.90,h*.56,.08,'#e0e7e4',.12,.52);
        part(root,'briefing-screen',0,h*.57,d*.32,w*.43,h*.24,.035,'#183c31',.1,.20);
        part(root,'briefing-screen-header',0,h*.84,d*.285,w*.47,.07,.02,'#d9e9df',.1,.22);
        // Board front is turned toward the open entrance so Korean lettering does not mirror in the tour view.
        statusPanel(root,'briefing-safety-board','안전 수칙','촬영 허용구역',0,h*.62,d*.286,w*.41,h*.20,Math.PI);
        part(root,'briefing-sidewall-left',-w*.45,.20,0,.06,h*.55,d*.72,'#c7d7da',.12,.54);
        part(root,'briefing-sidewall-right',w*.45,.20,0,.06,h*.55,d*.72,'#c7d7da',.12,.54);
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
        // Raised rear shutter keeps the cargo compartment visibly open.
        part(root,'loading-truck-open-tail',-w*.48,h*.47,d*.30,w*.12,h*.43,.035,'#d7e3e6',.30,.42);
        part(root,'loading-truck-cargo-opening',-w*.24,.22,d*.245,w*.48,h*.36,.025,'#243d4d',.12,.24);
        for(const xx of [-w*.33,w*.20])for(const zz of [-d*.22,d*.22])
          round(root,'loading-truck-wheel',xx,.08,zz,.12,.14,'#202d35');
        part(root,'loading-ls-logo',-w*.17,h*.49,-d*.25,w*.22,.13,.02,'#d9363e',.20,.30);
        part(root,'loading-ls-logo-stripe',-w*.17,h*.44,-d*.27,w*.31,.028,.018,'#f3f8f6',.05,.32);
        decal(root,'loading-truck-ls','LS CABLE',-w*.17,h*.50,-d*.275,w*.30,h*.10,'#d9363e','#ffffff');
        // Parked counterbalance forklift: orange body, black mast/cab frame, driver seat, wheels and raised forks.
        const forkX=w*.24,forkZ=d*.18;
        part(root,'loading-forklift-chassis',forkX,.07,forkZ,w*.30,h*.14,d*.26,'#f47b20',.45,.33);
        part(root,'loading-forklift-counterweight',forkX-w*.10,.16,forkZ,w*.14,h*.19,d*.24,'#ff9b2d',.40,.32);
        part(root,'loading-forklift-engine-cover',forkX-w*.09,h*.31,forkZ,w*.16,h*.14,d*.22,'#f68a24',.34,.30);
        for(const xx of [forkX-w*.12,forkX+w*.11])for(const zz of [forkZ-d*.12,forkZ+d*.12])
          round(root,'loading-forklift-wheel-'+xx+'-'+zz,xx,.045,zz,Math.min(w,d)*.045,.11,'#171c20');
        part(root,'loading-forklift-mast-left',forkX+w*.18,.12,forkZ-d*.09,.040,h*.58,.040,'#314a59',.55,.28);
        part(root,'loading-forklift-mast-right',forkX+w*.18,.12,forkZ+d*.09,.040,h*.58,.040,'#314a59',.55,.28);
        part(root,'loading-forklift-mast-crossbar',forkX+w*.18,h*.67,forkZ,.055,.045,d*.26,'#314a59',.55,.28);
        part(root,'loading-forklift-fork-left',forkX+w*.31,.07,forkZ-d*.07,w*.22,.035,.035,'#465158',.50,.32);
        part(root,'loading-forklift-fork-right',forkX+w*.31,.07,forkZ+d*.07,w*.22,.035,.035,'#465158',.50,.32);
        part(root,'loading-forklift-seat',forkX-w*.02,h*.30,forkZ,w*.10,h*.11,d*.12,'#1a2024',.22,.38);
        part(root,'loading-forklift-seatback',forkX-w*.10,h*.32,forkZ,.045,h*.23,d*.13,'#1a2024',.22,.38);
        for(const xx of [forkX-w*.10,forkX+w*.10]) {
          part(root,'loading-forklift-guard-post-'+xx,xx,h*.18,forkZ-d*.105,.024,h*.51,.024,'#314a59',.52,.30);
          part(root,'loading-forklift-guard-post-rear-'+xx,xx,h*.18,forkZ+d*.105,.024,h*.51,.024,'#314a59',.52,.30);
        }
        part(root,'loading-forklift-roof',forkX,h*.69,forkZ,w*.28,.035,d*.27,'#314a59',.46,.30);
        part(root,'loading-forklift-safety-bumper',forkX+w*.17,.11,forkZ,w*.07,.11,d*.29,'#ffd442',.26,.28);
        part(root,'loading-forklift-steering',forkX+w*.04,h*.44,forkZ-d*.05,.12,.025,.12,'#222a30',.46,.28);
        decal(root,'loading-forklift-mark','FORKLIFT',forkX-w*.10,h*.45,forkZ-d*.145,w*.17,h*.065,'#f47b20','#ffffff');
        const shippingBrands=['Cisco','LS Cable & System','NVIDIA'];
        for(let i=0;i<3;i++){
          const bx=[-w*.30,-w*.02,w*.26][i],brand=shippingBrands[i];
          part(root,'loading-box-'+i,bx,.10,d*.31,w*.10,h*.12,d*.09,'#c99b54',.18,.56);
          // Large, nearly edge-to-edge carton-face labels make the supplier names readable in a wide tour shot.
          shippingBrandPanel(root,'box-'+i,brand,bx,.225,d*.359,w*.097,h*.1025);
          shippingBrandPanel(root,'box-rear-'+i,brand,bx,.225,d*.261,w*.097,h*.1025);
        }
        // Green platform trolley in front of the cartons: white handrail and red caster wheels.
        const trolleyX=-w*.02,trolleyZ=d*.04;
        part(root,'loading-trolley-platform',trolleyX,.09,trolleyZ,w*.23,.055,d*.22,'#62bb7b',.34,.38);
        part(root,'loading-trolley-edge',trolleyX,.145,trolleyZ,w*.24,.035,d*.23,'#3a9162',.32,.32);
        for(const xx of [-w*.075,w*.075])for(const zz of [-d*.075,d*.075]){
          round(root,'loading-trolley-caster-'+xx+'-'+zz,trolleyX+xx,.025,trolleyZ+zz,Math.min(w,d)*.030,.08,'#d9534f');
          round(root,'loading-trolley-wheel-'+xx+'-'+zz,trolleyX+xx,.022,trolleyZ+zz,Math.min(w,d)*.018,.095,'#3a4650');
        }
        for(const xx of [-w*.095,w*.095])part(root,'loading-trolley-handle-post-'+xx,trolleyX+xx,.13,trolleyZ+d*.09,.030,h*.42,.030,'#e8f0ed',.34,.25);
        part(root,'loading-trolley-handle-top',trolleyX,h*.55,trolleyZ+d*.09,w*.22,.035,.035,'#e8f0ed',.34,.25);
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
    // Tour cameras often show assets at medium LOD: retain the operational
    // identity of switching and cooling equipment instead of a plain cabinet.
    if(['transformer','ats','switchgear','cooling-tower'].includes(kind))detailed(mid,dims,kind);
    else {
      part(mid,'mid-shell',0,.08,0,w*.88,h*.84,d*.84,shellColor,.43,.52);
      part(mid,'mid-door',0,h*.24,-d*.44,w*.7,h*.43,.08,'#2e5269');
    }
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
