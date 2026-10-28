// Open-front preparation workshop: delivery -> inspection / setup / test -> Hall.
// Illustrative equipment geometry, independent of procurement/BOM quantities.
export function buildStagingRoom({part,round,decal},root,{w,d,h}){
  const station=[-w*.32,0,w*.32], rear=-d*.36;
  const colors=['#ce9e55','#53bbcd','#64cb98'];
  part(root,'staging-floor',0,.02,0,w,.10,d,'#344f60');
  part(root,'staging-rear-wall',0,.10,-d*.47,w,h*.78,.10,'#7898a7');
  decal(root,'staging-title','STAGING · 서버 설치 전 준비실',0,h*.89,-d*.455,w*.90,.48);
  const titles=['① 입고 · 개봉','② 검수 · 설정 · 시험','③ 서버홀 반입 대기'];
  const subtitles=['모델·수량·파손 확인','부품 장착 · IP/펌웨어 · 전원/통신','라벨 부착 · 카트 운반 → 서버홀'];
  for(let i=0;i<3;i++){
    part(root,'staging-zone-'+i,station[i],.08,0,w*.29,.025,d*.84,colors[i],.08,.85);
    decal(root,'staging-step-'+i,titles[i],station[i],h*.70,rear,w*.295,.43,['#705326','#175565','#1d6249'][i]);
    decal(root,'staging-note-'+i,subtitles[i],station[i],h*.54,rear+.015,w*.295,.28);
  }
  // 1. Open carton with four outward flaps, foam and removed packaging.
  const bx=station[0],bz=-d*.02,bw=w*.18,bd=d*.22;
  part(root,'staging-carton-base',bx,.16,bz,bw,.08,bd,'#bb8b47');
  for(const sign of [-1,1]){
    part(root,'staging-carton-side-'+sign,bx+sign*bw*.48,.20,bz,.06,.55,bd,'#d2aa62');
    part(root,'staging-carton-end-'+sign,bx,.20,bz+sign*bd*.47,bw,.55,.06,'#c99b54');
    const flap=part(root,'staging-open-flap-'+sign,bx+sign*bw*.65,.72,bz,bw*.35,.035,bd,'#deb977');
    flap.rotation.z=sign*.35;
    const end=part(root,'staging-open-end-flap-'+sign,bx,.70,bz+sign*bd*.64,bw,.035,bd*.36,'#deb977');
    end.rotation.x=-sign*.35;
  }
  part(root,'staging-packing-foam',bx,.27,bz,bw*.72,.35,bd*.66,'#edf0e7');
  part(root,'staging-unpacked-server',bx,.62,bz,bw*.62,.14,bd*.62,'#627d8d');
  part(root,'staging-packaging-bin',bx-w*.07,.11,d*.28,w*.12,.40,d*.17,'#526b72');
  decal(root,'staging-packaging-label','포장재 분리',bx-w*.07,.40,d*.37,w*.16,.22);
  decal(root,'staging-receipt-label','입고 검수',bx,.43,bz+bd*.51,bw*.88,.23,'#705326');
  // 2. ESD bench with actual server chassis, ports, test LEDs and laptop.
  const benchZ=-d*.06,benchW=w*.28,benchD=d*.32;
  part(root,'staging-esd-worktop',0,.94,benchZ,benchW,.10,benchD,'#469e9f');
  for(const x of [-benchW*.43,benchW*.43])for(const z of [-benchD*.40,benchD*.40])
    part(root,'staging-bench-leg-'+x+'-'+z,x,.13,benchZ+z,.08,.80,.08,'#a3b9bd');
  const sx=-w*.045,sz=benchZ+d*.02;
  part(root,'staging-test-server',sx,1.05,sz,w*.13,.23,d*.20,'#7d929f');
  part(root,'staging-server-front',sx,1.07,sz+d*.105,w*.13,.19,.04,'#172e3e');
  for(let n=0;n<6;n++){
    part(root,'staging-test-port-'+n,sx-w*.052+n*w*.019,1.11,sz+d*.113,w*.012,.07,.02,'#273e4c');
    part(root,'staging-test-led-'+n,sx-w*.052+n*w*.019,1.20,sz+d*.115,.035,.025,.02,'#5df5ad');
  }
  const lx=w*.085,lz=benchZ;
  part(root,'staging-laptop-keyboard',lx,1.05,lz,w*.095,.045,d*.14,'#a8bdc7');
  part(root,'staging-laptop-screen',lx,1.10,lz-d*.06,w*.095,.52,.045,'#203d52');
  decal(root,'staging-laptop-status','IP 설정 · TEST PASS',lx,1.38,lz-d*.055,w*.087,.32,'#124e52','#8bffd4');
  part(root,'staging-test-cable',w*.019,1.03,benchZ+d*.13,w*.12,.025,.025,'#f5d462');
  part(root,'staging-checklist-paper',-w*.075,1.005,benchZ+d*.13,w*.055,.015,d*.09,'#f1eee2');
  decal(root,'staging-test-status','전원 ON · 통신 OK',0,.73,benchZ+benchD*.52,benchW*.90,.25,'#175565');
  // 3. Labeled servers on a wheeled transfer cart with push handle.
  const cx=station[2],cz=d*.04,cw=w*.22,cd=d*.30;
  part(root,'staging-transfer-cart',cx,.32,cz,cw,.09,cd,'#469c71');
  for(const x of [-cw*.43,cw*.43])for(const z of [-cd*.40,cd*.40]){
    const wheel=round(root,'staging-cart-wheel-'+x+'-'+z,cx+x,.19,cz+z,.12,.09,'#263b48');
    wheel.rotation.z=Math.PI/2;
  }
  for(let n=0;n<2;n++){
    part(root,'staging-ready-server-'+n,cx,.42+n*.26,cz,cw*.77,.20,cd*.76,'#667f90');
    part(root,'staging-ready-front-'+n,cx,.45+n*.26,cz+cd*.39,cw*.75,.15,.035,'#203645');
  }
  for(const x of [-cw*.44,cw*.44])part(root,'staging-cart-handle-post-'+x,cx+x,.35,cz-cd*.43,.055,.90,.055,'#b2d9cc');
  part(root,'staging-cart-handle',cx,1.23,cz-cd*.43,cw*.94,.055,.055,'#b2d9cc');
  decal(root,'staging-ready-label','검수 완료 · RACK A-03',cx,1.02,cz+cd*.41,cw,.30,'#1d6249');
  // Floor flow arrows connect the stations, away from the equipment footprints.
  for(const x of [-w*.16,w*.16]){
    part(root,'staging-flow-shaft-'+x,x,.125,d*.36,w*.12,.025,.09,'#f6dc7b');
    for(const sign of [-1,1]){
      const tip=part(root,'staging-flow-arrow-'+x+'-'+sign,x+w*.055,.13,d*.36+sign*.06,.20,.025,.065,'#f6dc7b');tip.rotation.y=sign*.55;
    }
  }
  decal(root,'staging-hall-exit','서버홀 반입 →',station[2],.28,d*.47,w*.27,.28,'#1d6249');
}
