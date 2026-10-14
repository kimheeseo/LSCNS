// STEP 06: real articulated 3D personnel — no textures, sprite planes, glTF/CDN, or new RAF.
// Routes, worker camera, patrol points and simulation clock stay owned by the existing host.
// Conceptual visitor training models; no claims of manufacturer PPE certification.
export function createNpcCharacterFactory(THREE, {cube, cylinder}) {
  const roles={
    engineer:   {uniform:'#338d91',trousers:'#2f5865',helmet:'#edc85f',vest:'#dcbd62',boots:'#243645',trim:'#c2edf0'},
    technician: {uniform:'#8e779d',trousers:'#465368',helmet:'#e9c65b',vest:'#f0b951',boots:'#293746',trim:'#e3d5f0'},
    operator:   {uniform:'#416b8f',trousers:'#263f59',helmet:'#96bed9',vest:'#64a4cd',boots:'#1e3548',trim:'#d7e9f0'},
    guard:      {uniform:'#254a76',trousers:'#263c57',helmet:'#2d4b6d',vest:'#46688d',boots:'#1d2d42',trim:'#d1deeb'}
  };
  const metrics={built:0,animated:0,roleCounts:{},lastActionCounts:{walk:0,inspect:0,pause:0,idle:0}};
  function kind(role){return roles[role] ? role : 'engineer';}
  function piece(group,label,x,y,z,w,h,d,color,metal=.1,rough=.77) {
    const mesh=cube(group,x,y,z,w,h,d,color,metal,rough);
    mesh.name='npc-part:'+label;mesh.userData.npcPart=label;
    mesh.userData.noPick=true; mesh.castShadow=false;
    return mesh;
  }
  function round(group,label,x,y,z,r,h,color) {
    const mesh=cylinder(group,x,y,z,r,h,color);
    mesh.name='npc-part:'+label;mesh.userData.npcPart=label;
    mesh.userData.noPick=true;mesh.castShadow=false;
    return mesh;
  }
  function pivot(parent,name,x,y,z) {
    const g=new THREE.Group();g.name='npc-joint:'+name;g.position.set(x,y,z);
    g.userData.noPick=true;parent.add(g);return g;
  }
  function create(n) {
    const role=kind(n.role),palette=roles[role];
    const root=new THREE.Group();root.name='npc:'+n.id;
    root.userData.npcId=n.id;root.userData.npcRole=role;
    root.userData.noPick=true;
    const torso=pivot(root,'torso',0,1.03,0);
    piece(torso,'jacket',0,0,.015,.64,.84,.36,palette.uniform);
    piece(torso,'shoulder-yoke',0,.71,.015,.70,.20,.38,palette.uniform);
    piece(torso,'vest-front',0,.13,.20,.57,.64,.048,palette.vest,.08,.76);
    piece(torso,'vest-back',0,.11,-.20,.56,.59,.042,palette.vest);
    piece(torso,'belt',0,.01,.04,.64,.14,.39,'#273643',.45,.5);
    for(const sign of [-1,1]){
      piece(torso,'retroreflective-chest-'+sign,sign*.19,.46,.23,.082,.34,.027,
        role==='guard'?'#a9bed2':'#e6f9e2',.18,.31);
      piece(torso,'shoulder-tab-'+sign,sign*.27,.74,.09,.16,.07,.31,palette.trim);
      piece(torso,'reflective-waist-'+sign,sign*.24,.14,.242,.14,.055,.031,'#f6f3d2',.05,.42);
    }
    const idCard=pivot(torso,'id-badge',-.17,.53,.252);
    piece(idCard,'badge-strap',0,-.085,0,.045,.19,.026,'#e8eff2');
    piece(idCard,'badge-card',0,-.21,.026,.22,.265,.036,'#e9eff3',.12,.4);
    piece(idCard,'badge-photo',-.043,-.143,.051,.075,.075,.012,'#62849c');
    piece(idCard,'badge-stripe',.033,-.26,.052,.10,.025,.012,'#40b6d0');
    piece(torso,'radio-case',.265,.53,-.04,.155,.27,.17,'#1e354d',.36,.51);
    piece(torso,'radio-antenna',.27,.85,-.04,.030,.20,.030,'#283e4e');

    const head=pivot(torso,'neck-head',0,.9,0);
    round(head,'neck',0,-.02,0,.115,.18,'#c9a584');
    round(head,'head',0,.13,.03,.445,.45,'#ddb89b');
    piece(head,'nose',0,.115,.264,.10,.12,.11,'#d1a784');
    for(const sign of [-1,1]){
      round(head,'ear-'+sign,sign*.235,.14,.015,.06,.13,'#cba18a');
      piece(head,'eye-'+sign,sign*.092,.20,.265,.052,.044,.022,'#243e4c');
    }
    // Head protection visually differentiates security officers and field engineers.
    const helmet=pivot(head,'helmet',0,.34,0);
    if(role==='guard'){
      round(helmet,'security-cap-crown',0,-.03,0,.57,.25,palette.helmet);
      piece(helmet,'security-cap-visor',0,-.10,.205,.51,.054,.30,'#203f63',.36,.45);
      piece(helmet,'security-cap-emblem',0,.04,.253,.15,.10,.032,'#dfc17f');
    }else{
      round(helmet,'safety-helmet-shell',0,-.04,0,.60,.27,palette.helmet);
      piece(helmet,'helmet-front-brim',0,-.11,.22,.55,.06,.27,palette.helmet,.24,.52);
      piece(helmet,'helmet-center-ridge',0,.08,0,.12,.10,.44,'#f6e8b2');
      for(const sign of [-1,1])
        piece(helmet,'helmet-side-stripe-'+sign,sign*.29,.01,.04,.05,.115,.33,palette.trim);
    }
    if(role==='operator'){
      piece(head,'headset-band',0,.43,0,.49,.063,.27,'#24495e');
      piece(head,'headset-mic',-.21,-.03,.15,.048,.11,.21,'#dae5e5');
    }

    const arms={};
    for(const sign of [-1,1]){
      const isLeft=sign<0;
      const arm=pivot(torso,isLeft?'left-shoulder':'right-shoulder',sign*.43,.73,0);
      piece(arm,isLeft?'left-sleeve':'right-sleeve',sign*.015,-.07,0,.245,.50,.255,palette.uniform);
      piece(arm,'upper-arm-retroreflective',sign*.01,-.26,.13,.244,.065,.033,palette.trim);
      const elbow=pivot(arm,isLeft?'left-elbow':'right-elbow',sign*.01,-.48,0);
      piece(elbow,'forearm',sign*.012,-.02,0,.220,.40,.22,palette.uniform);
      piece(elbow,'glove',sign*.015,-.34,.013,.20,.18,.19,'#293a42');
      arms[isLeft?'left':'right']={upper:arm,lower:elbow};
    }
    // A tablet is held up during an inspect pose, rather than a flat UI sprite.
    const tablet=pivot(arms.left.lower,'tablet',-.012,-.29,.08);
    piece(tablet,'tablet-housing',0,0,.14,.33,.24,.045,'#253847',.42,.43);
    piece(tablet,'tablet-screen',0,0,.172,.27,.18,.015,'#45afbf',.12,.26);
    tablet.visible=false;
    const legs={};
    for(const sign of [-1,1]){
      const isLeft=sign<0;
      const hip=pivot(root,isLeft?'left-hip':'right-hip',sign*.17,1.09,0);
      piece(hip,'upper-trouser',0,-.04,0,.27,.49,.30,palette.trousers);
      const knee=pivot(hip,isLeft?'left-knee':'right-knee',0,-.53,0);
      piece(knee,'lower-trouser',0,-.02,0,.245,.47,.28,palette.trousers);
      piece(knee,'reflective-trouser-hem',0,-.38,.02,.25,.075,.30,palette.trim);
      piece(knee,'safety-boot',0,-.48,.105,.31,.24,.46,palette.boots,.26,.61);
      legs[isLeft?'left':'right']={upper:hip,lower:knee};
    }
    metrics.built++;metrics.roleCounts[role]=(metrics.roleCounts[role]||0)+1;
    root.userData.rig={torso,head,arms,legs,tablet,role,lastAction:'idle'};
    return root;
  }
  function update(root,n,pose,clock=0) {
    const rig=root.userData.rig;if(!rig)return 'idle';
    const action=pose.action==='walk'?'walk':pose.action==='inspect'?'inspect':
      pose.action==='pause'?'pause':'idle';
    const time=Number.isFinite(clock)?clock:0;
    const phase=time*(2.6+Math.max(.1,Math.min(5,Number(n.speed)||2.5))*.35);
    const gait=action==='walk'?Math.sin(phase)*.46:0;
    const bend=action==='walk'?Math.max(0,Math.sin(phase))* .32:0;
    rig.legs.left.upper.rotation.x=gait;
    rig.legs.right.upper.rotation.x=-gait;
    rig.legs.left.lower.rotation.x=action==='walk'?-bend:0;
    rig.legs.right.lower.rotation.x=action==='walk'?-Math.max(0,-Math.sin(phase))*.32:0;
    rig.arms.left.upper.rotation.x=action==='walk'?-gait*.76:
      action==='inspect'?-1.00:action==='pause'?-0.18:-.04;
    rig.arms.right.upper.rotation.x=action==='walk'?gait*.76:
      action==='inspect'?-0.53:action==='pause'?-0.10:.04;
    rig.arms.left.lower.rotation.x=action==='inspect'?-1.03:
      action==='walk'?-.16:-.22;
    rig.arms.right.lower.rotation.x=action==='inspect'?-1.01:
      action==='walk'?-.18:-.24;
    rig.arms.left.upper.rotation.z=action==='inspect'?.12:.06;
    rig.arms.right.upper.rotation.z=action==='inspect'?- .10:-.06;
    rig.torso.position.y=1.03+
      (action==='walk'?Math.abs(Math.sin(phase*2))*.028:Math.sin(time*1.4)*.009);
    rig.torso.rotation.z=action==='walk'?Math.sin(phase)*.035:Math.sin(time*.72)*.012;
    rig.head.rotation.y=action==='walk'?Math.sin(phase*.4)*.11:Math.sin(time*.67)*.17;
    rig.head.rotation.x=action==='inspect'?.22:Math.sin(time*.85)*.022;
    rig.tablet.visible=action==='inspect';
    rig.lastAction=action;
    metrics.animated++;
    return action;
  }
  return {create,update,roles:Object.keys(roles),metrics};
}
