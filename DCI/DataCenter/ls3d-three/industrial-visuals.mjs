// Batched contact shading follows existing visible asset footprints, never creates equipment.
export function createStudioContactShadows(THREE,root,texture){
 const geometry=new THREE.PlaneGeometry(1,1);geometry.rotateX(-Math.PI/2);
 const material=new THREE.MeshBasicMaterial({map:texture,transparent:true,opacity:.38,depthWrite:false,toneMapped:false});
 const mesh=new THREE.InstancedMesh(geometry,material,2048);mesh.name='studio-contact-shadows';mesh.userData.noPick=true;mesh.frustumCulled=false;mesh.count=0;root.add(mesh);
 const transform=new THREE.Object3D();let signature='';
 function update(registry,assets,enabled=true){
  mesh.visible=enabled;if(!enabled)return;
  const visible=[];
  for(const [id,item] of registry){const a=assets.get(id);if(item.g.visible&&a)visible.push([item.g.position.x,item.g.position.z,a.w,a.d]);}
  const next=visible.map(v=>v.join(',')).join('|');if(next===signature)return;signature=next;
  mesh.count=Math.min(visible.length,2048);
  for(let i=0;i<mesh.count;i++){const [x,z,w,d]=visible[i];transform.position.set(x,.053,z);transform.scale.set(w*1.55,1,d*1.55);transform.updateMatrix();mesh.setMatrixAt(i,transform.matrix);}
  mesh.instanceMatrix.needsUpdate=true;
 }
 function dispose(){mesh.removeFromParent();mesh.dispose();geometry.dispose();material.dispose();}
 return {mesh,update,dispose};
}
// Procedural surface finish and reflection lighting; no external HDR/model assets.
export function createStudioBevel(THREE){
 const s=new THREE.Shape(),r=.035,a=-.465,b=.465;
 s.moveTo(a,-.5);s.lineTo(b,-.5);s.quadraticCurveTo(.5,-.5,.5,a);s.lineTo(.5,b);s.quadraticCurveTo(.5,.5,b,.5);s.lineTo(a,.5);s.quadraticCurveTo(-.5,.5,-.5,b);s.lineTo(-.5,a);s.quadraticCurveTo(-.5,-.5,a,-.5);
 const geometry=new THREE.ExtrudeGeometry(s,{depth:.93,bevelEnabled:true,bevelSegments:1,steps:1,bevelSize:r,bevelThickness:r,curveSegments:2});
 geometry.center();geometry.computeBoundingBox();
 // Fit the unit box so existing dimensions, placement, and pick extents stay intact.
 const size=geometry.boundingBox.getSize(new THREE.Vector3());geometry.scale(1/size.x,1/size.y,1/size.z);geometry.computeBoundingBox();return geometry;
}
export function createIndustrialVisuals(THREE,renderer,scene){
 const textures=[],finished=new Set();let level='medium';
 function canvasTexture(width,height,paint,color=true){const c=document.createElement('canvas');c.width=width;c.height=height;const ctx=c.getContext('2d');paint(ctx,width,height);const t=new THREE.CanvasTexture(c);if(color)t.colorSpace=THREE.SRGBColorSpace;textures.push(t);return t;}
 const background=canvasTexture(16,256,(ctx,w,h)=>{const g=ctx.createLinearGradient(0,0,0,h);g.addColorStop(0,'#243340');g.addColorStop(.55,'#60727e');g.addColorStop(1,'#abb9c2');ctx.fillStyle=g;ctx.fillRect(0,0,w,h);});
 scene.background=background;scene.fog=new THREE.Fog('#91a3ae',180,420);
 const perforation=canvasTexture(256,256,(ctx,w,h)=>{ctx.fillStyle='#737c83';ctx.fillRect(0,0,w,h);for(let y=4;y<h;y+=12)for(let x=(y%24?5:11);x<w;x+=12){ctx.fillStyle='#21292f';ctx.beginPath();ctx.arc(x,y,3.1,0,Math.PI*2);ctx.fill();ctx.strokeStyle='#8c959c';ctx.lineWidth=.7;ctx.stroke();}});
 perforation.wrapS=perforation.wrapT=THREE.RepeatWrapping;perforation.repeat.set(1,2);
 const brush=canvasTexture(256,64,(ctx,w,h)=>{ctx.fillStyle='#d3d3d3';ctx.fillRect(0,0,w,h);for(let y=0;y<h;y++){ctx.fillStyle=`rgb(${192+(y*17%39)},${192+(y*17%39)},${192+(y*17%39)})`;ctx.fillRect(0,y,w,1);}},false);
 brush.wrapS=brush.wrapT=THREE.RepeatWrapping;brush.repeat.set(2,4);
 const envScene=new THREE.Scene();envScene.background=new THREE.Color('#647481');
 const panelGeometry=new THREE.PlaneGeometry(1,1),panels=[];
 for(const [pos,scale,power] of [[[-8,8,5],[10,8],3.5],[[7,5,-5],[7,10],2.3],[[0,12,0],[12,10],3]]){
  const mat=new THREE.MeshBasicMaterial({color:new THREE.Color('#e8f1f5').multiplyScalar(power),side:THREE.DoubleSide});const p=new THREE.Mesh(panelGeometry,mat);p.position.set(...pos);p.scale.set(...scale,1);p.lookAt(0,0,0);envScene.add(p);panels.push(mat);
 }
 const pmrem=new THREE.PMREMGenerator(renderer);const environment=pmrem.fromScene(envScene,.05,.1,60,{size:128});scene.environment=environment.texture;scene.environmentIntensity=.75;pmrem.dispose();panelGeometry.dispose();panels.forEach(m=>m.dispose());
 const contactTexture=canvasTexture(64,64,(ctx,w,h)=>{const g=ctx.createRadialGradient(w/2,h/2,2,w/2,h/2,w/2);g.addColorStop(0,'rgba(0,0,0,.7)');g.addColorStop(.5,'rgba(0,0,0,.35)');g.addColorStop(1,'rgba(0,0,0,0)');ctx.fillStyle=g;ctx.fillRect(0,0,w,h);});
 const contacts=createStudioContactShadows(THREE,scene,contactTexture);
 const bevel=createStudioBevel(THREE);
 const palette=new Map([['#829db1','#a6b0b8'],['#9db2c3','#747e88'],['#65839b','#505a63'],['#3c5b75','#353f48'],['#607a8b','#66717b'],['#b8cbd4','#c6cdd2'],['#36718a','#394d5c'],['#206d76','#304a50'],['#537e94','#687985'],['#708a9d','#4d606c'],['#182a39','#161e25'],['#b4c6d1','#b5bec6'],['#7898ac','#697a88'],['#8aa7bd','#9eabb5']]);
 function finish(material){if(material.metalness>=.5&&!material.transparent){finished.add(material);const map=level==='high'?brush:null;if(material.roughnessMap!==map){material.roughnessMap=map;material.needsUpdate=true;}material.envMapIntensity=level==='low'?0:.9;}}
 const lowBackground=new THREE.Color('#536672');
 function setQuality(value){level=value;scene.background=level==='low'?lowBackground:background;scene.environment=level==='low'?null:environment.texture;for(const material of finished)finish(material);} 
 function dispose(){contacts.dispose();scene.environment=null;environment.dispose();textures.forEach(t=>t.dispose());bevel.dispose();finished.clear();}
 return {bevel,perforation,palette,finish,dispose,contacts,setQuality,environment:environment.texture};
}
