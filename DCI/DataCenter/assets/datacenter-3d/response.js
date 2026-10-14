/* Preserve existing fault-response visual callbacks as pooled real Three geometry. */
import * as THREE from '../../vendor/three.module.js';
export function responseLayer(scene){
 const layer=new THREE.Group();scene.add(layer);const boxG=new THREE.BoxGeometry(1,1,1),cylinderG=new THREE.CylinderGeometry(1,1,1,12),ringG=new THREE.TorusGeometry(1,.025,5,32),pool=new Map();let counts={};
 function acquire(kind,geo){const n=counts[kind]||0;counts[kind]=n+1;const list=pool.get(kind)||[];let m=list[n];if(!m){m=new THREE.Mesh(geo,new THREE.MeshStandardMaterial({roughness:.65}));layer.add(m);list.push(m);pool.set(kind,list);}m.visible=true;return m;}
 function paint(m,c){m.material.color.set(Array.isArray(c)?new THREE.Color(c[0],c[1],c[2]):c);m.material.opacity=Array.isArray(c)?c[3]??1:1;m.material.transparent=m.material.opacity<1;m.material.depthWrite=!m.material.transparent;}
 const api={box(x,y,z,w,h,d,c,rot=0){const m=acquire('box',boxG);m.position.set(x,y+h/2,z);m.scale.set(w,h,d);m.rotation.set(0,rot,0);paint(m,c);},cylinder(x,y,z,r,h,c){const m=acquire('cylinder',cylinderG);m.position.set(x,y+h/2,z);m.scale.set(r,h,r);paint(m,c);},ring(x,y,z,r,c){const m=acquire('ring',ringG);m.position.set(x,y,z);m.scale.setScalar(r);m.rotation.x=Math.PI/2;paint(m,c);},line(a,b,c){const m=acquire('line',cylinderG),A=new THREE.Vector3(...a),B=new THREE.Vector3(...b),v=B.clone().sub(A);m.position.copy(A).add(B).multiplyScalar(.5);m.scale.set(.012,v.length(),.012);m.quaternion.setFromUnitVectors(new THREE.Vector3(0,1,0),v.normalize());paint(m,c);}};
 return {layer,update(callbacks,visible){layer.visible=visible;for(const list of pool.values())for(const m of list)m.visible=false;counts={};if(visible)for(const f of callbacks)f(api);}};
}
