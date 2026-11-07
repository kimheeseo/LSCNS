import test from 'node:test';
import assert from 'node:assert/strict';
import * as THREE from 'three';
import {createRackXray} from './rack-xray.mjs';
test('xray restores shared opaque material and never changes door pose or unrelated mesh',()=>{
  const material=new THREE.MeshStandardMaterial({metalness:.6,roughness:.5});
  const door=new THREE.Group(),mesh=new THREE.Mesh(new THREE.BoxGeometry(),material),neighbor=new THREE.Mesh(new THREE.BoxGeometry(),material);door.add(mesh);door.rotation.y=.3;
  const xray=createRackXray();xray.set('rack-a',[door]);
  assert.notEqual(mesh.material,material);assert.equal(mesh.material.opacity,.12);assert.equal(mesh.material.depthWrite,false);
  assert.equal(neighbor.material,material);assert.equal(material.transparent,false);assert.equal(door.rotation.y,.3);
  let disposed=0;mesh.material.addEventListener('dispose',()=>disposed++);xray.clear();
  assert.equal(mesh.material,material);assert.equal(disposed,1);assert.equal(xray.active,null);assert.equal(door.visible,true);assert.equal(door.rotation.y,.3);
});
test('switching racks and clearing twice releases clones and restores material arrays',()=>{
  const door=new THREE.Group(),original=[new THREE.MeshStandardMaterial(),new THREE.MeshStandardMaterial()];
  const mesh=new THREE.Mesh(new THREE.BoxGeometry(),original);door.add(mesh);const xray=createRackXray();
  xray.set('a',[door]);const clones=mesh.material;assert.equal(clones.length,2);
  xray.set('b',[]);assert.equal(mesh.material,original);assert.equal(xray.active,'b');xray.clear();xray.clear();assert.equal(xray.active,null);
});
