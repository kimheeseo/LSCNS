import test from 'node:test';
import assert from 'node:assert/strict';
import * as THREE from 'three';
import {createStudioBevel,createStudioContactShadows} from './industrial-visuals.mjs';
test('shared studio bevel retains unit pick bounds with finite normals and bounded triangles',()=>{
 const g=createStudioBevel(THREE);const b=g.boundingBox,size=b.getSize(new THREE.Vector3());
 for(const value of size.toArray())assert.ok(Math.abs(value-1)<1e-5);
 for(const value of g.attributes.normal.array)assert.ok(Number.isFinite(value));
 assert.ok(g.attributes.position.count/3<200,'Only selected structural surfaces use this modest bevel');
 for(const v of b.getCenter(new THREE.Vector3()).toArray())assert.ok(Math.abs(v)<1e-5);
 g.dispose();
});

test('contact shading batches only existing visible footprints without changing selection or asset data',()=>{
 const root=new THREE.Group(),a=new THREE.Group(),b=new THREE.Group();a.position.set(3,0,7);b.visible=false;
 const registry=new Map([['a',{g:a}],['b',{g:b}]]),assets=new Map([['a',{w:2,d:4}],['b',{w:3,d:5}]]);
 const before=JSON.stringify([...assets]);const contacts=createStudioContactShadows(THREE,root,new THREE.Texture());contacts.update(registry,assets);
 assert.equal(contacts.mesh.count,1);assert.equal(contacts.mesh.userData.noPick,true);
 const matrix=new THREE.Matrix4();contacts.mesh.getMatrixAt(0,matrix);const position=new THREE.Vector3().setFromMatrixPosition(matrix);
 assert.ok(Math.abs(position.x-3)<1e-5&&Math.abs(position.z-7)<1e-5);
 assert.equal(JSON.stringify([...assets]),before);contacts.update(registry,assets,false);assert.equal(contacts.mesh.visible,false);
 contacts.dispose();assert.equal(root.children.length,0);
});
