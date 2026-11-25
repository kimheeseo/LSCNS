import test from 'node:test';
import assert from 'node:assert/strict';
import * as THREE from '../vendor/three.module.js';
import {createVisualFinish} from './visual-finish.mjs';

test('surface enhancement preserves base colors and metalness, and leaves glass unchanged',()=>{
  const scene=new THREE.Scene(),finish=createVisualFinish(THREE,{},scene);
  for(const color of ['#57eac5','#f8cb6a','#dce0e1','#287858']){
    const mat=new THREE.MeshStandardMaterial({color,...finish.materialOptions(.6,.4,false)});
    assert.equal(mat.color.getHexString(),new THREE.Color(color).getHexString());
    assert.equal(mat.metalness,.6);assert.equal(mat.roughness,.43500000000000005);
    assert.equal(mat.bumpMap.image.width,64);assert.equal(mat.bumpScale,.0025);mat.dispose();
  }
  assert.deepEqual(finish.materialOptions(.1,.2,true),{metalness:.1,roughness:.2});
  assert.equal(scene.children.length,0,'no new equipment or geometry');finish.dispose();
});
test('shared microtexture is deterministic and released when the renderer is disposed',()=>{
  const scene=new THREE.Scene(),finish=createVisualFinish(THREE,{},scene);
  const a=finish.materialOptions(.3,.5,false),b=finish.materialOptions(.8,.3,false);
  assert.equal(a.bumpMap,b.bumpMap);let disposed=0;a.bumpMap.addEventListener('dispose',()=>disposed++);
  assert.ok([...a.bumpMap.image.data].every(x=>Number.isFinite(x)));
  finish.dispose();assert.equal(disposed,1);assert.equal(scene.environment,null);
});
