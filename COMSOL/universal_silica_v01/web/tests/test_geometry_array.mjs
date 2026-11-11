import assert from 'node:assert/strict';
import {arrayOffsets} from '../geometry_array.mjs';
assert.deepEqual(arrayOffsets({type:'Rectangular'}),[{ix:0,iy:0,x:0,y:0}]);
assert.deepEqual(arrayOffsets({type:'Rectangular',nx:2,ny:3,dx:8,dy:-6}),[
 {ix:0,iy:0,x:0,y:-0},{ix:0,iy:1,x:0,y:-6},{ix:0,iy:2,x:0,y:-12},
 {ix:1,iy:0,x:8,y:-0},{ix:1,iy:1,x:8,y:-6},{ix:1,iy:2,x:8,y:-12}
]);
assert.deepEqual(arrayOffsets({type:'Linear',size:4,dx:7,dy:-3}).map(o=>[o.x,o.y]),[[0,-0],[7,-3],[14,-6],[21,-9]]);
assert.equal(arrayOffsets({type:'Rectangular',nx:2,ny:2,dx:0,dy:0}).length,4,'Coincident positions are not silently removed.');
assert.equal(arrayOffsets({type:'Linear',size:450}).length,450);
for(const count of [0,-1,1.5,NaN,Infinity])assert.throws(()=>arrayOffsets({type:'Linear',size:count}),/positive integer/);
assert.throws(()=>arrayOffsets({type:'Rectangular',nx:2,ny:0}),/y size/);
assert.throws(()=>arrayOffsets({type:'Rectangular',nx:451,ny:1}),/limit/);
assert.throws(()=>arrayOffsets({type:'Linear',size:1,dx:Infinity}),/finite/);
assert.throws(()=>arrayOffsets({type:'Circular'}),/type/);
console.log('PASS: rectangular/linear offsets, count includes origin, signed displacement, coincident copies and validation limits');
