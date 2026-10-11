'use strict';
const assert=require('node:assert/strict');
const {roomBounds,assetInside,routeInside}=require('../server-room-geometry.js');
const hall={id:'hall',pos:[15,7],size:[42,36]};
const b=roomBounds(hall,'server-room');
assert.deepEqual([b.minX,b.maxX,b.minY,b.maxY,b.minZ,b.maxZ],[2,28,0,7,-5,19]);
const rooms=[
 {id:'company-rack-a',x:10,z:5,w:2.7,d:3,h:5},
 {id:'company-rack-b',x:15,z:5,w:2.7,d:3,h:5},
 {id:'company-storage',x:20,z:5,w:2.7,d:3,h:5},
 {id:'company-tor',x:10,z:0,w:2.7,d:3,h:5},
 {id:'company-odf',x:20,z:0,w:2.7,d:2.5,h:3},
 {id:'company-ups',x:8,z:11,w:2.8,d:2.5,h:2.8},
 {id:'company-pdu',x:12,z:11,w:2.2,d:2.2,h:2.5},
 {id:'company-crac',x:19,z:11,w:3.5,d:3,h:3}
];
assert.ok(rooms.every(a=>assetInside(a,b)),'all intended server-room equipment remains in bounds');
assert.equal(assetInside({id:'outside-busway',x:0,z:7,w:3,d:.4,h:.2},b),false);
assert.equal(assetInside({id:'ceiling-penetration',x:15,z:7,w:1,d:1,h:8},b),false);
assert.equal(routeInside([[4,6.35,0],[15,6.35,0],[26,6.35,0]],b),true);
assert.equal(routeInside([[-3,6.8,0],[32,6.8,0]],b),false);
console.log('server_room_geometry.test.cjs: 6 assertions passed');
// Execute the actual structural renderer: box Y is a base, not a centre.
const fs=require('node:fs'),path=require('node:path'),vm=require('node:vm');
const html=fs.readFileSync(path.join(__dirname,'../LS_Datacenter_3D.html'),'utf8');
const source=html.slice(html.indexOf('function zoneBase(z)'),html.indexOf('function campusGround()'));
const columns=[];vm.runInNewContext(source+';zoneBase(hall)',{hall:{...hall,color:'#fff'},facilityMode:'server-room',showRoof:false,showSection:false,line:()=>{},box:(x,y,z,w,h,d,color)=>{if(color==='#6d8a9c')columns.push({y,h})}});
assert.ok(columns.length>0&&columns.every(c=>c.y>=b.minY&&c.y+c.h<=b.maxY),'actual server-room columns remain below the room ceiling');
console.log('server-room structural columns: ceiling regression passed');
