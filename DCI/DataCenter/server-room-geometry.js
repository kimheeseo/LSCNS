/* Geometry contract shared by rendering filters, MEP generation, and regression tests. */
(function(root,factory){
  const api=factory();
  if(typeof module==='object'&&module.exports)module.exports=api;
  if(root)root.LS3DServerRoomGeometry=api;
})(typeof globalThis!=='undefined'?globalThis:this,function(){
  'use strict';
  function roomBounds(hall,mode){
    if(!hall||!Array.isArray(hall.pos)||hall.pos.length<2)throw new TypeError('A hall zone with world coordinates is required');
    const w=mode==='server-room'?26:Number(hall.size?.[0]),d=mode==='server-room'?24:Number(hall.size?.[1]);
    if(!Number.isFinite(w)||!Number.isFinite(d)||w<=0||d<=0)throw new TypeError('Hall footprint dimensions must be positive finite values');
    return {id:mode==='server-room'?'server-room-hall':'data-hall',roomId:mode==='server-room'?'server-room-hall':'data-hall',roomType:mode,minX:hall.pos[0]-w/2,maxX:hall.pos[0]+w/2,minY:0,maxY:7,minZ:hall.pos[1]-d/2,maxZ:hall.pos[1]+d/2,allowedEquipmentTypes:mode==='server-room'?['Compute rack','Storage rack','Top-of-Rack switch','ODF','UPS','PDU','Air handling']:[]};
  }
  function assetInside(asset,b){
    if(!b||!asset)return false;
    const x=Number(asset.x),z=Number(asset.z),w=Number(asset.w??asset.width??0),d=Number(asset.d??asset.depth??0),h=Number(asset.h??asset.height??0),y=Number(asset.y??0);
    return [x,z,w,d,h,y].every(Number.isFinite)&&x-w/2>=b.minX-1e-6&&x+w/2<=b.maxX+1e-6&&z-d/2>=b.minZ-1e-6&&z+d/2<=b.maxZ+1e-6&&y>=b.minY-1e-6&&y+h<=b.maxY+1e-6;
  }
  function routeInside(route,b){return !!b&&Array.isArray(route)&&route.every(p=>Array.isArray(p)&&p.length>=3&&p.every(Number.isFinite)&&p[0]>=b.minX&&p[0]<=b.maxX&&p[1]>=b.minY&&p[1]<=b.maxY&&p[2]>=b.minZ&&p[2]<=b.maxZ);}
  return {roomBounds,assetInside,routeInside};
});
