/**
 * Self-contained 2D viewport Boolean union (preview only; NOT B-rep CAD/FEM).
 * Supports filled disks and rotated rectangles, including Array-derived copies.
 * The contour is extracted from a bounded occupancy grid. Sampling introduces
 * approximation at the outer boundary and must not be used for a solver mesh.
 */
export function rectangleCorners(shape){
 const {x=0,y=0,w,h,deg=0}=shape;
 if(!(w>0&&h>0))return [];
 const a=deg*Math.PI/180,c=Math.cos(a),s=Math.sin(a);
 return [[-w/2,-h/2],[w/2,-h/2],[w/2,h/2],[-w/2,h/2]]
  .map(([u,v])=>({x:x+u*c-v*s,y:y+u*s+v*c}));
}
export function boundsOf(shape){
 if(shape?.kind==="union"&&shape.members?.length){
  return shape.members.map(boundsOf).reduce((b,r)=>({
   minX:Math.min(b.minX,r.minX),minY:Math.min(b.minY,r.minY),
   maxX:Math.max(b.maxX,r.maxX),maxY:Math.max(b.maxY,r.maxY)
  }),{minX:Infinity,minY:Infinity,maxX:-Infinity,maxY:-Infinity});
 }
 if(shape?.w>0&&shape?.h>0){
  const corners=rectangleCorners(shape);
  return {minX:Math.min(...corners.map(p=>p.x)),minY:Math.min(...corners.map(p=>p.y)),
   maxX:Math.max(...corners.map(p=>p.x)),maxY:Math.max(...corners.map(p=>p.y))};
 }
 const {x=0,y=0,r=0}=shape||{};
 return {minX:x-r,minY:y-r,maxX:x+r,maxY:y+r};
}
export function containsPoint(shape,x,y){
 if(shape?.kind==="union")return (shape.members||[]).some(m=>containsPoint(m,x,y));
 if(shape?.w>0&&shape?.h>0){
  const a=(shape.deg||0)*Math.PI/180,c=Math.cos(a),s=Math.sin(a);
  const u=x-shape.x,v=y-shape.y;
  const px=u*c+v*s,py=-u*s+v*c;
  return Math.abs(px)<=shape.w/2&&Math.abs(py)<=shape.h/2;
 }
 const dx=x-shape.x,dy=y-shape.y;
 return dx*dx+dy*dy<=shape.r*shape.r;
}
function cleanLoop(vertices){
 const pts=[];
 for(let i=0;i<vertices.length;i++){
  const p=vertices[i],prev=vertices[(i+vertices.length-1)%vertices.length],next=vertices[(i+1)%vertices.length];
  const dx1=p.x-prev.x,dy1=p.y-prev.y,dx2=next.x-p.x,dy2=next.y-p.y;
  if(Math.abs(dx1*dy2-dy1*dx2)<=1e-8*Math.hypot(dx1,dy1)*Math.hypot(dx2,dy2)&&dx1*dx2+dy1*dy2>=0)continue;
  pts.push(p);
 }
 return pts;
}
export function unionContours(shapes,resolution=340){
 if(!Array.isArray(shapes)||!shapes.length)throw Error("Union requires input shapes.");
 if(shapes.length>200)throw Error("Union preview supports up to 200 input objects.");
 if(shapes.some(s=>!s?.solid||((s?.sector??360)<359.999&&!s?.w))){
  throw Error("Union supports solid full circles and solid rectangles only.");
 }
 const bound=shapes.map(boundsOf).reduce((b,r)=>({
  minX:Math.min(b.minX,r.minX),minY:Math.min(b.minY,r.minY),
  maxX:Math.max(b.maxX,r.maxX),maxY:Math.max(b.maxY,r.maxY)
 }),{minX:Infinity,minY:Infinity,maxX:-Infinity,maxY:-Infinity});
 const spanX=bound.maxX-bound.minX,spanY=bound.maxY-bound.minY;
 if(!(spanX>0&&spanY>0&&Number.isFinite(spanX)&&Number.isFinite(spanY)))throw Error("Union input bounds are invalid.");
 const cell=Math.max(spanX,spanY)/Math.max(80,Math.min(500,Math.floor(resolution)));
 const x0=bound.minX-2*cell,y0=bound.minY-2*cell;
 const nx=Math.ceil(spanX/cell)+4,ny=Math.ceil(spanY/cell)+4;
 if(nx*ny>270000)throw Error("Union preview sampling grid is too large.");
 const grid=new Uint8Array(nx*ny);
 for(let j=0;j<ny;j++)for(let i=0;i<nx;i++){
  const x=x0+(i+.5)*cell,y=y0+(j+.5)*cell;
  if(shapes.some(s=>containsPoint(s,x,y)))grid[j*nx+i]=1;
 }
 const occupied=(i,j)=>i>=0&&i<nx&&j>=0&&j<ny&&grid[j*nx+i]===1;
 const edges=[],starts=new Map();
 function edge(a,b){
  const key=a.join(",");const idx=edges.length;
  edges.push({a,b,used:false});
  if(!starts.has(key))starts.set(key,[]);
  starts.get(key).push(idx);
 }
 for(let j=0;j<ny;j++)for(let i=0;i<nx;i++){
  if(!occupied(i,j))continue;
  if(!occupied(i,j-1))edge([i,j],[i+1,j]);
  if(!occupied(i+1,j))edge([i+1,j],[i+1,j+1]);
  if(!occupied(i,j+1))edge([i+1,j+1],[i,j+1]);
  if(!occupied(i-1,j))edge([i,j+1],[i,j]);
 }
 const loops=[];
 for(let k=0;k<edges.length;k++){
  if(edges[k].used)continue;
  const first=edges[k].a.join(",");
  let pos=k,vertices=[],closed=false;
  for(let hop=0;hop<=edges.length;hop++){
   const e=edges[pos];
   if(e.used)break;
   e.used=true;
   vertices.push({x:x0+e.a[0]*cell,y:y0+e.a[1]*cell});
   const end=e.b.join(",");
   if(end===first){closed=true;break;}
   const options=starts.get(end)||[];
   const next=options.find(v=>!edges[v].used);
   if(next===undefined)break;
   pos=next;
  }
  if(closed&&vertices.length>=4){
   const loop=cleanLoop(vertices);
   if(loop.length>=3)loops.push(loop);
  }
 }
 if(!loops.length)throw Error("Union could not construct an outer contour.");
 return {loops,bounds:bound,cell};
}
