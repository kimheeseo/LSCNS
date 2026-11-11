import assert from 'node:assert/strict';
import {geometryBounds,fitGeometryView,geometryView} from '../geometry_view.mjs';
const plot={left:55,top:31,right:618,bottom:452},near=(a,b)=>assert.ok(Math.abs(a-b)<1e-9,`${a} != ${b}`);
const shapes=Array.from({length:6},(_,i)=>({x:i*400e-6,y:0,r:200e-6}));
const b=geometryBounds(shapes);near(b.minX,-200e-6);near(b.maxX,2200e-6);near(b.centerX,1000e-6);near(b.centerY,0);
const assertInside=(objects,p=plot)=>{
 const fit=fitGeometryView(objects,p),v=geometryView(objects,p,{extent:fit.extent,centerX:fit.centerX,centerY:fit.centerY});
 for(const s of objects){const x=v.ox+s.x*v.scale,y=v.oy-s.y*v.scale,r=s.r*v.scale;
  assert.ok(x-r>=p.left-1e-8&&x+r<=p.right+1e-8);assert.ok(y-r>=p.top-1e-8&&y+r<=p.bottom+1e-8);
  near((x-v.ox)/v.scale,s.x);near((v.oy-y)/v.scale,s.y);
 }
 near(v.ox+fit.centerX*v.scale,(p.left+p.right)/2);near(v.oy-fit.centerY*v.scale,(p.top+p.bottom)/2);
 return {fit,v};
};
assertInside(shapes);
assertInside(shapes.map(s=>({...s,x:0,y:s.x})));
assertInside(shapes.map(s=>({...s,x:5000e-6-s.x,y:-3000e-6-s.x*.625})));
assertInside(shapes,{left:0,top:0,right:1100,bottom:300});
assertInside(shapes,{left:0,top:0,right:300,bottom:1100});
const {fit,v}=assertInside([{x:-.1,y:.15,r:.0002}]);
const zoomed=geometryView([{x:-.1,y:.15,r:.0002}],plot,{extent:fit.extent,centerX:fit.centerX,centerY:fit.centerY,zoom:2,panX:17,panY:-11});
near(zoomed.ox+fit.centerX*zoomed.scale,(plot.left+plot.right)/2+17);
near(zoomed.oy-fit.centerY*zoomed.scale,(plot.top+plot.bottom)/2-11);
near(zoomed.scale,2*v.scale);
// Preserve the original concentric model's 84% vertical occupancy.
near(fitGeometryView([{x:0,y:0,r:40e-6}],plot).scale,(452-31)*.42/40e-6);
assert.ok(Number.isFinite(fitGeometryView([],plot).scale));
assert.throws(()=>geometryBounds([{x:NaN,y:0,r:1}]),/finite/);
console.log('PASS: horizontal/vertical/translated bounds, aspect-aware fit, world-coordinate centering, zoom/pan, inverse mapping, origin model compatibility and empty view');
