/** World-coordinate camera for 2D circular Geometry previews. */
export function geometryBounds(shapes){
 if(!shapes.length)return {minX:-1e-6,maxX:1e-6,minY:-1e-6,maxY:1e-6,centerX:0,centerY:0,width:2e-6,height:2e-6};
 let minX=Infinity,maxX=-Infinity,minY=Infinity,maxY=-Infinity;
 for(const {x,y,r} of shapes){
  if(![x,y,r].every(Number.isFinite)||r<0)throw Error('Geometry bounds require finite centers and nonnegative radii.');
  minX=Math.min(minX,x-r);maxX=Math.max(maxX,x+r);minY=Math.min(minY,y-r);maxY=Math.max(maxY,y+r);
 }
 const width=Math.max(maxX-minX,1e-12),height=Math.max(maxY-minY,1e-12);
 const centerX=minX+width/2,centerY=minY+height/2;
 if(![minX,maxX,minY,maxY,width,height,centerX,centerY].every(Number.isFinite))throw Error('Geometry bounds exceed the supported coordinate range.');
 return {minX,maxX,minY,maxY,centerX,centerY,width,height};
}
export function fitGeometryView(shapes,plot,padding=.08){
 const bounds=geometryBounds(shapes),width=plot.right-plot.left,height=plot.bottom-plot.top;
 if(!(width>0&&height>0&&padding>=0&&padding<.5))throw Error('Invalid Geometry viewport.');
 const scale=Math.min(width*(1-2*padding)/bounds.width,height*(1-2*padding)/bounds.height);
 return {...bounds,scale,extent:Math.min(width,height)*.42/scale};
}
export function geometryView(shapes,plot,{extent=null,centerX=0,centerY=0,zoom=1,panX=0,panY=0}={}){
 const fit=fitGeometryView(shapes,plot);
 const reference=extent??fit.extent,cx=extent===null?fit.centerX:centerX,cy=extent===null?fit.centerY:centerY;
 const scale=Math.min(plot.right-plot.left,plot.bottom-plot.top)*.42*zoom/reference;
 if(!(scale>0&&Number.isFinite(scale)))throw Error('Invalid Geometry zoom.');
 return {plot,scale,ox:(plot.left+plot.right)/2+panX-cx*scale,oy:(plot.top+plot.bottom)/2+panY+cy*scale};
}
