/** Translation offsets for 2D COMSOL-style Array previews. Counts include the
 * original position; this does not perform Boolean union or FEM meshing. */
export function arrayOffsets({type, nx=1, ny=1, size=1, dx=0, dy=0}, maxCopies=450){
 if(!['Rectangular','Linear'].includes(type))throw Error('Array type must be Rectangular or Linear.');
 const positiveInteger=(n,label)=>{if(!Number.isFinite(n)||!Number.isInteger(n)||n<1)throw Error(label+' must be a positive integer.');return n;};
 const width=positiveInteger(type==='Linear'?size:nx,type==='Linear'?'Size':'x size');
 const height=type==='Linear'?1:positiveInteger(ny,'y size');
 if(!Number.isFinite(dx)||!Number.isFinite(dy))throw Error('Displacement must be finite.');
 if(width*height>maxCopies)throw Error('Array exceeds the '+maxCopies+' object preview limit.');
 const offsets=[];
 for(let ix=0;ix<width;ix++)for(let iy=0;iy<height;iy++)
  offsets.push({ix,iy,x:ix*dx,y:type==='Linear'?ix*dy:iy*dy});
 return offsets;
}
