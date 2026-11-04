/**
 * Real WebGL2 3D viewer for Universal Silica Fiber Lab.
 * Visualizes an actual x/y/z scalar BPM intensity volume; it does NOT solve Maxwell 3D.
 * No external libraries/CDN or GitHub token required.
 * Fiber shell/tubes are physical geometry EXTRUSIONS (not 3D FEM meshes).
 */
export const VIEWER_VERSION = '1.0.0';
const TAU = Math.PI * 2;
const clamp = (x,a,b) => Math.min(b,Math.max(a,x));
function perspective(fovy,aspect,near,far){
 const f=1/Math.tan(fovy/2),t=1/(near-far);
 return Float32Array.from([f/aspect,0,0,0,0,f,0,0,0,0,(far+near)*t,-1,0,0,2*far*near*t,0]);
}
function cross(a,b){return [a[1]*b[2]-a[2]*b[1],a[2]*b[0]-a[0]*b[2],a[0]*b[1]-a[1]*b[0]];}
function unit(v){const l=Math.hypot(...v)||1;return v.map(x=>x/l);}
function lookAt(eye,target=[0,0,0],up=[0,1,0]){
 const z=unit(eye.map((v,i)=>v-target[i])),x=unit(cross(up,z)),y=cross(z,x);
 return Float32Array.from([
 x[0],y[0],z[0],0,x[1],y[1],z[1],0,x[2],y[2],z[2],0,
 -x.reduce((s,v,i)=>s+v*eye[i],0),-y.reduce((s,v,i)=>s+v*eye[i],0),-z.reduce((s,v,i)=>s+v*eye[i],0),1]);
}
function mul4(a,b){const o=new Float32Array(16);for(let c=0;c<4;c++)for(let r=0;r<4;r++)for(let k=0;k<4;k++)o[c*4+r]+=a[k*4+r]*b[c*4+k];return o;}
function compile(gl,type,source){const s=gl.createShader(type);gl.shaderSource(s,source);gl.compileShader(s);if(!gl.getShaderParameter(s,gl.COMPILE_STATUS))throw Error(gl.getShaderInfoLog(s));return s;}
function program(gl,vs,fs){const p=gl.createProgram(),a=compile(gl,gl.VERTEX_SHADER,vs),b=compile(gl,gl.FRAGMENT_SHADER,fs);gl.attachShader(p,a);gl.attachShader(p,b);gl.linkProgram(p);gl.deleteShader(a);gl.deleteShader(b);if(!gl.getProgramParameter(p,gl.LINK_STATUS))throw Error(gl.getProgramInfoLog(p));return p;}
const VS=`#version 300 es
precision highp float;
layout(location=0) in vec3 aPos;
layout(location=1) in vec4 aColor;
uniform mat4 uMVP;
uniform float uPointScale;
out vec4 vColor;
out float vFrac;
void main(){gl_Position=uMVP*vec4(aPos,1.0);vColor=aColor;vFrac=(aPos.z+2.6)/5.2;
 gl_PointSize=clamp(uPointScale/max(gl_Position.w,0.8),1.8,15.0);
}`;
const FS=`#version 300 es
precision highp float;
in vec4 vColor;
in float vFrac;
uniform bool uPoints;
uniform float uReveal;
uniform float uTime;
out vec4 outColor;
void main(){
 if(uPoints){
   if(vFrac>uReveal+0.006)discard;
   vec2 p=gl_PointCoord*2.0-1.0;
   float halo=exp(-dot(p,p)*3.7);
   if(halo<0.035)discard;
   float pulse=0.73+0.27*cos((vFrac*20.0-uTime*2.6)*6.2831853);
   outColor=vec4(vColor.rgb*(0.70+0.3*pulse),vColor.a*halo*0.75);
 } else outColor=vColor;
}`;
function pushVertex(out,x,y,z,color){out.push(x,y,z,...color);}
function segment(out,a,b,c){pushVertex(out,...a,c);pushVertex(out,...b,c);}
function ringLines(out,x,y,r,z,scale,c,segments=48){
 for(let i=0;i<segments;i++){const a=TAU*i/segments,b=TAU*(i+1)/segments;
   segment(out,[(x+r*Math.cos(a))/scale,(y+r*Math.sin(a))/scale,z],[(x+r*Math.cos(b))/scale,(y+r*Math.sin(b))/scale,z],c);
 }
}
function shell(out,x,y,r,wall,scale,color,z0=-2.6,z1=2.6){
 if(!(r>0&&scale>0))return;
 const inner=Math.max(0.001,r-wall);
 for(const z of [z0,z1]){ringLines(out,x,y,r,z,scale,color,48);if(wall>0)ringLines(out,x,y,inner,z,scale,[...color.slice(0,3),color[3]*.72],48);}
 for(let i=0;i<12;i++){
  const a=TAU*i/12,cs=Math.cos(a),sn=Math.sin(a),c=[...color.slice(0,3),color[3]*.5];
  segment(out,[(x+r*cs)/scale,(y+r*sn)/scale,z0],[(x+r*cs)/scale,(y+r*sn)/scale,z1],c);
 }
}
/** Convert nested tube circles or solid layers to true xyz line geometry (floats: xyzrgba). */
export function buildFiberWireframe({profile='step',coreRadiusUm=5.5,coreDiameterUm=32,hcfGeometry=null,solidLayers=null,trenchStartUm=null,trenchEndUm=null}={}){
 const lines=[];
 if(profile==='hcf'){
  if(!hcfGeometry?.rings?.length)throw Error('HCF geometry is not available');
  const ext=Math.max(1,hcfGeometry.extentUm*1.13);
  shell(lines,0,0,hcfGeometry.coreRadiusUm,0,ext,[0.29,0.93,0.82,.72]);
  for(const ring of hcfGeometry.rings)for(const t of ring.layers){
   const palette=[[.40,.83,1,.70],[1,.74,.46,.66],[1,.48,.65,.64]];
   shell(lines,t.x,t.y,t.radius,t.wall,ext,palette[t.level-1]||palette[0]);
  }
  return {lines:Float32Array.from(lines),visualRadiusUm:ext,kind:'HCF_geometry_only',lineCount:lines.length/14};
 }
 const a=Math.max(.01,profile==='multilayer'&&solidLayers?.length?solidLayers[0]:coreRadiusUm),visR=Math.max(12,a*2.8,profile==='trench'&&trenchEndUm?trenchEndUm*1.25:0,solidLayers?.length?solidLayers.at(-1)*1.25:0);
 shell(lines,0,0,visR*1.02,0,visR,[.24,.62,.93,.20]);
 shell(lines,0,0,a,0,visR,[.38,1,.83,.85]);
 if(profile==='trench'){
  shell(lines,0,0,Number.isFinite(trenchStartUm)?trenchStartUm:Math.max(a*1.38,a+1),0,visR,[.98,.70,.49,.50]);
  shell(lines,0,0,Number.isFinite(trenchEndUm)?trenchEndUm:Math.max(a*2.05,a+2),0,visR,[.98,.70,.49,.38]);
 }else if(profile==='multilayer'&&Array.isArray(solidLayers)){
  for(let i=0;i<solidLayers.length;i++)shell(lines,0,0,solidLayers[i],0,visR,[.45,.72,1,.37]);
 }
 return {lines:Float32Array.from(lines),visualRadiusUm:visR,kind:'solid_geometry',lineCount:lines.length/14};
}
/** Interpolates actual saved BPM xy intensity slices to a 3D emission cloud.
 * All intensities are from the solver. Interpolation between slices only visualizes them.
 * Does NOT infer uncomputed propagation features or optical loss.
 */
export function volumeFromBpm(data,{visualRadiusUm=15,axialSamples=42,maxPoints=38000}={}){
 if(!data?.slices?.length || !Array.isArray(data.axis_um))throw Error('BPM snapshots are required');
 const N=data.grid,axis=data.axis_um,slides=data.slices;
 if(!Number.isInteger(N)||N<8||axis.length!==N)throw Error('Invalid BPM grid');
 if(slides.some(x=>x.intensity.length!==N*N))throw Error('Invalid BPM intensity shape');
 if(!Number.isFinite(visualRadiusUm)||visualRadiusUm<=0)throw Error('Invalid radius');
 const stride=Math.max(2,Math.ceil(N/25));
 let peak=0;for(const sl of slides){for(let j=0;j<sl.intensity.length;j+=stride){const v=sl.intensity[j];if(Number.isFinite(v)&&v>peak)peak=v;}}
 if(peak<=0)throw Error('BPM intensity is zero');
 const out=[],nz=clamp(Math.floor(axialSamples),6,96);
 for(let zi=0;zi<nz;zi++){
  const frac=zi/(nz-1),zActual=frac*(data.config.length_um||slides.at(-1).z_um||1);
  let upper=1;while(upper<slides.length-1 && slides[upper].z_um<zActual)upper++;
  const lower=upper-1,s0=slides[lower],s1=slides[upper],span=Math.max(1e-15,s1.z_um-s0.z_um),t=clamp((zActual-s0.z_um)/span,0,1);
  for(let iy=0;iy<N;iy+=stride)for(let ix=0;ix<N;ix+=stride){
   const j=iy*N+ix,raw=s0.intensity[j]*(1-t)+s1.intensity[j]*t,value=raw/peak;
   if(!Number.isFinite(value)||value<.021)continue;
   const q=Math.sqrt(clamp(value,0,1));
   // x/y are genuine transverse coordinates; z geometry is compressed for viewing.
   const xx=axis[ix]/visualRadiusUm,yy=axis[iy]/visualRadiusUm,zz=-2.6+5.2*frac;
   out.push(xx,yy,zz, .25+.75*q, .40+.52*q, 1-.43*q, clamp(.10+.48*q,0,1));
   if(out.length/7>=maxPoints)break;
  }
  if(out.length/7>=maxPoints)break;
 }
 return {points:Float32Array.from(out),pointCount:out.length/7,peak,axialSamples:nz,stride};
}
function gpuBuffer(gl,arr){const b=gl.createBuffer();gl.bindBuffer(gl.ARRAY_BUFFER,b);gl.bufferData(gl.ARRAY_BUFFER,arr,gl.STATIC_DRAW);return b;}
/** Attaches a genuine WebGL2 3D canvas above the old projected 2D fallback canvas. */
export function mountFiberWebGL({panel,q,getData,getHcfGeometry,getSolidModel=()=>({})}){
 const scene2D=q('f3-scene'),parent=scene2D.parentElement;
 const toolbar=document.createElement('div');toolbar.className='f3-webgl-controls';toolbar.innerHTML='<button type="button" id="f3-view-3d" aria-pressed="true">◉ 3D WebGL</button><button type="button" id="f3-view-2d" aria-pressed="false">2D 투영</button><button type="button" id="f3-auto-orbit" aria-pressed="false">↻ 자동 회전</button><span id="f3-webgl-label">마우스 드래그: 회전 · 휠: 확대/축소</span>';
 const c=document.createElement('canvas');c.id='f3-webgl';c.width=970;c.height=460;c.setAttribute('aria-label','3D WebGL animated optical fiber and BPM intensity volume');c.style.cssText='width:100%;height:auto;aspect-ratio:970/460;display:block;background:#071423;border:1px solid #38617b;border-radius:9px;touch-action:none';
 const css=document.createElement('style');css.textContent='#fiber-3d .f3-webgl-controls{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin:0 0 8px}#fiber-3d .f3-webgl-controls button{padding:6px 10px;font-size:11px}#fiber-3d .f3-webgl-controls [aria-pressed="true"]{border-color:#78e0e0;background:#16607a}#fiber-3d .f3-webgl-controls span{color:#a4c0d4;font-size:11px}';document.head.appendChild(css);
 parent.insertBefore(toolbar,scene2D);parent.insertBefore(c,scene2D);
 const label=toolbar.querySelector('#f3-webgl-label');
 let gl=null,p=null,vertexBuffer=null,volumeBuffer=null,wireframe=null,cloud=null;
 let active='3d',yaw=1.06,pitch=.36,distance=10,autoOrbit=false,now=0,reveal=0;
 let prev=0,raf=0,lastGeometryKey='',lastData=null,disposed=false,pointer=null,failed=false;
 try{
  gl=c.getContext('webgl2',{alpha:false,antialias:true,preserveDrawingBuffer:true});
  if(!gl)throw Error('WebGL2 is unavailable (GPU or browser policy)');
  p=program(gl,VS,FS);
  gl.enable(gl.DEPTH_TEST);gl.depthFunc(gl.LEQUAL);
 }catch(e){failed=true;active='2d';c.hidden=true;scene2D.hidden=false;toolbar.querySelector('#f3-view-3d').disabled=true;toolbar.querySelector('#f3-view-3d').setAttribute('aria-pressed','false');toolbar.querySelector('#f3-view-2d').setAttribute('aria-pressed','true');label.textContent='WebGL2 지원 불가: 기존 2D 투영으로 표시 ('+e.message+')';}
 function setActive(next){if(failed&&next==='3d')return;active=next;c.hidden=next!=='3d';scene2D.hidden=next==='3d';toolbar.querySelector('#f3-view-3d').setAttribute('aria-pressed',String(next==='3d'));toolbar.querySelector('#f3-view-2d').setAttribute('aria-pressed',String(next==='2d'));if(next==='3d')draw();}
 toolbar.querySelector('#f3-view-3d').onclick=()=>setActive('3d');toolbar.querySelector('#f3-view-2d').onclick=()=>setActive('2d');
 toolbar.querySelector('#f3-auto-orbit').onclick=()=>{autoOrbit=!autoOrbit;toolbar.querySelector('#f3-auto-orbit').setAttribute('aria-pressed',String(autoOrbit));};
 function updateGeometry(){
  if(failed)return;
  const profile=q('f3-profile').value;
  const solid=getSolidModel();
  let config={profile,coreRadiusUm:Number(q('f3-radius').value),solidLayers:solid.shells?.map(x=>x[0])||null,trenchStartUm:solid.trench_start_um,trenchEndUm:solid.trench_end_um};
  if(profile==='hcf'){
   try{config.hcfGeometry=getHcfGeometry();}catch(e){if(vertexBuffer){gl.deleteBuffer(vertexBuffer);vertexBuffer=null;}wireframe=null;lastGeometryKey='';label.textContent='HCF 단면 입력을 확인하세요: '+e.message;return;}
  }
  const key=JSON.stringify({p:profile,a:config.coreRadiusUm,layer:config.solidLayers,trench:[config.trenchStartUm,config.trenchEndUm],hcf:profile==='hcf'?config.hcfGeometry.input:null});
  if(key===lastGeometryKey)return;
  try{
   const geo=buildFiberWireframe(config);
   if(vertexBuffer)gl.deleteBuffer(vertexBuffer);
   vertexBuffer=gpuBuffer(gl,geo.lines);
   wireframe=geo;
   lastGeometryKey=key;
   // Rebuild light volume for the new visual scaling.
   if(lastData)updateVolume(lastData);
  }catch(e){if(vertexBuffer){gl.deleteBuffer(vertexBuffer);vertexBuffer=null;}wireframe=null;lastGeometryKey='';label.textContent='3D 구조 오류: '+e.message;}
 }
 function updateVolume(data){
  if(failed)return;
  if(volumeBuffer){gl.deleteBuffer(volumeBuffer);volumeBuffer=null;}
  lastData=data;
  if(!data?.slices?.length||!wireframe||q('f3-profile').value==='hcf'){cloud=null;return;}
  try{cloud=volumeFromBpm(data,{visualRadiusUm:wireframe.visualRadiusUm});volumeBuffer=gpuBuffer(gl,cloud.points);}
  catch(e){cloud=null;label.textContent='광장 3D 변환 실패: '+e.message;}
 }
 function bind(buf){gl.bindBuffer(gl.ARRAY_BUFFER,buf);gl.enableVertexAttribArray(0);gl.vertexAttribPointer(0,3,gl.FLOAT,false,28,0);gl.enableVertexAttribArray(1);gl.vertexAttribPointer(1,4,gl.FLOAT,false,28,12);}
 function draw(){
  if(failed||disposed||active!=='3d')return;
  const dpr=Math.min(2,window.devicePixelRatio||1),width=Math.max(1,Math.round(c.clientWidth*dpr)),height=Math.max(1,Math.round(c.clientHeight*dpr));
  if(c.width!==width||c.height!==height){c.width=width;c.height=height;}
  gl.viewport(0,0,c.width,c.height);gl.clearColor(.025,.060,.113,1);gl.clear(gl.COLOR_BUFFER_BIT|gl.DEPTH_BUFFER_BIT);
  const eye=[distance*Math.cos(pitch)*Math.sin(yaw),distance*Math.sin(pitch),distance*Math.cos(pitch)*Math.cos(yaw)];
  const vp=mul4(perspective(.82,c.width/c.height,.1,100),lookAt(eye));
  gl.useProgram(p);gl.uniformMatrix4fv(gl.getUniformLocation(p,'uMVP'),false,vp);
  gl.uniform1f(gl.getUniformLocation(p,'uPointScale'),Math.max(9,Math.min(24,16*dpr)));
  gl.uniform1f(gl.getUniformLocation(p,'uReveal'),reveal);
  gl.uniform1f(gl.getUniformLocation(p,'uTime'),now);
  if(wireframe&&vertexBuffer){gl.depthMask(true);gl.disable(gl.BLEND);gl.uniform1i(gl.getUniformLocation(p,'uPoints'),0);bind(vertexBuffer);gl.drawArrays(gl.LINES,0,wireframe.lineCount*2);}
  if(cloud?.pointCount&&volumeBuffer){gl.depthMask(false);gl.disable(gl.DEPTH_TEST);gl.enable(gl.BLEND);gl.blendFunc(gl.SRC_ALPHA,gl.ONE);gl.uniform1i(gl.getUniformLocation(p,'uPoints'),1);bind(volumeBuffer);gl.drawArrays(gl.POINTS,0,cloud.pointCount);gl.disable(gl.BLEND);gl.enable(gl.DEPTH_TEST);gl.depthMask(true);}
 }
 function tick(ts){if(disposed)return;const dt=prev?Math.min(.07,(ts-prev)/1000):0;prev=ts;now+=dt;if(autoOrbit)yaw+=dt*.18;draw();raf=requestAnimationFrame(tick);}
 if(!failed){scene2D.hidden=true;raf=requestAnimationFrame(tick);}
 let start=null;
 c.addEventListener('pointerdown',e=>{pointer={x:e.clientX,y:e.clientY};c.setPointerCapture(e.pointerId);});
 c.addEventListener('pointermove',e=>{if(!pointer)return;yaw+=(e.clientX-pointer.x)*.008;pitch=clamp(pitch+(e.clientY-pointer.y)*.008,-1.4,1.4);pointer={x:e.clientX,y:e.clientY};draw();});
 for(const event of ['pointerup','pointercancel','lostpointercapture'])c.addEventListener(event,()=>pointer=null);
 c.addEventListener('wheel',e=>{e.preventDefault();distance=clamp(distance*Math.exp(e.deltaY*.0013),4.1,27);draw();},{passive:false});
 c.addEventListener('webglcontextlost',e=>{e.preventDefault();failed=true;setActive('2d');label.textContent='WebGL 컨텍스트 손실: 2D 화면으로 전환';});
 return {
  update(data){if(failed)return;updateGeometry();if(data!==lastData)updateVolume(data);const idx=Number(q('f3-frame').value)||0;const total=data?.slices?.length||1;reveal=data?.slices?.length?clamp(idx/(total-1),0,1):0;draw();},
  png(){if(failed||active!=='3d')return null;draw();return c.toDataURL('image/png');},
  is3D(){return !failed&&active==='3d';},
  stats(){return {webgl:!failed,active,points:cloud?.pointCount||0,lines:wireframe?.lineCount||0,reveal,version:VIEWER_VERSION};},
  dispose(){disposed=true;if(raf)cancelAnimationFrame(raf);if(gl){if(vertexBuffer)gl.deleteBuffer(vertexBuffer);if(volumeBuffer)gl.deleteBuffer(volumeBuffer);if(p)gl.deleteProgram(p);}}
 };
}
