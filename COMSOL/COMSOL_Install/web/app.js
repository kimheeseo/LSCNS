/* LS Wave Optics web laboratory · analytic preview, NOT in-browser FEM.
   The real Python FEM remains at COMSOL/waveoptics_m7/src/fiber_modes.py.
   Standalone, no external libraries. Pure math exports are Node-testable. */
(function(root){
'use strict';
const C=299792458,U01=2.404825557695773;
const SB=[.6961663,.4079426,.8974794],SC=[.0684043**2,.1162414**2,9.896161**2];
function silica(w){if(!(w>.21&&w<3.7))throw Error('실리카 Sellmeier 범위는 0.21–3.7 µm입니다.');return Math.sqrt(1+SB.reduce((s,b,i)=>s+b*w*w/(w*w-SC[i]),0))}
function J0(x){let y=x*x/4,t=1,s=1;for(let k=1;k<35;k++){t*=-y/(k*k);s+=t;if(Math.abs(t)<1e-16)break}return s}
function J1(x){let y=x*x/4,t=x/2,s=t;for(let k=1;k<35;k++){t*=-y/(k*(k+1));s+=t;if(Math.abs(t)<1e-16)break}return s}
function I0(x){let y=x*x/4,t=1,s=1;for(let k=1;k<55;k++){t*=y/(k*k);s+=t;if(Math.abs(t)<1e-15)break}return s}
function I1(x){let y=x*x/4,t=x/2,s=t;for(let k=1;k<55;k++){t*=y/(k*(k+1));s+=t;if(Math.abs(t)<1e-15)break}return s}
function K0(x){
 if(x<=0)throw Error('K0 입력은 양수여야 합니다.');
 if(x<=2){let y=x*x/4;return -Math.log(x/2)*I0(x)+(-.57721566+y*(.42278420+y*(.23069756+y*(.03488590+y*(.00262698+y*(.00010750+y*.00000740))))))}
 let y=2/x;return Math.exp(-x)/Math.sqrt(x)*(1.25331414+y*(-.07832358+y*(.02189568+y*(-.01062446+y*(.00587872+y*(-.00251540+y*.00053208))))))
}
function K1(x){
 if(x<=0)throw Error('K1 입력은 양수여야 합니다.');
 if(x<=2){let y=x*x/4;return Math.log(x/2)*I1(x)+(1+y*(.15443144+y*(-.67278579+y*(-.18156897+y*(-.01919402+y*(-.00110404+y*(-.00004686)))))))/x}
 let y=2/x;return Math.exp(-x)/Math.sqrt(x)*(1.25331414+y*(.23498619+y*(-.03655620+y*(.01504268+y*(-.00780353+y*(.00325614-y*.00068245))))))
}
function fiber(p){
 const wl=+p.wavelength_um,a=+p.radius_um,dn=+p.delta_n;
 if(!Number.isFinite(wl)||!Number.isFinite(a)||!Number.isFinite(dn)||wl<.8||wl>2.5||a<=0||a>30||dn<=0||dn>.04)throw Error('유효 범위: λ 0.8–2.5 µm, a 0–30 µm, Δn 0–0.04.');
 const nclad=silica(wl),ncore=nclad+dn,k0=2*Math.PI/wl,V=k0*a*Math.sqrt(ncore*ncore-nclad*nclad);
 if(V<.4||V>8)throw Error('현재 LP01 해석기에서는 V=0.4–8 범위만 지원합니다. 파라미터를 조정해 주세요.');
 const end=Math.min(V*(1-1e-10),2.4048255576);
 function fn(u){const w=Math.sqrt(Math.max(1e-20,V*V-u*u));return u*J1(u)/J0(u)-w*K1(w)/K0(w)}
 let left=Math.min(end*.00001,.00001),right=end;
 if(fn(left)*fn(right)>0)throw Error('LP01 해석근을 찾지 못했습니다. FEM 해석기로 확인해 주세요.');
 for(let i=0;i<100;i++){const mid=(left+right)/2;if(fn(mid)>0)right=mid;else left=mid}
 const u=(left+right)/2,w=Math.sqrt(V*V-u*u),neff=Math.sqrt((k0*ncore)**2-(u/a)**2)/k0;
 const jcore=J0(u),kedge=K0(w);
 const radial=r=>r<=a?J0(u*r/a)/jcore:K0(w*r/a)/kedge;
 const rmax=Math.max(12*a,a+16*a/w),N=1600,h=rmax/N;
 let s2=0,s4=0,smom=0,rad=[];
 for(let i=0;i<=N;i++){
  const r=i*h,v=radial(r),I=v*v,weight=(i===0||i===N)?1:(i%2?4:2);
  s2+=weight*I*r;s4+=weight*I*I*r;smom+=weight*I*r*r*r;
  if(i%8===0)rad.push([r,I/(radial(0)**2)]);
 }
 const A=2*Math.PI*s2*h/3,B=2*Math.PI*s4*h/3,moment=(smom/s2);
 return{wl,a,dn,ncore,nclad,V,u,w,neff,Aeff:A*A/B,MFD:2*Math.sqrt(2*moment),singleMode:V<2.4048255577,radial,radialSeries:rad,rmax};
}
function capillary(wl,radius=15,nAir=1.00027){
 wl=+wl;radius=+radius;nAir=+nAir;
 if(!(wl>.2&&wl<4&&radius>2&&radius<500&&nAir>=1&&nAir<=1.1))throw Error('HCF 입력 범위를 확인해 주세요.');
 const q=(U01/(2*Math.PI*radius))**2;
 const squared=nAir*nAir-q*wl*wl;
 if(squared<=0)throw Error('이 파장에서는 단순 모세관 모델의 전파 해가 없습니다.');
 const neff=Math.sqrt(squared),d2=-q*nAir*nAir/(neff**3),D=-wl/C*1e12*d2;
 return{wl,radius,nAir,neff,D,model:'ideal capillary (no confinement loss)'};
}
const MCF=[[12,.0002952222507259794],[16,.00007427797622971966],[20,.000018904298877409076],[24,.00000496565386076675]];
function mcf(pitch){
 pitch=+pitch;if(!Number.isFinite(pitch)||pitch<12||pitch>24)throw Error('MCF 참조 보간은 pitch 12–24 µm에 한정됩니다.');
 const i=Math.min(2,Math.floor((pitch-12)/4)),a=MCF[i],b=MCF[i+1],t=(pitch-a[0])/(b[0]-a[0]);
 const split=Math.exp(Math.log(a[1])*(1-t)+Math.log(b[1])*t);
 return{pitch,split,Lmm:1.55/(2*split)/1000,source:'M8 fixed-parameter FEM output, log interpolation (NOT new FEM)'};
}
const API={fiber,capillary,mcf,silica,J0,J1,K0,K1,MCF};
if(typeof module!=='undefined'&&module.exports)module.exports=API;
if(typeof document==='undefined')return;
const $=s=>document.querySelector(s),$$=s=>Array.from(document.querySelectorAll(s));
let mode='fiber',latest=null,series=[];
function num(id){return +$('#'+id).value}
function fmt(n,d=5){return Number.isFinite(n)?Number(n).toFixed(d):'—'}
function metric(label,value,unit=''){return '<div class="metric"><span>'+label+'</span><strong>'+value+'</strong><small>'+unit+'</small></div>'}
function chart(canvas,rows,{title,xLabel,yLabel,yMin,yMax,unit='',tone='#55e0ca'}={}){
 const box=canvas.getBoundingClientRect(),dpr=Math.min(2,window.devicePixelRatio||1);
 const width=Math.max(280,Math.round(box.width||510)),height=Math.max(200,Math.round(box.height||280));
 canvas.width=width*dpr;canvas.height=height*dpr;const ctx=canvas.getContext('2d');ctx.scale(dpr,dpr);
 const pad={l:62,r:20,t:36,b:42};ctx.fillStyle='#0d2033';ctx.fillRect(0,0,width,height);
 const xs=rows.map(r=>r[0]),ys=rows.map(r=>r[1]);
 let xmin=Math.min(...xs),xmax=Math.max(...xs),ymin=yMin??Math.min(...ys),ymax=yMax??Math.max(...ys);
 if(ymax-ymin<1e-12){ymin-=.001;ymax+=.001}let delta=ymax-ymin;ymin-=delta*.07;ymax+=delta*.07;
 const sx=x=>pad.l+(x-xmin)/(xmax-xmin||1)*(width-pad.l-pad.r);
 const sy=y=>height-pad.b-(y-ymin)/(ymax-ymin)*(height-pad.t-pad.b);
 ctx.font='11px system-ui,sans-serif';ctx.strokeStyle='#30485d';ctx.lineWidth=1;
 for(let i=0;i<5;i++){let y=pad.t+(height-pad.t-pad.b)*i/4;ctx.beginPath();ctx.moveTo(pad.l,y);ctx.lineTo(width-pad.r,y);ctx.stroke();ctx.fillStyle='#9cb4c6';ctx.textAlign='right';let value=ymax-(ymax-ymin)*i/4;ctx.fillText(Math.abs(value)>=100?value.toFixed(0):Math.abs(value)<.01?value.toExponential(1):value.toFixed(3),pad.l-8,y+4)}
 for(let i=0;i<5;i++){let x=pad.l+(width-pad.l-pad.r)*i/4;ctx.fillStyle='#9cb4c6';ctx.textAlign='center';ctx.fillText((xmin+(xmax-xmin)*i/4).toFixed(2),x,height-pad.b+18)}
 ctx.lineWidth=2;ctx.strokeStyle=tone;ctx.beginPath();rows.forEach((r,i)=>{const x=sx(r[0]),y=sy(r[1]);if(i===0)ctx.moveTo(x,y);else ctx.lineTo(x,y)});ctx.stroke();
 ctx.fillStyle='#e4f2ff';ctx.font='600 13px system-ui,sans-serif';ctx.textAlign='left';ctx.fillText(title||'',pad.l,21);
 ctx.fillStyle='#9cb4c6';ctx.font='11px system-ui,sans-serif';ctx.textAlign='center';ctx.fillText(xLabel||'',width/2,height-6);
 ctx.save();ctx.translate(14,height/2);ctx.rotate(-Math.PI/2);ctx.fillText(yLabel||'',0,0);ctx.restore();
}
function heatmap(render){
 const canvas=$('#mode-map'),ctx=canvas.getContext('2d');
 const size=240;canvas.width=size;canvas.height=size;
 const buf=ctx.createImageData(size,size),a=render.a||15,range=render.range||a*3.4;
 for(let y=0;y<size;y++)for(let x=0;x<size;x++){
  const xx=(x-size/2)/(size/2)*range,yy=(y-size/2)/(size/2)*range,rad=Math.hypot(xx,yy),theta=Math.atan2(yy,xx);
  const I=Math.min(1,Math.max(0,render.intensity(rad,xx,yy,theta)));
  const v=Math.pow(I,.48),k=(y*size+x)*4;
  buf.data[k]=Math.floor(10+230*Math.pow(v,1.8));buf.data[k+1]=Math.floor(30+200*v);buf.data[k+2]=Math.floor(51+165*(1-v)*v+130*Math.pow(v,3));buf.data[k+3]=255;
 }
 ctx.putImageData(buf,0,0);
 ctx.save();ctx.beginPath();ctx.strokeStyle='rgba(255,229,147,.86)';ctx.lineWidth=1.2;ctx.setLineDash([4,4]);
 for(const dx of render.cores||[0]){ctx.beginPath();ctx.arc(size/2+dx/range*size/2,size/2,a/range*size/2,0,Math.PI*2);ctx.stroke()}
 ctx.restore();
}
function validation(text,warning=false){$('#notice').textContent=text;$('#notice').className=warning?'notice warn':'notice'}
function run(){
 try{
  let result;
  if(mode==='fiber'){
   result=fiber({wavelength_um:num('wl'),radius_um:num('radius'),delta_n:num('dn')});const r=result,center=r.radial(0);
   $('#metrics').innerHTML=metric('Effective index',fmt(r.neff,8))+metric('Mode-field diameter',fmt(r.MFD,3),'µm · second moment')+metric('Effective area',fmt(r.Aeff,3),'µm²')+metric('Normalized V',fmt(r.V,4),r.singleMode?'LP01 regime':'LP01 shown; higher modes possible');
   heatmap({a:r.a,range:Math.max(r.a*3.3,r.MFD*.8),intensity:rad=>Math.pow(r.radial(rad)/center,2)});
   chart($('#chart-a'),r.radialSeries.map(([x,y])=>[x/r.a,y]),{title:'LP01 radial intensity · I / I(0)',xLabel:'Radius / core radius',yLabel:'Relative intensity',yMin:0});
   const base={wavelength_um:r.wl,radius_um:r.a,delta_n:r.dn},spect=[];
   for(let i=0;i<=48;i++){let w=Math.max(.8,r.wl-.18)+i*.36/48;try{spect.push([w,fiber({...base,wavelength_um:w}).neff])}catch(_){}}
   if(spect.length>1)chart($('#chart-b'),spect,{title:'LP01 effective index · wavelength sweep',xLabel:'Wavelength (µm)',yLabel:'n_eff',tone:'#eabc72'});
   series=r.radialSeries.map(([rr,y])=>[rr,y]);validation(r.singleMode?'해석식 LP01 결과입니다. 동일 파라미터의 실제 2D FEM은 Python M7에서 실행하십시오.':'주의: V ≥ 2.405로 고차 모드가 존재할 수 있습니다. 화면에는 LP01만 표시합니다.',!r.singleMode);
  }else if(mode==='mcf'){
   const pitch=num('pitch');result=mcf(pitch);
   $('#metrics').innerHTML=metric('Core spacing',fmt(pitch,1),'µm')+metric('Δn_eff',result.split.toExponential(5),'even − odd · log interpolation')+metric('Coupling length',fmt(result.Lmm,2),'mm · 1.55 µm');
   const r=4.1,sigma=5.3;
   heatmap({a:r,range:Math.max(16,pitch/2+7),cores:[-pitch/2,pitch/2],intensity:(rad,x,y)=>Math.min(1,Math.exp(-((x-pitch/2)**2+y*y)/sigma**2)+Math.exp(-((x+pitch/2)**2+y*y)/sigma**2))});
   const rows=Array.from({length:97},(_,i)=>{let p=12+i/8;return[p,mcf(p).split*1e6]});
   chart($('#chart-a'),rows,{title:'MCF supermode splitting · reference interpolation',xLabel:'Core pitch (µm)',yLabel:'Δn_eff × 10⁶',tone:'#55e0ca'});
   chart($('#chart-b'),rows.map(([p])=>[p,mcf(p).Lmm]),{title:'MCF power-transfer coupling length',xLabel:'Core pitch (µm)',yLabel:'Length (mm)',tone:'#eabc72'});
   series=rows.map(([p,val])=>[p,val]);validation('M8 고정 조건(λ=1.55 µm, a=4.1 µm, Δn=0.005)의 FEM 참조값을 로그 보간합니다. 새 FEM 계산이 아닙니다.',true);
  }else{
   result=capillary(num('hwl'),num('hradius'),num('air'));
   $('#metrics').innerHTML=metric('Effective index',fmt(result.neff,9))+metric('Dispersion D',fmt(result.D,3),'ps/(nm·km) · capillary')+metric('Core diameter',fmt(result.radius*2,1),'µm');
   heatmap({a:result.radius,range:result.radius*1.25,intensity:rad=>rad<=result.radius?Math.pow(J0(U01*rad/result.radius),2):0});
   const wl=Array.from({length:81},(_,i)=>Math.max(.8,result.wl-.35)+i*.70/80);
   const ne=wl.map(w=>[w,capillary(w,result.radius,result.nAir).neff]);
   const di=wl.map(w=>[w,capillary(w,result.radius,result.nAir).D]);
   chart($('#chart-a'),ne,{title:'Ideal hollow capillary n_eff',xLabel:'Wavelength (µm)',yLabel:'n_eff'});
   chart($('#chart-b'),di,{title:'Ideal capillary dispersion',xLabel:'Wavelength (µm)',yLabel:'D · ps/(nm·km)',tone:'#eabc72'});
   series=ne;validation('이상적인 중공 모세관 해석식입니다. NANF/ARF의 반공진, confinement loss 또는 복소 FEM 모드를 나타내지 않습니다.',true);
  }
  latest=result;$('#computed-at').textContent='계산 완료 · '+new Date().toLocaleTimeString('ko-KR');
 }catch(e){latest=null;$('#metrics').innerHTML='';validation(e.message||'계산에 실패했습니다.',true)}
}
function downloadFile(name,blob){const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),1500)}
function csv(){
 if(!latest||!series.length)return validation('먼저 계산을 실행해 주세요.',true);
 const hdr=mode==='fiber'?'radius_um,intensity_relative':mode==='mcf'?'pitch_um,index_split_times_1e6':'wavelength_um,n_eff';
 const str='\ufeff'+hdr+'\n'+series.map(r=>r.map(v=>Number(v).toPrecision(12)).join(',')).join('\n');
 downloadFile('waveoptics_'+mode+'.csv',new Blob([str],{type:'text/csv;charset=utf-8'}));
}
$$('[data-mode]').forEach(button=>button.addEventListener('click',()=>{
 mode=button.dataset.mode;$$('[data-mode]').forEach(b=>b.setAttribute('aria-selected',String(b===button)));
 $$('.input-block').forEach(el=>el.hidden=el.dataset.for!==mode);
 $('#web-method').textContent=mode==='fiber'?'LP01 Bessel 해석식':mode==='mcf'?'M8 FEM 참조값 보간':'중공 모세관 근사식';run();
}));
$('#calculate').addEventListener('click',run);
$('#export-csv').addEventListener('click',csv);
$('#export-png').addEventListener('click',()=>{$('#chart-a').toBlob(b=>{if(b)downloadFile('waveoptics_'+mode+'.png',b)},'image/png')});
$$('.input-block input').forEach(el=>el.addEventListener('change',run));
window.addEventListener('resize',()=>{if(latest)run()});
window.WaveOpticsAnalytic=API;
run();
})(typeof window!=='undefined'?window:globalThis);
