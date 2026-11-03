import {VERSION,PRESETS,silicaIndex,validate,deltaIndex,solve,dispersion,spectrum,meshStudy,outputSummary} from './solver.mjs';
import {calculateMCF,mcfLengthSeries,MCF_REFERENCE,MCF_ASSUMPTIONS} from './mcf_xt.mjs';

const $=id=>document.getElementById(id);
const keys=['wavelength_nm','core_radius_um','delta_n_core','outer_radius_um','graded_alpha','trench_start_um','trench_end_um','trench_delta_n','n2_m2_per_W','elements'];
let last=null, busy=false, sweepRows=[],meshRows=[],lastModel=null,modeCanvas=$('crossSection');
const C={teal:'#42d5d1',blue:'#68a7fb',orange:'#f9b876',grid:'#273b55',text:'#9fb4cc',pale:'#d8ecfd'};
const fnum=(v,n=4)=>Number.isFinite(v)?Number(v).toFixed(n):'—';
function errorText(msg,kind='warning'){$('notice').textContent=msg;$('notice').className='notice '+kind;}
function changeProfile(){const v=$('profile').value;document.querySelectorAll('.conditional').forEach(el=>el.hidden=el.dataset.for!==v);document.querySelectorAll('[data-fiber-only]').forEach(el=>el.hidden=v==='mcf');setModeHeadings(v==='mcf');}
function setModeHeadings(isMCF){
 const label=isMCF?['Supermode Δneff','Cross-talk · 10log₁₀(P₂/P₁)','결합 길이 Lc','결합 계수 κ']:['유효굴절률 · neff','모드 필드 직경 · MFD (PII)','유효면적 · Aeff','색분산 · D'];
 const unit=isMCF?['짝/홀 슈퍼모드 차이','dB · ideal coherent','mm','1/m']:['모드 고유값','µm · Petermann-II','µm²','ps/(nm·km) · FEM 파장 미분'];
 document.querySelectorAll('.metrics .metric').forEach((el,i)=>{el.querySelector('span').textContent=label[i];el.querySelector('small').textContent=unit[i];});
 $('calculateBtn').innerHTML=isMCF?'<span>▶</span> XT 계산 실행':'<span>▶</span> 모드 계산 실행';
 $('sweepBtn').disabled=isMCF;$('meshBtn').disabled=isMCF;
}
function applyPreset(name){const p=structuredClone(PRESETS[name==='mcf'?'g652':name]);if(!p)throw Error('알 수 없는 예제');if(name==='mcf'){p.profile='mcf';p.wavelength_nm=1550;p.core_radius_um=4.1;p.delta_n_core=0.005;$('pitch_um').value='16';$('length_mm').value='5';}$('profile').value=p.profile;for(const k of keys){$(k).value=k==='n2_m2_per_W'?(p[k]/1e-20).toString():p[k];}$('shells').value=p.shells.map(x=>x.join(',')).join('\n');$('meshLabel').textContent=p.elements; changeProfile();invalidate();}
function readModel(){if($('profile').value==='mcf'){
  if(!$('pitch_um').value.trim()||!$('length_mm').value.trim())throw Error('MCF 코어 간격과 길이를 입력하세요.');
  const pitch_um=Number($('pitch_um').value),length_mm=Number($('length_mm').value);
  calculateMCF(pitch_um,length_mm);
  return {profile:'mcf',pitch_um,length_mm,wavelength_nm:1550,core_radius_um:4.1,delta_n_core:0.005,reference:'M8 legacy FEM four-point interpolation'};
 }
 const p={profile:$('profile').value,shells:[]};for(const k of keys){const x=Number($(k).value);if($(k).value.trim()===''||!Number.isFinite(x))throw Error(`${k} 숫자를 입력하세요.`);p[k]=x;}p.n2_m2_per_W*=1e-20;
 p.elements=Math.round(p.elements);
 p.shells=$('shells').value.trim().split(/\n+/).map(s=>s.trim()).filter(Boolean).map((s,i)=>{const a=s.split(',').map(x=>Number(x.trim()));if(a.length!==2||a.some(x=>!Number.isFinite(x)))throw Error(`쉘 ${i+1}행은 반경,Δn 2개 숫자여야 합니다.`);return a;});
 return validate(p);
}
function clearCanvas(canvas,message){const c=$(canvas),ctx=c.getContext('2d');ctx.clearRect(0,0,c.width,c.height);ctx.fillStyle='#101d30';ctx.fillRect(0,0,c.width,c.height);ctx.fillStyle='#6f8ba8';ctx.textAlign='center';ctx.font='24px sans-serif';ctx.fillText(message||'계산 대기',c.width/2,c.height/2);}
function invalidate(){last=null;lastModel=null;sweepRows=[];meshRows=[];for(const id of ['neff','mfd','aeff','disp','residual','mfd2','gamma','nclad','nElements','calcTime'])$(id).textContent='—';$('neffSub').textContent='새 계산 필요';$('spectrumTag').textContent='계산 대기';$('meshStatus').textContent='계산 대기';$('meshRows').innerHTML='<tr><td colspan="5">메시 수렴 버튼을 누르면 계산됩니다.</td></tr>';for(const id of ['crossSection','indexChart','modeChart','spectrumChart'])clearCanvas(id);errorText('입력 조건이 변경됐습니다. 「모드 계산 실행」을 눌러 결과를 다시 계산하세요.');}
function renderMCF(m,model){
 last={...m,isMCF:true};lastModel=structuredClone(model);sweepRows=[];meshRows=[];
 $('neff').textContent=m.delta_neff.toExponential(6);$('mfd').textContent=m.xt_label;
 $('aeff').textContent=m.coupling_length_mm.toFixed(3);$('disp').textContent=m.kappa_per_m.toFixed(3);
 $('neffSub').textContent='M8 FEM 4점 로그 보간';$('residual').textContent='실시간 FEM 미실행';
 $('mfd2').textContent='Core 2: '+(100*m.core2_power).toFixed(3)+'%';
 $('gamma').textContent='Core 1: '+(100*m.core1_power).toFixed(3)+'%';
 $('nclad').textContent=m.delta_neff.toExponential(5);
 $('nElements').textContent='고정 M8 참조';
 $('calcTime').textContent=m.length_mm.toFixed(3)+' mm · 이상적 결합';
 $('geometryBadge').textContent='MCF 2 CORE · SCHEMATIC';
 const c=$('crossSection'),ctx=c.getContext('2d');
 ctx.fillStyle='#101d30';ctx.fillRect(0,0,c.width,c.height);
 const cx=348,cy=223,scale=9,r=4.1*scale,gap=m.pitch_um*scale;
 [-1,1].forEach((v,i)=>{
   const x=cx+v*gap/2,p=i===0?m.core1_power:m.core2_power;
   const grd=ctx.createRadialGradient(x,cy,0,x,cy,r*2.4);
   grd.addColorStop(0,i===0?'rgba(66,213,209,.85)':'rgba(249,184,118,.85)');
   grd.addColorStop(1,'rgba(16,29,48,0)');
   ctx.fillStyle=grd;ctx.beginPath();ctx.arc(x,cy,r*2.4,0,2*Math.PI);ctx.fill();
   ctx.strokeStyle=i===0?'#42d5d1':'#f9b876';ctx.lineWidth=3;ctx.beginPath();ctx.arc(x,cy,r,0,2*Math.PI);ctx.stroke();
   ctx.fillStyle='#e6f3ff';ctx.textAlign='center';ctx.font='16px sans-serif';
   ctx.fillText('Core '+(i+1),x,cy-85);ctx.fillText((p*100).toFixed(2)+'%',x,cy+94);
 });
 ctx.fillStyle='#aec3d8';ctx.textAlign='left';ctx.font='15px sans-serif';
 ctx.fillText('2-core schematic · NOT FEM field',535,160);
 ctx.fillText('Pitch '+m.pitch_um.toFixed(1)+' µm',535,188);
 ctx.fillText('z '+m.length_mm.toFixed(2)+' mm',535,216);
 ctx.fillText('XT = '+m.xt_label,535,244);
 const pts=[];for(let i=0;i<=48;i++){const pitch=12+i*.25;pts.push([pitch,calculateMCF(pitch,0).delta_neff]);}
 makeLineChart('indexChart',[{name:'Δneff vs pitch',color:C.blue,data:pts}],{xLabel:'코어 간격 pitch (µm)',xDecimals:1,yDecimals:6,legend:true});
 const limit=Math.min(100,Math.max(20,m.length_mm*1.3,m.coupling_length_mm*2));
 const rows=mcfLengthSeries(m.pitch_um,limit,201);sweepRows=rows;
 makeLineChart('modeChart',[{name:'XT (clipped)',color:C.orange,data:rows.map(x=>[x.length_mm,x.xt_plot_db])}],{xLabel:'전파 길이 (mm)',xDecimals:1,yMin:-60,yMax:40,yDecimals:0,legend:true});
 makeLineChart('spectrumChart',[{name:'P₁',color:C.blue,data:rows.map(x=>[x.length_mm,x.core1_power])},{name:'P₂',color:C.teal,data:rows.map(x=>[x.length_mm,x.core2_power])}],{xLabel:'전파 길이 (mm)',xDecimals:1,yMin:0,yMax:1,yDecimals:2,legend:true});
 $('spectrumTag').textContent='MCF POWER TRANSFER';
 $('spectrumFoot').textContent='무손실 동일 코어 2개의 P1/P2. XT 그래프는 −60~+40 dB로 시각화만 제한합니다. +∞/−∞는 완전 결합/영전력의 수학적 한계입니다.';
 $('meshStatus').textContent='MCF FEM 재계산 없음';
 $('meshRows').innerHTML='<tr><td colspan="5">M8 저장 FEM 참조점(12,16,20,24 µm)을 로그 보간합니다. MCF 신규 메시 검사는 별도 Python M8에서 실행해야 합니다.</td></tr>';
 errorText('MCF XT 계산 완료 · 이상적 coherent 두 코어, 실제 측정값·장거리 랜덤 누화·XT(dB/km)가 아닙니다.','warning');
}
function indexAt(r,model){return silicaIndex(model.wavelength_nm)+deltaIndex(model,r);}
function makeLineChart(id,series,opts={}){
 const can=$(id),ctx=can.getContext('2d'),W=can.width,H=can.height;
 ctx.clearRect(0,0,W,H);ctx.fillStyle='#101d30';ctx.fillRect(0,0,W,H);
 const left=85,right=23,top=38,bottom=61,w=W-left-right,h=H-top-bottom;
 const pts=series.flatMap(s=>s.data).filter(p=>Number.isFinite(p[0])&&Number.isFinite(p[1]));if(!pts.length){clearCanvas(id);return;}
 let loX=opts.xMin??Math.min(...pts.map(p=>p[0])),hiX=opts.xMax??Math.max(...pts.map(p=>p[0]));let loY=opts.yMin??Math.min(...pts.map(p=>p[1])),hiY=opts.yMax??Math.max(...pts.map(p=>p[1]));
 if(hiX===loX){loX-=1;hiX+=1;}if(hiY===loY){loY-=1;hiY+=1;}
 const ygap=(hiY-loY)*.12;loY-=ygap;hiY+=ygap;
 const xp=x=>left+(x-loX)/(hiX-loX)*w, yp=y=>top+h-(y-loY)/(hiY-loY)*h;
 ctx.strokeStyle=C.grid;ctx.lineWidth=1;ctx.font='16px sans-serif';ctx.textAlign='right';ctx.fillStyle=C.text;
 for(let i=0;i<=5;i++){let y=top+h*i/5,v=hiY-(hiY-loY)*i/5;ctx.beginPath();ctx.moveTo(left,y);ctx.lineTo(left+w,y);ctx.stroke();ctx.fillText(Math.abs(v)>=100?String(v.toFixed(1)):v.toFixed(opts.yDecimals??3),left-10,y+5);}
 ctx.textAlign='center';for(let i=0;i<=5;i++){let x=left+w*i/5,v=loX+(hiX-loX)*i/5;ctx.fillText(v.toFixed(opts.xDecimals??0),x,top+h+27);}
 for(const s of series){ctx.beginPath();ctx.strokeStyle=s.color||C.teal;ctx.lineWidth=s.width??3;ctx.lineJoin='round';s.data.forEach(([x,y],i)=>i?ctx.lineTo(xp(x),yp(y)):ctx.moveTo(xp(x),yp(y)));ctx.stroke();}
 ctx.fillStyle='#b0c9dc';ctx.font='16px sans-serif';ctx.textAlign='center';ctx.fillText(opts.xLabel||'반경 (µm)',W/2,H-15);
 if(opts.legend){let x=left+8;ctx.textAlign='left';ctx.font='15px sans-serif';for(const s of series){ctx.fillStyle=s.color||C.teal;ctx.fillRect(x,15,24,3);ctx.fillStyle='#c3d7e6';ctx.fillText(s.name,x+32,20);x+=Math.min(240,60+(s.name||'').length*10);}}
}
function interp(r,field,rad){if(rad<=r[0])return field[0];if(rad>=r.at(-1))return 0;let lo=0,hi=r.length-1;while(hi-lo>1){let mid=(lo+hi)>>1;if(r[mid]<=rad)lo=mid;else hi=mid;}const t=(rad-r[lo])/(r[lo+1]-r[lo]);return field[lo]*(1-t)+field[lo+1]*t;}
function crossSection(model,sol){const can=modeCanvas,ctx=can.getContext('2d'),W=can.width,H=can.height;ctx.clearRect(0,0,W,H);ctx.fillStyle='#101d30';ctx.fillRect(0,0,W,H);
 const cx=350,cy=H/2,radPx=Math.min(H*.42,185),renderR=Math.min(model.outer_radius_um,Math.max(model.core_radius_um*3.6,model.profile==='multilayer'?model.shells.at(-1)[0]*2.2:model.profile==='trench'?model.trench_end_um*1.9:model.core_radius_um*3.5));
 const N=290;const off=document.createElement('canvas');off.width=off.height=N;const oc=off.getContext('2d'),im=oc.createImageData(N,N);let high=sol.field[0]**2;for(const q of sol.field)high=Math.max(high,q*q);
 for(let y=0;y<N;y++)for(let x=0;x<N;x++){
  const ra=Math.hypot(x-N/2,y-N/2)/(N/2)*renderR;const z=ra<=renderR?Math.max(0,Math.min(1,interp(sol.r,sol.field,ra)**2/high)):0;
  const col=heat(z),p=(y*N+x)*4;im.data[p]=col[0];im.data[p+1]=col[1];im.data[p+2]=col[2];im.data[p+3]=ra<=renderR?255:0;
 }oc.putImageData(im,0,0);ctx.save();ctx.beginPath();ctx.arc(cx,cy,radPx,0,2*Math.PI);ctx.clip();ctx.drawImage(off,cx-radPx,cy-radPx,radPx*2,radPx*2);ctx.restore();
 const borders=model.profile==='trench'?[model.core_radius_um,model.trench_start_um,model.trench_end_um]:model.profile==='multilayer'?model.shells.map(x=>x[0]):[model.core_radius_um];ctx.strokeStyle='#f2d0a7';ctx.lineWidth=2;ctx.setLineDash([7,5]);for(const b of borders){ctx.beginPath();ctx.arc(cx,cy,b/renderR*radPx,0,Math.PI*2);ctx.stroke();}ctx.setLineDash([]);
 ctx.fillStyle='#b0c6d8';ctx.textAlign='left';ctx.font='17px sans-serif';ctx.fillText('LP₀₁ intensity',565,145);ctx.font='15px sans-serif';ctx.fillText('0',605,190);ctx.fillText('1',738,190);
 for(let i=0;i<130;i++){let rgb=heat(i/129);ctx.fillStyle=`rgb(${rgb.join(',')})`;ctx.fillRect(625+i*.85,177,.9,10);}
 ctx.fillStyle='#8ba5bc';ctx.font='14px sans-serif';ctx.fillText(`r = 0 ~ ${renderR.toFixed(1)} µm`,570,235);ctx.fillText('실선 점선 = 계면',570,264);
}
function heat(t){t=Math.max(0,Math.min(1,t));const stops=[[0,[13,25,50]],[.13,[19,69,127]],[.36,[31,152,184]],[.6,[100,219,205]],[.82,[250,187,102]],[1,[255,243,184]]];for(let i=1;i<stops.length;i++){if(t<=stops[i][0]){const l=stops[i-1],h=stops[i],z=(t-l[0])/(h[0]-l[0]);return l[1].map((v,j)=>Math.round(v+(h[1][j]-v)*z));}}return stops.at(-1)[1];}
function drawIndex(m){const R=Math.min(m.outer_radius_um,Math.max(m.core_radius_um*3.5,m.profile==='multilayer'?m.shells.at(-1)[0]*1.5:m.profile==='trench'?m.trench_end_um*1.4:m.core_radius_um*3.5));const data=[];for(let i=0;i<650;i++){const rad=R*i/649;data.push([rad,indexAt(rad,m)]);}makeLineChart('indexChart',[{name:'굴절률 n(r)',color:C.blue,data}],{xMin:0,xMax:R,yDecimals:4});}
function drawMode(sol){const max=sol.field.reduce((a,b)=>Math.max(a,b*b),0);const rmax=Math.min(sol.r.at(-1),sol.r.find((r,i)=>i>0&&(sol.field[i]**2/max)<0.0002)||sol.r.at(-1));const rows=sol.r.map((r,i)=>[r,sol.field[i]**2/max]).filter(p=>p[0]<=Math.max(rmax,15));makeLineChart('modeChart',[{name:'Normalized |E|²',color:C.teal,data:rows}],{xDecimals:1,yDecimals:2,legend:true});}
function drawSpectrum(){if(!sweepRows.length){clearCanvas('spectrumChart','파장 스윕 대기');return;}const s=sweepRows.map(x=>[x.wavelength_nm,x.D_ps_nm_km]);makeLineChart('spectrumChart',[{name:'D from FEM neff',color:C.orange,data:s}],{xLabel:'파장 (nm)',yDecimals:2,legend:true});$('spectrumTag').textContent=`${sweepRows.length} wavelengths`;$('spectrumFoot').textContent='각 파장에서 FEM을 다시 실행하고 5점 유한차분으로 D를 계산했습니다. 입력한 Δn은 파장에 대해 일정하다고 가정합니다.';}
function drawMeshes(){if(!meshRows.length)return;const tbody=$('meshRows');tbody.replaceChildren();const fine=meshRows.at(-1).n_eff;for(const x of meshRows){const row=document.createElement('tr');for(const v of [x.elements,x.n_eff.toFixed(9),x.MFD_PII_um.toFixed(4),x.Aeff_um2.toFixed(4),Math.abs(x.n_eff-fine).toExponential(2)]){const td=document.createElement('td');td.textContent=v;row.appendChild(td);}tbody.appendChild(row);}$('meshStatus').textContent='FEM 재계산 '+meshRows.length+'회';}
async function task(action,msg){if(busy)return;busy=true;for(const id of ['calculateBtn','meshBtn','sweepBtn'])$(id).disabled=true;errorText(msg);$('engineState').innerHTML='<span class="status-dot"></span> FEM 계산 중';await new Promise(r=>setTimeout(r,32));try{await action();}catch(err){console.error(err);errorText('해석 실패: '+(err?.message||err),'error');}finally{busy=false;for(const id of ['calculateBtn','meshBtn','sweepBtn'])$(id).disabled=($('profile').value==='mcf'&&id!=='calculateBtn');$('engineState').innerHTML='<span class="status-dot"></span> 브라우저 FEM 엔진 준비';}}
function updateBase(m,sol,D,ms){last=sol;lastModel=structuredClone(m);$('neff').textContent=sol.n_eff.toFixed(9);$('mfd').textContent=sol.MFD_PII_um.toFixed(4);$('aeff').textContent=sol.Aeff_um2.toFixed(4);$('disp').textContent=D.toFixed(4);$('neffSub').textContent=sol.guided_proxy?'Guided mode (scalar proxy)':'Guiding 확인 필요';$('residual').textContent=sol.residual.toExponential(2);$('mfd2').textContent=sol.MFD_2nd_um.toFixed(4)+' µm';$('gamma').textContent=sol.gamma_Wkm.toFixed(4)+' /(W·km)';$('nclad').textContent=sol.n_clad.toFixed(7);$('nElements').textContent=String(sol.elements);$('calcTime').textContent=ms.toFixed(0)+' ms · 현재 입력';$('geometryBadge').textContent=m.profile.toUpperCase();crossSection(m,sol);drawIndex(m);drawMode(sol);}
async function calc({autoSweep=false}={}){await task(async()=>{const m=readModel();if(m.profile==='mcf'){renderMCF(calculateMCF(m.pitch_um,m.length_mm),m);return;}const start=performance.now(),s=solve(m),D=dispersion(m);updateBase(m,s,D,performance.now()-start);meshRows=[];$('meshRows').innerHTML='<tr><td colspan="5">메시 수렴 버튼을 누르면 계산됩니다.</td></tr>';$('meshStatus').textContent='검사 대기';sweepRows=[];if(autoSweep){const from=Number($('sweepStart').value),to=Number($('sweepEnd').value),step=Number($('sweepStep').value);sweepRows=spectrum(m,from,to,step);}drawSpectrum();errorText(`모드 계산 완료 · ${s.elements} elements · 고유값 잔차 ${s.residual.toExponential(2)} · 현재 입력 조건`,solGood(s)?'success':'warning');},'Radial FEM 모드와 분산을 계산하고 있습니다…');}
function solGood(sol){return sol.iteration_converged&&sol.guided_proxy&&sol.residual<1e-7;}
async function doSweep(){await task(async()=>{const m=readModel();const s=spectrum(m,Number($('sweepStart').value),Number($('sweepEnd').value),Number($('sweepStep').value));sweepRows=s;if(!last){const one=solve(m);updateBase(m,one,dispersion(m),0);}drawSpectrum();errorText(`${s.length}개 파장에 대해 FEM 스윕 계산 완료`,'success');},'파장별 FEM 스윕 계산 중…');}
async function doMesh(){await task(async()=>{const m=readModel();const levels=[Math.max(100,Math.round(m.elements*.4/50)*50),Math.max(150,Math.round(m.elements*.65/50)*50),m.elements,Math.min(3000,Math.round(m.elements*1.4/50)*50)];meshRows=meshStudy(m,[...new Set(levels)].sort((a,b)=>a-b));if(!last){const one=solve(m);updateBase(m,one,dispersion(m),0);}drawMeshes();errorText(`${meshRows.length}개 메시 수준의 FEM 수렴 검사 완료 · neff 차이를 확인하세요.`,'success');},'메시 조건별 FEM 재계산 중…');}
function saveFile(name,text,mime){const blob=new Blob([text],{type:mime}),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);}
function exportJSON(){if(last?.isMCF){const doc={solver:'M8 legacy reference + ideal coherent transfer',generated_at:new Date().toISOString(),model:lastModel,result:{...last,xt_db:Number.isFinite(last.xt_db)?last.xt_db:null,xt_db_special:Number.isFinite(last.xt_db)?null:last.xt_label},length_series:sweepRows,limitations:MCF_ASSUMPTIONS};saveFile('mcf_ideal_xt_result.json',JSON.stringify(doc,null,2),'application/json');return;}if(!last){errorText('먼저 해석을 실행해야 결과를 저장할 수 있습니다.','error');return;}const doc={solver:'LS Universal Silica Fiber Web v'+VERSION,generated_at:new Date().toISOString(),basis:'radial scalar FEM; silica Sellmeier + wavelength-independent delta n; not real product certification',model:lastModel,result:outputSummary(last),D_ps_nm_km:Number($('disp').textContent),spectrum:sweepRows,mesh_convergence:meshRows};saveFile('silica_fiber_result.json',JSON.stringify(doc,null,2),'application/json');}
function exportCSV(){if(last?.isMCF){const lines=[['length_mm','core1_power','core2_power','xt_db','xt_plot_db_clipped'].join(','),...sweepRows.map(r=>[r.length_mm,r.core1_power,r.core2_power,r.xt_db===null?'':r.xt_db,r.xt_plot_db].join(','))];saveFile('mcf_ideal_xt_curve.csv',lines.join('\\n')+'\\n','text/csv;charset=utf-8');return;}if(!last){errorText('먼저 해석을 실행해야 CSV를 저장할 수 있습니다.','error');return;}const p=[['parameter','value','unit'],['profile',lastModel.profile,''],['wavelength_nm',lastModel.wavelength_nm,'nm'],['core_radius_um',lastModel.core_radius_um,'um'],['delta_n_core',lastModel.delta_n_core,''],...Object.entries(outputSummary(last)).filter(x=>typeof x[1]==='number').map(([k,v])=>[k,v,'']),['D_ps_nm_km',$('disp').textContent,'ps/(nm km)']];saveFile('silica_fiber_result.csv',p.map(x=>x.join(',')).join('\n')+'\n','text/csv;charset=utf-8');if(sweepRows.length){const ks=Object.keys(sweepRows[0]);saveFile('silica_fiber_spectrum.csv',[ks.join(','),...sweepRows.map(row=>ks.map(k=>row[k]).join(','))].join('\n')+'\n','text/csv;charset=utf-8');}}
function exportPNG(){if(last?.isMCF){const a=document.createElement('a');a.download='mcf_ideal_xt_curve.png';a.href=$('modeChart').toDataURL('image/png');a.click();return;}if(!last){errorText('먼저 모드를 계산하세요.','error');return;}const a=document.createElement('a');a.download='silica_fiber_mode.png';a.href=modeCanvas.toDataURL('image/png');a.click();}
function exportDesign(){const m=readModel();saveFile('silica_fiber_design.json',JSON.stringify({version:VERSION,model:m},null,2),'application/json');}
function importDesignFile(f){if(!f)return;if(f.size>500000){errorText('설계 JSON은 500 KB 이하여야 합니다.','error');return;}const reader=new FileReader();reader.onload=()=>{try{let o=JSON.parse(reader.result),m=o.model||o;if(m.profile==='mcf'){calculateMCF(m.pitch_um,m.length_mm);applyPreset('mcf');$('preset').value='mcf';$('profile').value='mcf';$('pitch_um').value=m.pitch_um;$('length_mm').value=m.length_mm;changeProfile();invalidate();errorText('MCF 설계 불러오기 완료. XT 계산 실행을 누르세요.','success');return;}validate(m);$('preset').value='g654';$('profile').value=m.profile;for(const k of keys)$(k).value=k==='n2_m2_per_W'?m[k]/1e-20:m[k];$('shells').value=m.shells.map(s=>s.join(',')).join('\n');$('meshLabel').textContent=m.elements;changeProfile();invalidate();errorText('설계를 불러왔습니다. 모드 계산을 실행하세요.','success');}catch(e){errorText('설계 JSON 오류: '+e.message,'error');}};reader.readAsText(f);}
$('preset').addEventListener('change',e=>applyPreset(e.target.value));$('profile').addEventListener('change',()=>{changeProfile();invalidate();});for(const el of document.querySelectorAll('.sidebar input:not([type="file"]),.sidebar textarea,.sidebar select:not(#preset):not(#profile)'))el.addEventListener('input',()=>{$('meshLabel').textContent=$('elements').value;invalidate();});$('calculateBtn').addEventListener('click',()=>calc());$('resetBtn').addEventListener('click',()=>{applyPreset('g654');$('preset').value='g654';calc({autoSweep:true});});$('sweepBtn').addEventListener('click',doSweep);$('meshBtn').addEventListener('click',doMesh);$('jsonBtn').addEventListener('click',exportJSON);$('csvBtn').addEventListener('click',exportCSV);$('pngBtn').addEventListener('click',exportPNG);$('designBtn').addEventListener('click',()=>{try{exportDesign();}catch(e){errorText(e.message,'error');}});$('importDesign').addEventListener('change',e=>{importDesignFile(e.target.files?.[0]);e.target.value='';});
applyPreset('g654');calc({autoSweep:true});
