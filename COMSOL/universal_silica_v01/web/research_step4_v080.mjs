/** STEP4 v0.8.0: result-only physical quantities, verified field-tracked curves,
 * and fixed-index ARROW resonance overlays. NO browser vector FEM or STEP5 Bloch solver.
 */
const $=id=>document.getElementById(id);
const fmt=v=>Number.isFinite(v)?Number(v).toPrecision(5):'TODO';
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;','\'':'&#39;'}[c]));
let metrics=null,sweep=null;
function getJSON(f,maxSize=9000000){
  if(!f||f.size>maxSize)throw Error('파일 크기 제한 또는 빈 파일');
  return f.text().then(t=>JSON.parse(t));
}
function validateMetrics(data){
  if(data?.schema!=='LSCNS_step4_metrics_v080'||data.engine_version!=='0.8.0'||!Array.isArray(data.modes))throw Error('STEP4 v0.8.0 metrics JSON 아님');
  if(data.modes.some(m=>!Number.isFinite(m.neff_real)||!Number.isFinite(m.alpha_total_dB_m)))throw Error('유한성 오류');
  return data;
}
function validateSweep(data){
  if(data?.schema!=='LSCNS_step4_tracked_dispersion_v080'||!Array.isArray(data.curves))throw Error('overlap 추적 STEP4 데이터 아님');
  for(const c of data.curves){
    if(c.overlap_verified!==true||!Array.isArray(c.wavelength_nm)||c.wavelength_nm.length<5)throw Error('모드 추적 미인증');
    for(const k of ['neff_real','D_ps_nm_km','S_ps_nm2_km'])if(!Array.isArray(c[k])||c[k].length!==c.wavelength_nm.length||c[k].some(v=>!Number.isFinite(v)))throw Error(k+' 데이터 오류');
  }
  return data;
}
function chart(canvas,field,lines=[]){
  const ctx=canvas.getContext('2d');const W=canvas.width,H=canvas.height;
  ctx.fillStyle='#0a1c30';ctx.fillRect(0,0,W,H);
  const series=(sweep?.curves??[]).map(c=>({x:c.wavelength_nm,y:c[field],name:c.mode_id}));
  const xdata=[...series.flatMap(c=>c.x),...lines.map(l=>l.lambda_nm)].filter(Number.isFinite);
  const ydata=series.flatMap(c=>c.y).filter(Number.isFinite);
  const bx=57,by=22,bw=W-94,bh=H-65;
  ctx.strokeStyle='#7d97a9';ctx.lineWidth=1;ctx.strokeRect(bx,by,bw,bh);
  if(!xdata.length){ctx.fillStyle='#d4e9f8';ctx.fillText('결과 파일을 가져오면 그래프가 표시됩니다.',bx+20,by+30);return;}
  let xmin=Math.min(...xdata),xmax=Math.max(...xdata);
  if(xmax<=xmin){xmin-=5;xmax+=5;}
  const px=x=>bx+(x-xmin)/(xmax-xmin)*bw;
  let ymin=Math.min(...ydata),ymax=Math.max(...ydata);
  if(!ydata.length){ymin=0;ymax=1;}
  if(ymax<=ymin){ymin-=1;ymax+=1;}
  const pad=(ymax-ymin)*.09;ymin-=pad;ymax+=pad;
  const py=y=>by+(ymax-y)/(ymax-ymin)*bh;
  ctx.fillStyle='#d3e7f1';ctx.font='11px sans-serif';
  for(let k=0;k<5;k++){
    const x=xmin+(xmax-xmin)*k/4;
    ctx.fillText(x.toFixed(0),px(x)-9,by+bh+19);
    const y=ymin+(ymax-ymin)*k/4;
    ctx.fillText(y.toPrecision(4),4,py(y)+4);
  }
  ctx.fillText('파장 [nm]',bx+bw/2,by+bh+40);
  ctx.fillText(field==='neff_real'?'Re(n_eff) vs λ':'D [ps/(nm·km)] vs λ',bx+11,by+14);
  // ARROW lines: only constant index formula, NOT FEM loss peaks.
  ctx.save();ctx.setLineDash([6,5]);ctx.strokeStyle='#ebb776';ctx.lineWidth=1.5;
  lines.forEach(l=>{
    const x=px(l.lambda_nm);
    ctx.beginPath();ctx.moveTo(x,by);ctx.lineTo(x,by+bh);ctx.stroke();
    ctx.fillStyle='#ffd6a7';ctx.fillText('ARROW m='+l.m,x+3,by+27);
  });ctx.restore();
  const palette=['#74daf2','#c5a7ff','#86dc9b','#f2a8bb','#bade9e','#ffca7b'];
  series.forEach((c,i)=>{
    ctx.beginPath();ctx.lineWidth=2;ctx.strokeStyle=palette[i%palette.length];
    c.x.forEach((x,j)=>{if(j===0)ctx.moveTo(px(x),py(c.y[j]));else ctx.lineTo(px(x),py(c.y[j]));});
    ctx.stroke();
    ctx.fillStyle=palette[i%palette.length];ctx.fillText(esc(c.name||'mode '+(i+1)),bx+bw-115,by+36+i*16);
  });
  if(!series.length){ctx.fillStyle='#dccda9';ctx.fillText('ARROW 이론 위치만 표시됨 · FEM 파장 스윕 없음',bx+16,by+bh/2);}
}
function redraw(){
  const lines=metrics?.arrow_overlay?.arrow_lines??[];
  chart($('r4-neff'),'neff_real',lines);
  chart($('r4-D'),'D_ps_nm_km');
  if(metrics){
    const info=metrics.modes.map((m,i)=>
      '<tr><td>'+Number(i+1)+'</td><td>'+fmt(m.alpha_total_dB_m)+'</td><td>'+fmt(m.Aeff_Et_um2)+'</td><td>'+fmt(m.MFD_second_moment_Et_um)+'</td><td>'+fmt(m.gamma_W_inv_m_inv)+'</td></tr>').join('');
    $('r4-table').innerHTML='<table style="width:100%;border-collapse:collapse;overflow-wrap:anywhere;font-size:12px"><thead><tr><th>모드</th><th>총손실 후보 dB/m</th><th>Aeff µm² (Et)</th><th>MFD 2nd µm</th><th>γ 1/(W·m)</th></tr></thead><tbody>'+info+'</tbody></table>';
  }
  const m=metrics?'메트릭 '+metrics.modes.length+'개 모드':'메트릭 미입력';
  const sw=sweep?'인접 overlap 추적 '+sweep.curves.length+'개 모드':'모드 추적 스윕 미입력';
  $('r4-status').textContent=m+' / '+sw+' · FEM/PML 손실의 정량 검증은 아직 별도로 필요합니다.';
}
function init(){
  const target=$('research-fiber-v1')||document.querySelector('main.workspace')||document.querySelector('main');
  if(!target||$('research-step4-v080'))return;
  const s=document.createElement('section');
  s.id='research-step4-v080';s.className='engineering-panel';s.style.cssText='margin:22px 0;padding:18px;background:#11243c;color:#e4f3ff;border:1px solid #3f6f89;border-radius:12px';
  s.innerHTML='<h2 style="font-size:18px;margin:0 0 9px">STEP 4 · 물리량 계산 및 모드 추적 <small>v0.8.0</small></h2>'+
    '<p style="font-size:12px">정적 웹 뷰어입니다. Python 후처리 JSON만 표시하며, 브라우저에서 FEM이나 신규 밴드갭 계산은 실행하지 않습니다.</p>'+
    '<p style="font-size:12px;color:#eac594">미인증: HCF 손실, HOMER, HCF 유효 n₂, 비축대칭 Petermann-II, surface-mode anticrossing. 출처·모드 데이터가 없으면 TODO로 유지합니다.</p>'+
    '<p><label style="display:inline-block;padding:9px;border:1px solid #567d98;border-radius:7px;cursor:pointer">STEP4 메트릭 JSON 선택<input id="r4-metrics" type="file" accept=".json" hidden></label> '+
    '<label style="display:inline-block;padding:9px;border:1px solid #567d98;border-radius:7px;cursor:pointer">overlap 추적 스윕 JSON 선택<input id="r4-sweep" type="file" accept=".json" hidden></label></p>'+
    '<p id="r4-status" role="status" style="font-size:12px">메트릭과 스윕 결과를 가져오세요.</p>'+
    '<div id="r4-table" style="overflow-x:auto"></div>'+
    '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,350px),1fr));gap:12px;margin-top:18px">'+
    '<figure style="margin:0;min-width:0"><figcaption>모드 Re(n_eff) + ARROW 이론 공진선</figcaption><canvas id="r4-neff" width="620" height="280" style="width:100%;max-width:100%"></canvas></figure>'+
    '<figure style="margin:0;min-width:0"><figcaption>Overlap 추적 모드의 분산 D(λ)</figcaption><canvas id="r4-D" width="620" height="280" style="width:100%;max-width:100%"></canvas></figure></div>'+
    '<p style="font-size:11px;color:#bdcdda">파장축 ARROW 수직선은 고정 굴절률 근사 공진값입니다. FEM 곡선은 유효한 overlap 추적 JSON을 가져온 경우에만 표시합니다. 계산된 dB/m는 흡수 포함 총손실 후보입니다. 정량 유효성 미보증.</p>';
  target.after(s);
  $('r4-metrics').addEventListener('change',async e=>{
    try{metrics=validateMetrics(await getJSON(e.target.files[0]));redraw();}
    catch(err){$('r4-status').textContent='메트릭 읽기 실패: '+err.message;}
  });
  $('r4-sweep').addEventListener('change',async e=>{
    try{sweep=validateSweep(await getJSON(e.target.files[0]));redraw();}
    catch(err){$('r4-status').textContent='스윕 읽기 실패: '+err.message;}
  });
  redraw();
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
