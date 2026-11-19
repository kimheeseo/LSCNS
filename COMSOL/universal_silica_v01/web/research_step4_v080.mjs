/** STEP4 v0.8.0: result-only physical quantities, verified field-tracked curves,
 * and fixed-index ARROW resonance overlays. NO browser vector FEM or STEP5 Bloch solver.
 */
import {inputsMatch} from './research_design_mesh_v084.mjs?v=20261119-v084';
const $=id=>document.getElementById(id);
const compatible=(a,b)=>{const x=structuredClone(a),y=structuredClone(b);delete x.wavelength_nm;delete y.wavelength_nm;return inputsMatch(x,y);};
const fmt=v=>Number.isFinite(v)?Number(v).toPrecision(5):'TODO';
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;','\'':'&#39;'}[c]));
let metrics=null,sweep=null;
function getJSON(f,maxSize=9000000){
  if(!f||f.size>maxSize)throw Error('파일 크기 제한 또는 빈 파일');
  return f.text().then(t=>JSON.parse(t));
}
function validateMetrics(data){
  if(!['LSCNS_step4_metrics_v080','LSCNS_step4_metrics_v086'].includes(data?.schema)||!Array.isArray(data.modes))throw Error('STEP4 v0.8.0 metrics JSON 아님');
  if(data.modes.some(m=>!Number.isFinite(m.neff_real)||!Number.isFinite(m.alpha_total_dB_m)))throw Error('유한성 오류');
  return data;
}
function validateSweep(data){
  if(!['LSCNS_step4_tracked_dispersion_v080','LSCNS_step4_tracked_sweep_v086'].includes(data?.schema)||!Array.isArray(data.curves))throw Error('overlap 추적 STEP4 데이터 아님');
  for(const c of data.curves){
    if(c.overlap_verified!==true||!Array.isArray(c.wavelength_nm)||c.wavelength_nm.length<5)throw Error('모드 추적 미인증');
    for(const k of ['neff_real','D_ps_nm_km','S_ps_nm2_km'])if(!Array.isArray(c[k])||c[k].length!==c.wavelength_nm.length||c[k].some(v=>!Number.isFinite(v)))throw Error(k+' 데이터 오류');
  }
  return data;
}
function chart(canvas,field,lines=[]){
  const ctx=canvas.getContext('2d');const W=canvas.width,H=canvas.height;
  ctx.fillStyle='#0a1c30';ctx.fillRect(0,0,W,H);
  const series=(sweep?.curves??[]).map(c=>({x:c.wavelength_nm,y:c[field]||c.wavelength_nm.map(()=>null),name:c.mode_id}));
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
  ctx.fillText(({neff_real:'Re(n_eff)',D_ps_nm_km:'D [ps/(nm·km)]',S_ps_nm2_km:'S [ps/(nm²·km)]',alpha_total_dB_m:'α total [dB/m]',P_glass_fraction:'P glass / P total'})[field]||field,bx+11,by+14);
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
    c.x.forEach((x,j)=>{if(!Number.isFinite(c.y[j]))return;if(j===0)ctx.moveTo(px(x),py(c.y[j]));else ctx.lineTo(px(x),py(c.y[j]));});
    ctx.stroke();
    ctx.fillStyle=palette[i%palette.length];ctx.fillText(esc(c.name||'mode '+(i+1)),bx+bw-115,by+36+i*16);
  });
  if(!series.length){ctx.fillStyle='#dccda9';ctx.fillText('ARROW 이론 위치만 표시됨 · FEM 파장 스윕 없음',bx+16,by+bh/2);}
}
function redraw(){
  const lines=sweep?.arrow_overlay?.resonance_lines??metrics?.arrow_overlay?.resonance_lines??metrics?.arrow_overlay?.arrow_lines??[];
  chart($('r4-neff'),'neff_real',lines);
  chart($('r4-D'),'D_ps_nm_km');chart($('r4-S'),'S_ps_nm2_km');chart($('r4-loss'),'alpha_total_dB_m',lines);chart($('r4-glass'),'P_glass_fraction');
  if(metrics){
    const columns=[['mode_index','모드'],['alpha_total_dB_m','α total dB/m'],['Aeff_Et_um2','Aeff µm²'],['MFD_second_moment_Et_um','MFD 2nd µm'],['MFD_PetermannII_um','PII µm'],['P_glass_fraction','유리 파워 비율'],['n2_eff_m2_W','유효 n₂ m²/W'],['gamma_W_inv_m_inv','γ 1/(W·m)']];
    $('r4-table').innerHTML='<table><thead><tr>'+columns.map(([k,t])=>'<th>'+t+'</th>').join('')+'</tr></thead><tbody>'+metrics.modes.map((m,i)=>'<tr>'+columns.map(([k])=>'<td>'+fmt(k==='mode_index'?i+1:m[k])+'</td>').join('')+'</tr>').join('')+'</tbody></table>';
    $('r4-additional').textContent='B='+fmt(metrics.birefringence?.B)+' · L_B[m]='+fmt(metrics.birefringence?.beat_length_m)+' · HOMER='+fmt(metrics.HOMER?.HOMER_ratio)+' · 출처/모드 식별 부족 값은 TODO. '+(sweep?.curves?.map(c=>c.kind||'individual_mode').join(', ')||'');
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
  s.innerHTML='<h2 style="font-size:18px;margin:0 0 9px">STEP 4 · 물리량 계산 및 모드 추적 <small>v0.8.6</small></h2>'+
    '<p style="font-size:12px">정적 웹 뷰어입니다. Python 후처리 JSON만 표시하며, 브라우저에서 FEM이나 신규 밴드갭 계산은 실행하지 않습니다.</p>'+
    '<p style="font-size:12px;color:#eac594">미인증: HCF 손실, HOMER, HCF 유효 n₂, 비축대칭 Petermann-II, surface-mode anticrossing. 출처·모드 데이터가 없으면 TODO로 유지합니다.</p>'+
    '<p><button id="r4-example">실계산 STEP4 검증 예시 적용</button></p><details><summary>공식·단위·사용 조건</summary><p>α=(20/ln10)k₀Im(neff) [dB/m], Aeff=(∫|Et|²)²/∫|Et|⁴ [µm²], MFD₂=2√(2⟨r²⟩), PII=2√(2∫|E|²rdr/∫|E′|²rdr) (축대칭 scalar만), D=−(λ/c)n″ [ps/(nm·km)], S=dD/dλ, γ=2πn₂/(λAeff) [1/(W·m)], B=|nx−ny|, L_B=λ/B [m]. 유리 파워는 Poynting 적분, HCF 유효 n₂는 |Et|⁴ 가중 근사. HOMER=최소 HOM 손실/최대 FM 손실. 출처 없는 n₂와 미확인 모드 식별은 계산을 보류합니다.</p></details>'+
    '<p><label style="display:inline-block;padding:9px;border:1px solid #567d98;border-radius:7px;cursor:pointer">STEP4 메트릭 JSON 선택<input id="r4-metrics" type="file" accept=".json" hidden></label> '+
    '<label style="display:inline-block;padding:9px;border:1px solid #567d98;border-radius:7px;cursor:pointer">overlap 추적 스윕 JSON 선택<input id="r4-sweep" type="file" accept=".json" hidden></label></p>'+
    '<p id="r4-status" role="status" style="font-size:12px">메트릭과 스윕 결과를 가져오세요.</p>'+
    '<p id="r4-additional"></p><div id="r4-table" style="overflow-x:auto"></div>'+
    '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,350px),1fr));gap:12px;margin-top:18px">'+
    '<figure style="margin:0;min-width:0"><figcaption>모드 Re(n_eff) + ARROW 이론 공진선</figcaption><canvas id="r4-neff" width="620" height="280" style="width:100%;max-width:100%"></canvas></figure>'+
    '<figure style="margin:0;min-width:0"><figcaption>Overlap 추적 모드의 분산 D(λ)</figcaption><canvas id="r4-D" width="620" height="280" style="width:100%;max-width:100%"></canvas></figure></div>'+
    '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,300px),1fr));gap:12px">'+['S','loss','glass'].map(k=>'<figure style="margin:0;min-width:0"><figcaption>'+({S:'분산 기울기 S',loss:'FEM 총손실 + ARROW',glass:'유리 파워 비율'})[k]+'</figcaption><canvas id="r4-'+k+'" width="620" height="280" style="width:100%"></canvas></figure>').join('')+'</div>'+
    '<p style="font-size:11px;color:#bdcdda">파장축 ARROW 수직선은 벽 모델의 근사 공진값이며 v0.8.6은 파장 의존 굴절률로 자기일치 근을 구합니다. FEM 곡선은 유효한 overlap 추적 JSON을 가져온 경우에만 표시합니다. 계산된 dB/m는 흡수 포함 총손실 후보입니다. 정량 유효성 미보증.</p>';
  if(target.id==='research-fiber-v1')target.after(s);else target.append(s);
  $('r4-metrics').addEventListener('change',async e=>{
    try{const m=validateMetrics(await getJSON(e.target.files[0]));if(m.input&&!inputsMatch(m.input,JSON.parse($('rf-json').value)))throw Error('현재 설계와 입력 불일치');metrics=m;if(sweep?.input&&m.input&&!compatible(sweep.input,m.input))sweep=null;redraw();}
    catch(err){metrics=null;$('r4-table').replaceChildren();$('r4-status').textContent='메트릭 읽기 실패: '+err.message;}
  });
  $('r4-sweep').addEventListener('change',async e=>{
    try{const q=validateSweep(await getJSON(e.target.files[0]));if(q.input&&!compatible(q.input,JSON.parse($('rf-json').value)))throw Error('현재 설계와 입력 불일치');sweep=q;if(metrics?.input&&q.input&&!compatible(metrics.input,q.input))metrics=null;redraw();}
    catch(err){sweep=null;redraw();$('r4-status').textContent='스윕 읽기 실패: '+err.message;}
  });
  $('rf-json')?.addEventListener('input',()=>{metrics=null;sweep=null;$('r4-table').replaceChildren();$('r4-additional').textContent='입력 변경: 이전 STEP4 결과 무효';redraw();});
  document.addEventListener('ls-step4-import',e=>{try{const m=validateMetrics(e.detail);if(m.input&&!inputsMatch(m.input,JSON.parse($('rf-json').value)))throw Error('현재 설계와 불일치');metrics=m;sweep=null;redraw();}catch(err){$('r4-status').textContent='STEP4 적용 실패: '+err.message;}});
  $('r4-example').addEventListener('click',async()=>{
    try{
      const load=async n=>{const r=await fetch('../fiber_research_v1/benchmarks/step4_v086/'+n,{cache:'no-store'});if(!r.ok)throw Error('HTTP '+r.status);return r.json();};
      const [m,q]=await Promise.all([load('example_metrics.json'),load('example_sweep.json')]);
      $('rf-json').value=JSON.stringify(m.input,null,2);$('rf-json').dispatchEvent(new Event('input',{bubbles:true}));
      metrics=validateMetrics(m);sweep=validateSweep(q);redraw();$('r4-status').textContent+=' · 실제 PEC TE11 합성 검증 예시 (HCF 제품값 아님)';
    }catch(e){$('r4-status').textContent='예시 로드 실패: '+e.message;}
  });
  redraw();
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
