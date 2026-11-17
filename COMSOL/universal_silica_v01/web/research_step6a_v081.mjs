/** v0.8.1 STEP6a ONLY — 2D conformal macro-bend result viewer.
 * No FEM is executed in browser. Loss is experimental / NOT COMSOL V3 certified.
 */
const $=id=>document.getElementById(id);
const finite=v=>typeof v==='number'&&Number.isFinite(v);
const num=v=>finite(v)?v.toPrecision(6):'근거 필요';
const metrics={
  alpha_total_dB_m:{label:'총손실 α [dB/m]',axis:'α [dB/m]'},
  delta_total_dB_m:{label:'직선 대비 총손실 변화 Δα [dB/m]',axis:'Δα [dB/m]'},
  loss_per_turn_total_dB:{label:'1회 굽힘 총손실 추정 [dB/turn]',axis:'dB/turn'}
};
let file=null,stale=true;
function sortKeys(value){
  if(Array.isArray(value))return value.map(sortKeys);
  if(value&&typeof value==='object')return Object.fromEntries(
    Object.keys(value).sort().map(k=>[k,sortKeys(value[k])]));
  return value;
}
function currentDesignMatches(){
  const editor=$('rf-json');
  if(!editor||!file)return false;
  try{
    return JSON.stringify(sortKeys(JSON.parse(editor.value)))===
           JSON.stringify(sortKeys(file.source_design));
  }catch{return false;}
}
function textStatus(msg){const p=$('r6-status');if(p)p.textContent=msg;}
function paint(){
  const canvas=$('r6-plot');if(!canvas)return;
  const ctx=canvas.getContext('2d'),W=canvas.width,H=canvas.height;
  ctx.fillStyle='#071b2e';ctx.fillRect(0,0,W,H);
  const B={x:76,y:27,w:W-107,h:H-91};
  ctx.strokeStyle='#64889c';ctx.strokeRect(B.x,B.y,B.w,B.h);
  ctx.fillStyle='#d9e8f0';ctx.font='13px sans-serif';
  if(!file||stale){
    ctx.fillText(stale&&file?'설계 JSON이 달라 결과 무효화':'STEP 6a bend_sweep_result.json을 선택하세요.',B.x+14,B.y+40);
    return;
  }
  const field=$('r6-metric').value;
  const byDirection=new Map();
  file.cases.forEach(c=>{
    if(c.status!=='COMPUTED_UNVERIFIED'||!finite(c[field])||!finite(c.radius_mm))return;
    const key=Number(c.direction_deg);
    if(!byDirection.has(key))byDirection.set(key,[]);
    byDirection.get(key).push({x:c.radius_mm,y:c[field]});
  });
  const data=[...byDirection.values()].flat();
  if(!data.length){
    ctx.fillStyle='#f7cea0';ctx.fillText('Overlap을 통과한 FEM 모드 결과가 없습니다.',B.x+12,B.y+40);
    return;
  }
  let xmin=Math.min(...data.map(d=>d.x)),xmax=Math.max(...data.map(d=>d.x));
  let ymin=Math.min(...data.map(d=>d.y)),ymax=Math.max(...data.map(d=>d.y));
  if(xmin===xmax){xmin-=.05*xmin||1;xmax+=.05*xmax||1;}
  if(ymin===ymax){const pad=Math.max(Math.abs(ymin)*.05,1e-10);ymin-=pad;ymax+=pad;}
  const pad=(ymax-ymin)*.07;ymin-=pad;ymax+=pad;
  const px=x=>B.x+(x-xmin)/(xmax-xmin)*B.w;
  const py=y=>B.y+B.h-(y-ymin)/(ymax-ymin)*B.h;
  ctx.fillStyle='#d9e8f0';ctx.font='11px sans-serif';
  for(let i=0;i<=4;i++){
    const r=xmin+(xmax-xmin)*i/4;
    ctx.fillText(r.toPrecision(3),px(r)-12,B.y+B.h+22);
    const v=ymin+(ymax-ymin)*i/4;
    ctx.fillText(v.toExponential(2),8,py(v)+4);
  }
  ctx.font='12px sans-serif';
  ctx.fillText('굽힘 반경 R [mm]',B.x+B.w/2-47,B.y+B.h+44);
  ctx.fillText(metrics[field].axis,B.x+6,B.y+16);
  const palette=['#74daf2','#f6bd80','#92e1a7','#d5b5ff','#fab4cc','#ffe49d'];
  let colorIdx=0;
  for(const [direction,pts] of byDirection){
    pts.sort((a,b)=>a.x-b.x);
    const col=palette[colorIdx%palette.length];
    ctx.strokeStyle=col;ctx.fillStyle=col;ctx.lineWidth=2;
    ctx.beginPath();pts.forEach((p,i)=>i?ctx.lineTo(px(p.x),py(p.y)):ctx.moveTo(px(p.x),py(p.y)));
    ctx.stroke();
    pts.forEach(p=>{ctx.beginPath();ctx.arc(px(p.x),py(p.y),3.7,0,2*Math.PI);ctx.fill();});
    ctx.fillText('방향 '+direction+'°',B.x+B.w-110,B.y+20+colorIdx*17);
    colorIdx++;
  }
}
function rebuild(){
  if(!file){paint();return;}
  stale=!currentDesignMatches();
  const total=file.cases.length;
  const good=file.cases.filter(c=>c.status==='COMPUTED_UNVERIFIED').length;
  const entries=file.cases.filter(c=>c.status!=='COMPUTED_UNVERIFIED')
    .map(c=>c.radius_mm+'mm/'+c.direction_deg+'°: '+c.status).slice(0,6);
  $('r6-metadata').textContent='엔진 '+file.engine_version+' · 입력 SHA-256 '+file.input_sha256+
    ' · '+good+'/'+total+' overlap 통과 후보 · V3: '+(file.validation?.V3||'EVIDENCE_REQUIRED')+
    ' · '+(file.stress_correction_requested?'문헌 명시 응력 보정':'응력 보정 없음');
  $('r6-failures').textContent=entries.length?'실패/미검증: '+entries.join(' / '):'';
  textStatus(stale?'설계 JSON과 결과의 원본 설계가 다릅니다. 현재 결과 무효 — [결과 설계 적용]을 선택하세요.':
      '입력 일치 · 총손실 후보값이며 실제 bend radiation loss 인증 아님 (V3/V4/V5 미검증)');
  paint();
}
function validate(data){
  if(data?.schema!=='LSCNS_step6a_bend_sweep_v081'||data.engine_version!=='0.8.1'||
     !Array.isArray(data.cases)||!data.source_design||!Array.isArray(data.radii_mm_descending)||
     !/^[0-9a-f]{64}$/i.test(data.input_sha256||'')){
    throw Error('STEP6a v0.8.1 결과 JSON이 아닙니다');
  }
  const d=data.source_design;
  if(!d.family||!d.parameters||!d.boundary||!d.mesh)throw Error('원본 설계 누락');
  if(data.cases.some(c=>!finite(c.radius_mm)||!finite(c.direction_deg)||!c.status))throw Error('손실 사례 형식 오류');
  return data;
}
function init(){
  const anchor=$('research-step4-v080')||$('research-fiber-v1')||document.querySelector('main');
  if(!anchor||$('research-step6a-v081'))return;
  const section=document.createElement('section');
  section.id='research-step6a-v081';section.className='engineering-panel';
  section.style.cssText='margin:22px 0;padding:18px;border:1px solid #44758c;background:#10273c;color:#edf7ff;border-radius:11px';
  section.innerHTML='<h2 style="font-size:18px">STEP 6a · 굽힘 손실 해석 <small>v0.8.1 · 실험용</small></h2>'+
    '<p style="font-size:12px">등가 굴절률 n_eq = n(x,y) × [1 + (x cosθ + y sinθ)/R_eff] · R_eff = 문헌 출처가 있는 값으로만 선택 보정.</p>'+
    '<p style="font-size:12px;color:#f2c98b">현재 Python 2D conformal FEM 기반이며 COMSOL V3 축대칭 검증 및 PML/메시 수렴은 미통과. 수치는 후보 결과로만 사용하세요.</p>'+
    '<div style="display:flex;gap:9px;flex-wrap:wrap;align-items:center">'+
    '<label style="display:inline-block;border:1px solid #6c97af;border-radius:6px;padding:8px;cursor:pointer">굽힘 결과 JSON 열기<input type="file" accept=".json" id="r6-open" hidden></label>'+
    '<button id="r6-adopt" style="border:1px solid #6c97af;border-radius:6px;padding:8px;background:#1f4663;color:#e8f5ff;cursor:pointer">결과 설계 적용</button>'+
    '<label>표시 지표 <select id="r6-metric" style="background:#173b53;color:#edf7ff;padding:7px;border:1px solid #6c97af">'+
      Object.entries(metrics).map(([k,v])=>'<option value="'+k+'">'+v.label+'</option>').join('')+
    '</select></label></div>'+
    '<p id="r6-status" role="status" style="font-size:12px">굽힘 결과 파일을 선택하세요.</p>'+
    '<canvas id="r6-plot" width="770" height="360" style="width:100%;max-width:900px;height:auto;display:block"></canvas>'+
    '<p id="r6-metadata" style="font-size:11px;overflow-wrap:anywhere"></p>'+
    '<p id="r6-failures" style="font-size:11px;color:#f3cead"></p>'+
    '<details style="font-size:12px"><summary style="cursor:pointer">계산 근거 보기 · 단위 / 출처 / 검증 한계</summary>'+
    '<p>R[mm] × 1000 = R[µm], x_out=x cosθ+y sinθ [µm], ε_eq=ε·(1+x_out/R_eff)². '+
    '총 α=(20/ln10)k₀ Im(n_eff) [dB/m]; 전체 1회 굽힘의 추정 총손실=α·2πR[mm]·10⁻³ [dB]. '+
    '직선 대비 Δα=α_bend−α_straight이며 부호는 변경하지 않습니다.</p>'+
    '<p><a href="https://www.comsol.com/paper/study-of-bending-losses-in-optical-fibers-using-comsol-multiphysics-61111" target="_blank" rel="noopener">2D 등각 변환 참고</a> · '+
    '<a href="https://optiwave.wiki/kb/optical_fiber_bending_loss_with_stress_effect/" target="_blank" rel="noopener">응력 R_eff 문헌</a> · '+
    '<a href="https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.step_index_fiber_bend/step_index_fiber_bend.html" target="_blank" rel="noopener">COMSOL V3 기준 예제</a>. '+
    'V3는 다른 축대칭 해석 방식이므로 수치 근거 없이는 PASS가 아닙니다.</p></details>';
  anchor.after(section);
  $('r6-open').addEventListener('change',async e=>{
    try{
      const f=e.target.files[0];if(!f||f.size>12000000)throw Error('12MB 제한/파일 없음');
      file=validate(JSON.parse(await f.text()));rebuild();
    }catch(err){file=null;stale=true;textStatus('파일 불러오기 실패: '+err.message);paint();}
  });
  $('r6-adopt').addEventListener('click',()=>{
    if(!file){textStatus('먼저 굽힘 결과 JSON을 선택하세요.');return;}
    const ed=$('rf-json');
    if(!ed){textStatus('기존 설계 편집기를 찾을 수 없습니다.');return;}
    ed.value=JSON.stringify(file.source_design,null,2);
    ed.dispatchEvent(new Event('input',{bubbles:true}));
    rebuild();
  });
  $('r6-metric').addEventListener('change',paint);
  $('rf-json')?.addEventListener('input',rebuild);
  paint();
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);
else init();
