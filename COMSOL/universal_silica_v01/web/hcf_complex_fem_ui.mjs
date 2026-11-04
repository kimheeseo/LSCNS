/** Browser viewer & design export for the separate Python HCF scalar-complex FEM.
 * This DOES NOT run FEM in the browser. Only imported, provenance-matched data.
 * It does not replace the analytical ARROW or existing full-vector solid-fiber panels.
 */
import {buildGeometry,DEFAULT as HCF_DEFAULT} from './hcf_designer.mjs';
const selector='#hcf-designer';
const parent=document.querySelector(selector);
if(parent){
const style=document.createElement('style');
style.textContent=`
#hcf-fem{border:1px solid #3b5d7a;border-radius:16px;background:#101e34;color:#e6f6ff;margin:20px 0 26px;padding:20px;scroll-margin-top:16px}
#hcf-fem *{box-sizing:border-box}#hcf-fem h3{font-size:20px;margin:0 0 8px}#hcf-fem p{font-size:12px;line-height:1.8;color:#a9c5d9}#hcf-fem .fem-tags{display:flex;flex-wrap:wrap;gap:7px;margin:12px 0}#hcf-fem .fem-tag{font-size:10px;color:#cff9f1;border:1px solid #376d7b;padding:5px 9px;border-radius:14px}#hcf-fem .fem-commands{display:flex;gap:8px;flex-wrap:wrap;margin:15px 0}#hcf-fem button,#hcf-fem a.fem-button{padding:9px 12px;background:#0d7688;color:#fff;border:1px solid #48b2c4;border-radius:7px;text-decoration:none;font-size:12px;font-weight:bold}#hcf-fem button{cursor:pointer}#hcf-fem button:hover{background:#148ba0}#hcf-fem input[type=file]{color:#d4e9f3;max-width:100%}#hcf-fem .fem-metrics{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:9px;margin:15px 0}#hcf-fem .fem-metric{background:#192e47;padding:13px;border:1px solid #355471;border-radius:9px}#hcf-fem .fem-metric span{font-size:10px;color:#9eb9d0;display:block}#hcf-fem .fem-metric strong{font-size:18px;display:block;font-variant-numeric:tabular-nums;margin:6px 0;color:#d9f6f8}#hcf-fem .fem-grid{display:grid;grid-template-columns:1fr 1fr;gap:15px}#hcf-fem .fem-box{background:#0b192b;border:1px solid #294862;border-radius:10px;padding:12px;min-width:0}#hcf-fem canvas,#hcf-fem svg{display:block;width:100%;height:auto;border-radius:5px}#hcf-fem h4{font-size:12px;color:#d8eafc;margin:4px 0 10px}#hcf-fem .fem-scroll{max-height:235px;overflow:auto}#hcf-fem table{width:100%;font-size:10px;border-collapse:collapse}#hcf-fem td,#hcf-fem th{padding:7px 5px;border-bottom:1px solid #2e465b;text-align:right}#hcf-fem td:first-child,#hcf-fem th:first-child{text-align:left}#hcf-fem th{color:#81b0c7}#hcf-fem .fem-warning{color:#fae5b0;padding:11px;border:1px solid #856d3c;background:#45392655;border-radius:8px;font-size:11px;line-height:1.8}#hcf-fem .fem-status{font-size:12px;color:#b6e3eb;line-height:1.7}#hcf-fem .fem-subtle{color:#809bb6;font-size:11px}
@media(max-width:920px){#hcf-fem .fem-grid{grid-template-columns:1fr}#hcf-fem .fem-metrics{grid-template-columns:1fr 1fr}#hcf-fem{padding:13px}}
`;document.head.appendChild(style);
const section=document.createElement('section');section.id='hcf-fem';section.innerHTML=`
<h3>HCF FEM — 복소 neff · 공진/반공진 스펙트럼 <small style="color:#80c2cc">Research v0.1</small></h3>
<p>상단 HCF Designer의 현재 코어·튜브 치수로 <b>Python/Colab에서 실제 2D Scalar P1 FEM + 복소 PML</b>을 수행합니다. 본 패널은 설계를 내려받고 FEM 결과를 불러오는 뷰어입니다. <b>브라우저에서 FEM을 즉시 재계산하지 않습니다.</b></p>
<div class="fem-tags"><span class="fem-tag">진짜 FEM 고유값 계산 · Python</span><span class="fem-tag">복소 PML</span><span class="fem-tag">파장별 FEM 스윕</span><span class="fem-tag">Full-vector FEM 아님</span><span class="fem-tag">누설손실 수렴 검증 필요</span></div>
<div class="fem-commands">
<button id="fem-download">① 현재 HCF 설계 JSON 저장</button>
<a class="fem-button" target="_blank" rel="noopener noreferrer" href="https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/universal_silica_v01/hcf_complex_fem/HCF_Complex_FEM_Research.ipynb">② Colab에서 FEM 실행</a>
<label style="font-size:11px;color:#b8d6e7;align-self:center">③ 결과 JSON 불러오기 <input id="fem-import" type="file" accept=".json,application/json"/></label>
<button id="fem-example" type="button">저장된 FEM 예시 보기</button>
</div>
<p id="fem-status" class="fem-status" role="status" aria-live="polite">현재 설계를 JSON으로 저장한 다음 Colab에서 FEM을 실행하고 계산 결과를 다시 가져오세요.</p>
<div class="fem-warning" id="fem-warning">주의: 이 모듈의 dB/km 값은 전부 **수렴 미검증 예비 FEM 수치**입니다. 이를 인증 손실 또는 정확한 반공진 파장으로 해석하지 마세요. 공진·반공진 후보는 FEM 파장 스윕에서 검출한 국소 극값입니다.</div>
<div class="fem-metrics">
 <div class="fem-metric"><span>Re(neff) · Scalar FEM</span><strong id="fem-re">—</strong></div>
 <div class="fem-metric"><span>Im(neff) · Scalar FEM/PML</span><strong id="fem-im">—</strong></div>
 <div class="fem-metric"><span>Air core power fraction</span><strong id="fem-core">—</strong></div>
 <div class="fem-metric"><span>Confinement loss (예비 · 미검증)</span><strong id="fem-loss">—</strong></div>
</div>
<div class="fem-grid"><div class="fem-box"><h4>파장 스윕: 복소 FEM/PML 기반 예비 누설손실 변화</h4><div id="fem-spectrum">파장별 FEM 결과를 불러오세요.</div><p>국소 최대/최소는 공진·반공진 <b>후보</b>입니다. FEM 결과에 대한 PML·메시·파장 스텝 수렴 검증이 필요합니다.</p></div>
<div class="fem-box"><h4>실제 FEM 스칼라 모드 단면 |ψ(x,y)|²</h4><canvas id="fem-field" width="520" height="420"></canvas><p id="fem-field-note">결과 JSON에 포함된 실수부/허수부 모드 데이터를 사용합니다.</p></div></div>
<div class="fem-grid" style="margin-top:14px"><div class="fem-box"><h4>FEM 공진/반공진 후보 (미검증)</h4><div class="fem-scroll" id="fem-features">결과가 없습니다.</div></div>
<div class="fem-box"><h4>모드 식별 및 수렴 경고</h4><div id="fem-diagnostics" class="fem-subtle">독립적인 mesh/PML/domain 수렴이 필요합니다.</div></div></div>
<p class="fem-subtle">계산: 2D Scalar FEM, 삼각형 P1 요소, Cartesian complex-coordinate PML, PML 외곽 Dirichlet. 벽 경계점 샘플링 기반 Delaunay 메시(엄밀한 경계 정합 메시 아님). Full-vector HE11, 신뢰할 수 있는 제품 손실, 산란·흡수·굽힘손실 및 3D 전파는 미포함.</p>`;
parent.after(section);
const q=id=>section.querySelector('#'+id);
const geomKeys=['coreDiameterUm','tubeCount','nesting','outerTubeRatio','innerRatio','thirdRatio','gapCoreUm','minNeighborGapUm','nestedClearanceUm','wall1Um','wall2Um','wall3Um','wavelengthNm','airIndex'];
const nf=(x,n=6)=>typeof x==='number'&&Number.isFinite(x)?x.toFixed(n):'—';
const save=(obj,name)=>{const b=new Blob([JSON.stringify(obj,null,2)],{type:'application/json'}),url=URL.createObjectURL(b),a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);};
function getCurrent(){
 const d={...HCF_DEFAULT};
 for(const key of geomKeys){const el=document.getElementById('hcf-'+key);if(el)d[key]=Number(el.value);}
 const flag=document.getElementById('hcf-autoFit');d.autoFit=flag?flag.checked:true;
 buildGeometry(d);// rejects invalid geometry
 return d;
}
function sameGeometry(a,b){return geomKeys.every(k=>Math.abs(Number(a[k])-Number(b[k]))<Math.max(1e-7,Math.abs(Number(a[k]))*1e-7))&&Boolean(a.autoFit)===Boolean(b.autoFit);}
let output=null;
function updateStatus(msg){q('fem-status').textContent=msg;}
function clearResults(reason){output=null;for(const id of ['re','im','core','loss'])q('fem-'+id).textContent='—';q('fem-spectrum').textContent='FEM 재실행 후 결과 JSON을 불러오세요.';q('fem-features').textContent='결과 없음';q('fem-diagnostics').textContent=reason;paintField(null);updateStatus(reason);}
function paintField(item){
 const c=q('fem-field'),ctx=c.getContext('2d'),W=c.width,H=c.height;
 ctx.fillStyle='#091626';ctx.fillRect(0,0,W,H);
 if(!item?.mode_points?.length){ctx.fillStyle='#a0c4d9';ctx.font='15px sans-serif';ctx.fillText('FEM 복소 모드 데이터 대기',20,33);return;}
 const g=item.geometry,extent=(g.radius_extent_um+item.input.domain_buffer_um)*1.05,scale=Math.min(W,H)*.435/extent,cx=W/2,cy=H/2;
 const points=item.mode_points,peak=Math.max(1e-30,...points.map(v=>v.intensity));
 for(const p of points){const f=Math.sqrt(Math.min(1,p.intensity/peak));if(!Number.isFinite(f))continue;const x=cx+p.x_um*scale,y=cy-p.y_um*scale;ctx.fillStyle=`rgba(${Math.round(18+240*f)},${Math.round(45+165*f*f)},${Math.round(115+100*(1-f))},${.18+.82*f})`;ctx.fillRect(x-1.3,y-1.3,2.6,2.6);}
 ctx.lineWidth=1.2;ctx.strokeStyle='#51e0cb';ctx.setLineDash([5,4]);ctx.beginPath();ctx.arc(cx,cy,g.core_radius_um*scale,0,Math.PI*2);ctx.stroke();ctx.setLineDash([]);
 for(const tube of g.tubes){ctx.strokeStyle=['#54bcf4','#e6cb84','#d9a5df'][tube.layer-1]||'#fff';ctx.lineWidth=.8;for(const rr of [tube.r,tube.r-tube.wall]){ctx.beginPath();ctx.arc(cx+tube.x*scale,cy-tube.y*scale,rr*scale,0,Math.PI*2);ctx.stroke();}}
}
function spectrumSvg(results){
 const vals=results.map(o=>[o.input?.wavelengthNm,o.confinement_loss_dB_km_unverified,o.mode_identification_pass]).filter(([l,v,ok])=>Number.isFinite(l)&&Number.isFinite(v)&&v>0&&ok);
 if(vals.length<2)return '<p>유효한 스펙트럼 점이 부족합니다.</p>';
 const logs=vals.map(x=>Math.log10(x[1]));const lo=Math.min(...logs)-.2,hi=Math.max(...logs)+.2;
 const lamMin=Math.min(...vals.map(x=>x[0])),lamMax=Math.max(...vals.map(x=>x[0]));
 const w=660,h=270,left=61,right=635,top=20,bottom=227;
 const X=l=>left+(right-left)*(l-lamMin)/Math.max(1,lamMax-lamMin);
 const Y=v=>bottom-(bottom-top)*(Math.log10(v)-lo)/Math.max(.001,hi-lo);
 const p=vals.map((x,i)=>`${i===0?'M':'L'}${X(x[0]).toFixed(1)},${Y(x[1]).toFixed(1)}`).join(' ');
 const circles=vals.map(v=>`<circle cx="${X(v[0]).toFixed(1)}" cy="${Y(v[1]).toFixed(1)}" r="2.5" fill="#6addd1"/>`).join('');
 const ticks=Array.from({length:5},(_,i)=>{const logval=lo+i*(hi-lo)/4;const yy=bottom-(bottom-top)*i/4;return `<line x1="${left}" y1="${yy}" x2="${right}" y2="${yy}" stroke="#28455e"/><text x="${left-7}" y="${yy+4}" fill="#a3c2d8" font-size="12" text-anchor="end">10^${logval.toFixed(1)}</text>`}).join('');
 return `<svg viewBox="0 0 ${w} ${h}" role="img" aria-label="FEM preliminary leakage versus wavelength"><rect width="${w}" height="${h}" fill="#0b192b"/>${ticks}<path d="${p}" fill="none" stroke="#66d6d1" stroke-width="2.1"/>${circles}<text x="${(left+right)/2}" y="259" fill="#c5e0ef" font-size="13" text-anchor="middle">λ (nm): ${lamMin}–${lamMax}</text><text x="${left}" y="13" fill="#aec9d9" font-size="11">예비 loss dB/km · log₁₀ · 검증되지 않음</text></svg>`;
}
function show(out){
 const raw=Array.isArray(out?.results)?out.results:[out];
 if(!raw.length||raw.some(x=>!x||!Number.isFinite(x.neff_real)||!Number.isFinite(x.neff_imag)||!x.input||!x.geometry))throw Error('FEM 결과 JSON 형식이 아닙니다.');
 let design;try{design=getCurrent();}catch(e){throw Error('상단 HCF 설계 입력 오류: '+e.message);}
 const match=raw.every(x=>sameGeometry(design,x.input));
 if(!match)throw Error('상단 HCF 설계와 결과 JSON의 튜브·코어 치수가 다릅니다. 설계를 원래 계산 조건으로 돌린 뒤 다시 불러오세요.');
 output=out;
 const desired=design.wavelengthNm;
 const selected=raw.reduce((a,b)=>Math.abs(b.input.wavelengthNm-desired)<Math.abs(a.input.wavelengthNm-desired)?b:a);
 q('fem-re').textContent=nf(selected.neff_real,9);
 q('fem-im').textContent=selected.neff_imag.toExponential(3);
 q('fem-core').textContent=(100*selected.air_core_power_fraction).toFixed(2)+'%';
 q('fem-loss').textContent=Number.isFinite(selected.confinement_loss_dB_km_unverified)?nf(selected.confinement_loss_dB_km_unverified,3)+'*':'미산출';
 q('fem-spectrum').innerHTML=spectrumSvg(raw);
 const extrema=(out.extrema_candidates||[]).filter(v=>Number.isFinite(v.wavelength_nm));
 const header='<table><thead><tr><th>유형</th><th>λ (nm)</th><th>f (THz)</th></tr></thead><tbody>';
 q('fem-features').innerHTML=extrema.length?header+extrema.slice(0,32).map(v=>'<tr><td>'+(v.kind?.startsWith('resonance_')?'공진 피크 후보':'반공진 딥 후보')+'</td><td>'+nf(v.wavelength_nm,1)+'</td><td>'+nf(v.frequency_THz,2)+'</td></tr>').join('')+'</tbody></table>':'FEM 국소 극값이 없거나 모드 식별 불안정으로 제외했습니다.';
 q('fem-diagnostics').textContent='계산 파장 '+selected.input.wavelengthNm+' nm · 모드 J0 중첩 '+nf(selected.ideal_J0_overlap,3)+' · FEM 메시 '+selected.mesh.nodes+' 노드 / '+selected.mesh.elements+' 요소 · 모드 식별 '+(selected.mode_identification_pass?'통과':'불안정')+' · Loss 수렴 검증 '+(selected.converged_and_validated?'완료':'미완료')+' · Full-vector 미구현. '+(Math.abs(selected.input.wavelengthNm-desired)>.01?'현재 HCF 목표 파장과 계산 결과 파장이 다릅니다. ':'');
 const field=raw.find(x=>x.mode_points?.length);paintField(field||null);
 q('fem-field-note').textContent=field?`실제 계산된 Scalar FEM 모드 · λ=${field.input.wavelengthNm} nm · 복소 모드 강도(정규화). 점 좌표 표시용 다운샘플링.`:'모드장 데이터가 없습니다. Colab에서 field export 옵션을 켜세요.';
 updateStatus(`실제 Python Scalar FEM 결과 ${raw.length}개 파장 표시 · ${match?'현재 상단 HCF 설계 치수와 일치':''} · 수렴 미검증. 표의 *는 예비값입니다.`);
}
q('fem-download').onclick=()=>{try{const d=getCurrent();save({...d,background_h_um:1.8,wall_edge_h_um:.65,domain_buffer_um:5,pml_thickness_um:8,pml_strength:2,pml_exponent:3,n_modes:6,max_nodes:28000},'HCF_design_for_scalar_complex_FEM.json');updateStatus('현재 HCF 설계 JSON을 저장했습니다. Colab에서 업로드한 뒤 실제 FEM을 실행하세요.');}catch(e){updateStatus('설계 오류: '+e.message);}};
q('fem-import').addEventListener('change',async e=>{const file=e.target.files?.[0];if(!file)return;try{const obj=JSON.parse(await file.text());show(obj);}catch(error){clearResults('불러오기 실패: '+error.message);}finally{e.target.value='';}});
q('fem-example').onclick=async()=>{try{const r=await fetch('./hcf_complex_fem/results/default_sweep.json');if(!r.ok)throw Error('GitHub Pages 배포 파일 확인 필요');show(await r.json());}catch(e){clearResults('예시 불러오기 실패: '+e.message);}};
parent.addEventListener('input',e=>{if(e.target.id==='hcf-external-im')return;if(!output)return;let changed=false;try{changed=!sameGeometry(getCurrent(),(output.results||[output])[0].input);}catch{changed=true;}if(changed)clearResults('설계 치수 변경: 기존 FEM 결과를 무효화했습니다. 새 설계로 Colab에서 다시 계산하세요.');});
}
