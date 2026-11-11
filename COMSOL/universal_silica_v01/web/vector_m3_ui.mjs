/**
 * M3 browser full-vector study, separate from existing Colab JSON viewer.
 * Feature flag: window.LSWaveFeatures.vectorM3=false before module import.
 * Worker is static-GitHub-Pages-compatible; no server/API/WASM dependency.
 */
const ENABLED=globalThis.LSWaveFeatures?.vectorM3!==false && globalThis.LSWaveFeatures?.fullVectorWeb!==false;
let worker=null,latest=null,busy=false,sequence=0;
const $=id=>document.getElementById(id);
const fmt=(n,d=8)=>Number.isFinite(n)?n.toFixed(d):"—";
function draw(canvas,matrix){
 if(!canvas||!Array.isArray(matrix))return;
 const n=matrix.length,ctx=canvas.getContext("2d"),size=canvas.width;
 const max=Math.max(1e-30,...matrix.flat().map(Number));
 ctx.fillStyle="#071527";ctx.fillRect(0,0,size,size);
 const pixel=size/n;
 for(let j=0;j<n;j++)for(let i=0;i<n;i++){
  const ratio=Math.pow(Math.max(0,Math.min(1,Number(matrix[j][i])/max)),.55);
  ctx.fillStyle="rgb("+Math.round(10+235*ratio)+","+Math.round(32+160*ratio)+","+Math.round(50+70*(1-ratio))+")";
  ctx.fillRect(i*pixel,j*pixel,pixel+1,pixel+1);
 }
}
function init(){
 if(!ENABLED||$("ls-m3"))return;
 const parent=$("full-vector-tool");if(!parent)return;
 const section=document.createElement("section");section.id="ls-m3";section.className="ls-m3-panel";
 section.innerHTML='<h3>브라우저 내 Full-vector FDFD · M3 v0.5</h3>'+
 '<p>실제 결합 Hx·Hy 2D Maxwell 고유모드 계산 → Ex, Ey, Ez 복원. 기존 Colab JSON 불러오기 기능은 위에 그대로 남아 있습니다. 아래 해석은 PML/복소 굴절률/굽힘손실을 포함하지 않습니다.</p>'+
 '<div class="ls-m3-controls">'+
 '<label>Wavelength nm<input id="m3-lambda" type="number" value="1550" min="401" max="1699"/></label>'+
 '<label>Core radius µm<input id="m3-radius" type="number" value="4.1" step=".1"/></label>'+
 '<label>n_core<input id="m3-core-n" type="number" value="1.4490236217" step=".0001"/></label>'+
 '<label>n_clad<input id="m3-clad-n" type="number" value="1.4440236217" step=".0001"/></label>'+
 '<label>Half-width µm<input id="m3-half" type="number" value="18" step="1"/></label>'+
 '<label>FDFD grid<select id="m3-grid"><option>29</option><option>41</option><option selected>51</option><option>61</option><option>71</option></select></label>'+
 '</div>'+
 '<div class="ls-m3-actions"><button id="m3-run">▶ 브라우저 Full-vector 계산</button>'+
 '<button id="m3-use-m2">M2 재료 계산값 가져오기</button>'+
 '<button id="m3-json" disabled>Colab 호환 결과 JSON 저장</button></div>'+
 '<p id="m3-state" role="status">브라우저 계산 준비 · 별도 Web Worker (로컬 JS) · 2개 편광 모드</p>'+
 '<div id="m3-numbers"></div><div class="ls-m3-plots">'+
 '<article><h4>|E|²</h4><canvas width="240" height="240" id="m3-total"></canvas></article>'+
 '<article><h4>|Ex|²</h4><canvas width="240" height="240" id="m3-ex"></canvas></article>'+
 '<article><h4>|Ey|²</h4><canvas width="240" height="240" id="m3-ey"></canvas></article>'+
 '<article><h4>|Ez|²</h4><canvas width="240" height="240" id="m3-ez"></canvas></article></div>'+
 '<p class="ls-m3-warning">제한: 격자 경계는 Dirichlet, 각 성분 맵은 개별 최대값 기준 색상입니다. neff는 grid 및 window 수렴 검증이 필요합니다. PML과 손실은 M4 범위입니다. M2의 κ는 아직 이 실수 벡터 FDFD에 적용되지 않습니다.</p>';
 parent.querySelector(".fv-actions")?.after(section);
 const run=$("m3-run"),state=$("m3-state"),save=$("m3-json");
 run.addEventListener("click",()=>{
  if(busy)return;
  try{
   const input={
    profile:"step",wavelength_nm:Number($("m3-lambda").value),core_radius_um:Number($("m3-radius").value),
    window_halfwidth_um:Number($("m3-half").value),grid_n:Number($("m3-grid").value),
    n_core:Number($("m3-core-n").value),n_clad:Number($("m3-clad-n").value)
   };
   if(!globalThis.Worker)throw Error("Module Web Worker unavailable in this browser.");
   worker?.terminate();worker=new Worker(new URL("./vector_fdfd_worker.mjs",import.meta.url),{type:"module"});
   const id=++sequence;busy=true;latest=null;run.disabled=true;save.disabled=true;
   state.textContent="2D Maxwell FDFD 고유값 계산 중…";
   worker.onmessage=event=>{
    const packet=event.data;if(packet?.id!==id)return;
    if(packet.type==="progress"){
     state.textContent="FDFD "+packet.detail.mode+" 편광 · "+packet.detail.iteration+"회 · 고유값 상대잔차 "+packet.detail.residual.toExponential(2);
     return;
    }
    busy=false;run.disabled=false;
    if(packet.type==="error"){state.textContent="FDFD 실패: "+packet.message;worker.terminate();return;}
    latest=packet.result;save.disabled=false;worker.terminate();
    const rows=latest.modes;
    $("m3-numbers").innerHTML='<table><thead><tr><th>Mode</th><th>neff</th><th>Eigen residual</th><th>Ex energy</th><th>Ey energy</th><th>Ez energy</th><th>MFD µm</th><th>Aeff µm²</th></tr></thead><tbody>'+
     rows.map(m=>'<tr><td>'+m.mode_index+'</td><td>'+fmt(m.neff,9)+'</td><td>'+m.residual.toExponential(2)+'</td>'+
     '<td>'+fmt(m.Ex_intensity_fraction,5)+'</td><td>'+fmt(m.Ey_intensity_fraction,5)+'</td><td>'+fmt(m.Ez_intensity_fraction,5)+'</td>'+
     '<td>'+fmt(m.MFD_2nd_um,3)+'</td><td>'+fmt(m.Aeff_um2,3)+'</td></tr>').join("")+'</tbody></table>';
    const field=rows[0].field_map.intensity_normalized_to_max_total;
    for(const [canvas,key] of [["m3-total","Itotal"],["m3-ex","Ex"],["m3-ey","Ey"],["m3-ez","Ez"]])draw($(canvas),field[key]);
    state.textContent="2개 편광 모드 계산 완료 · 스칼라 FEM과 다른 Maxwell 벡터 연산 · 실제 HE₁₁ 고정밀 검증은 별도 필요";
   };
   worker.onerror=e=>{busy=false;run.disabled=false;save.disabled=true;state.textContent="FDFD Worker 오류: "+(e.message||"unknown");worker?.terminate();};
   worker.postMessage({id,input,modes:2});
  }catch(error){busy=false;run.disabled=false;state.textContent="FDFD 입력/실행 실패: "+error.message;}
 });
 $("m3-use-m2").addEventListener("click",()=>{
  const result=globalThis.LSWaveMaterialM2?.last?.();
  if(!result?.result){state.textContent="M2에서 먼저 재료 n(λ) 계산을 실행하세요.";return;}
  const v=result.result,m=result.model;
  for(const [id,val] of [["m3-lambda",m.wavelength_nm],["m3-radius",m.core_radius_um],["m3-core-n",v.n_core],["m3-clad-n",v.n_clad]])
   $(id).value=val;
  state.textContent="M2 실수 굴절률 가져옴 · κ(흡수)는 M3 고유값 계산에 적용되지 않음.";
 });
 save.addEventListener("click",()=>{
  if(!latest)return;
  const blob=new Blob([JSON.stringify(latest,null,2)],{type:"application/json"});
  const url=URL.createObjectURL(blob),a=document.createElement("a");a.href=url;
  a.download="ls_wave_full_vector_browser_m3.json";a.click();
  setTimeout(()=>URL.revokeObjectURL(url),1000);
 });
 globalThis.LSWaveVectorM3=Object.freeze({enabled:true,latest:()=>latest});
}
if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",init,{once:true});else init();
