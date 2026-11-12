/**
 * M2 v0.6 optional, self-contained n(lambda)/k(lambda) preview.
 * No write to original scalar solver, M2 v0.4 editor, M3, or project JSON.
 * Feature flag: window.LSWaveFeatures.materialPreviewV06=false.
 */
import {M2_VERSION,M2_SOURCES,M2_CATALOG,evaluateM2Material,sampleM2Material,
 migrateM2Preview,makeM2PreviewDocument} from "./m2_preview_engine_v06.mjs?v=20261112-m3m6-v1";
import {parseIndexTable} from "./material_models.mjs";
const ENABLED=globalThis.LSWaveFeatures?.materialPreviewV06!==false;
const $=id=>document.getElementById(id);
const safe=s=>String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const fmt=(v,d=10)=>Number.isFinite(v)?Number(v).toPrecision(d):"—";
const modes=[
 ["offset","기본값: SiO₂ + Δn (기존 계산식)"],
 ["silica","순수 SiO₂"],
 ["geO2","GeO₂ 도핑 (mol% 보간)"],
 ["f_wt","F 도핑: 논문 wt% 기준 (0–2)"],
 ["f_mol","F 도핑: mol% + 사용자 Sellmeier 계수"],
 ["p2o5","P₂O₅ 도핑 (mol% 보간)"],
 ["table","사용자 n(λ) CSV / TSV"],
 ["custom_sellmeier","사용자 Sellmeier (B1–B3, L1–L3)"]
];
const modelDefault={core:{kind:"offset",deltaN:0.005,absorption:{kind:"none"}},
 clad:{kind:"silica",absorption:{kind:"none"}}};
const state=structuredClone(modelDefault);
let latest=null;
function form(role){
 const label=role==="core"?"CORE / 코어":"CLADDING / 클래딩";
 return '<fieldset class="m2v6-material" data-role="'+role+'"><legend>'+label+'</legend>'+
 '<label>재료 모델<select id="m2v6-'+role+'-kind">'+modes.map(([v,t])=>
  '<option value="'+v+'" '+(v===state[role].kind?"selected":"")+'>'+safe(t)+'</option>').join("")+'</select></label>'+
 '<div class="m2v6-fields">'+
 '<label>Δn (기본 방식)<input id="m2v6-'+role+'-offset" type="number" step=".0001" value="'+(role==="core"?.005:0)+'"/></label>'+
 '<label>도핑 농도 (mol%)<input id="m2v6-'+role+'-mol" type="number" min="0" max="100" step=".1" value="5"/></label>'+
 '<label>F 실험농도 (wt%)<input id="m2v6-'+role+'-wt" type="number" min="0" max="2" step=".1" value="1"/></label>'+
 '<label>보간 방식<select id="m2v6-'+role+'-interp"><option value="linear">Linear</option><option value="pchip">PCHIP (shape-preserving)</option></select></label>'+
 '<label>흡수 입력 방식<select id="m2v6-'+role+'-abs"><option value="none">기본 / 테이블 k</option><option value="k">κ 직접 입력</option><option value="alpha">α (dB/km) 입력</option></select></label>'+
 '<label>κ 또는 α 입력값<input id="m2v6-'+role+'-absvalue" type="number" value="0" min="0" step="any"/></label>'+
 '</div><details id="m2v6-'+role+'-csvbox"><summary>사용자 n(λ) 테이블 (CSV 붙여넣기)</summary>'+
 '<p>열 순서: wavelength_nm,n_real[,kappa]. Linear/PCHIP 선택 가능. 범위 밖 외삽 금지.</p>'+
 '<textarea id="m2v6-'+role+'-csv" rows="4" spellcheck="false" placeholder="1300,1.450,0&#10;1550,1.445,0&#10;1700,1.443,0"></textarea></details>'+
 '<details id="m2v6-'+role+'-coeffbox"><summary>사용자 Sellmeier 계수 및 출처 (F mol%는 필수)</summary>'+
 '<p>3-term B1,B2,B3 / L1,L2,L3 [µm] (실측 또는 확인 가능한 출처를 직접 입력)</p>'+
 '<label>B1, B2, B3<input id="m2v6-'+role+'-B" placeholder="B1,B2,B3"/></label>'+
 '<label>L1, L2, L3 [µm]<input id="m2v6-'+role+'-L" placeholder="L1,L2,L3"/></label>'+
 '<label>계수 출처 / 실험 조건<input id="m2v6-'+role+'-source" placeholder="논문 DOI/표준/측정 보고서"/></label></details>'+
 '<p class="m2v6-hint" id="m2v6-'+role+'-hint">기본 Δn 모델 · 파장 독립 추가값</p></fieldset>';
}
function readCoefficients(role){
 const B=$("m2v6-"+role+"-B").value.split(/[,\s]+/).filter(Boolean).map(Number);
 const L=$("m2v6-"+role+"-L").value.split(/[,\s]+/).filter(Boolean).map(Number);
 return {B,L,source:$("m2v6-"+role+"-source").value.trim()};
}
function readMaterial(role){
 const root="m2v6-"+role+"-",kind=$(root+"kind").value;
 const m={kind,absorption:{kind:$(root+"abs").value}};
 if(m.absorption.kind!=="none"){
  const raw=$(root+"absvalue").value.trim();
  if(raw==="")throw Error(role+" κ/α value required.");
  m.absorption.value=Number(raw);
 }
 if(kind==="offset")m.deltaN=Number($(root+"offset").value);
 else if(kind==="geO2"||kind==="p2o5"||kind==="f_mol")m.molPercent=Number($(root+"mol").value);
 else if(kind==="f_wt")m.weightPercent=Number($(root+"wt").value);
 if(kind==="f_mol"||kind==="custom_sellmeier")m.coefficients=readCoefficients(role);
 if(kind==="table"){
  m.table=parseIndexTable($(root+"csv").value);
  m.interpolation=$(root+"interp").value;
 }
 return m;
}
function writeMaterial(role,m){
 if(m.kind==="f_legacy_unresolved"){
  $("m2v6-"+role+"-hint").textContent=m.warning;
  $("m2v6-"+role+"-kind").value="f_mol";
  return;
 }
 if(!modes.some(([k])=>k===m.kind))throw Error("Unknown material kind "+m.kind);
 const root="m2v6-"+role+"-";
 $(root+"kind").value=m.kind;
 if(m.deltaN!=null)$(root+"offset").value=m.deltaN;
 if(m.molPercent!=null)$(root+"mol").value=m.molPercent;
 if(m.weightPercent!=null)$(root+"wt").value=m.weightPercent;
 if(m.absorption){
  $(root+"abs").value=m.absorption.kind||"none";
  $(root+"absvalue").value=m.absorption.value??0;
 }
 if(m.table)$(root+"csv").value=m.table.map(x=>[x.wavelength_nm,x.n,x.kappa??0].join(",")).join("\n");
 if(m.interpolation)$(root+"interp").value=m.interpolation;
 if(m.coefficients){
  $(root+"B").value=(m.coefficients.B||[]).join(",");
  $(root+"L").value=(m.coefficients.L||[]).join(",");
  $(root+"source").value=m.coefficients.source||"";
 }
}
function provenance(role){
 const kind=$("m2v6-"+role+"-kind").value;
 const sourceKey=kind==="f_wt"?"fwt":kind==="f_mol"?"fmol":kind==="table"?"table":kind==="offset"?"silica":kind;
 const src=M2_SOURCES[sourceKey]||M2_SOURCES.table;
 const hints={
  offset:"SiO₂ Sellmeier + 파장 무관 Δn · 기존 기본 모델(변경 없음)",
  silica:"순수 SiO₂ · 기존 Sellmeier 함수와 완전 동일",
  geO2:"0–30 mol% · SiO₂/GeO₂ endpoint 계수의 선형 혼합(중간 농도 실측 검증 아님)",
  f_wt:"원문 1/2 wt% 실측 피팅값 · mol%와 혼용 금지",
  f_mol:"F mol% 사용자 지정 · 공식 mol%-계수 매핑은 확인되지 않아 직접 계수+출처 필요",
  p2o5:"0–9.1 mol% · 9.1 mol% 실측 기준 계수 보간(중간 농도 실측 검증 아님)",
  table:"CSV/PCHIP/선형 · 범위 밖 외삽 거부 · 사용자 데이터 책임",
  custom_sellmeier:"사용자 실측 Sellmeier 6계수와 출처 필요"
 };
 $("m2v6-"+role+"-hint").textContent=hints[kind]||"";
 const info=$("m2v6-"+role+"-sourceinfo");
 if(info)info.textContent=src.label;
}
function draw(canvas,series,field,title){
 const ctx=canvas.getContext("2d");if(!ctx)return;
 const {width:W,height:H}=canvas;ctx.clearRect(0,0,W,H);
 ctx.fillStyle="#0b1f31";ctx.fillRect(0,0,W,H);
 const margins={l:68,r:14,t:26,b:40},pw=W-margins.l-margins.r,ph=H-margins.t-margins.b;
 const all=series.flatMap(s=>s.rows.map(v=>v[field])).filter(Number.isFinite);
 if(!all.length){ctx.fillStyle="#c2d5e0";ctx.fillText("No valid samples",35,H/2);return;}
 let lo=Math.min(...all),hi=Math.max(...all),span=hi-lo;
 if(span<1e-15){const d=Math.max(Math.abs(hi)*.03,field==="kappa"?1e-12:.001);lo-=d;hi+=d}
 else{lo-=.08*span;hi+=.08*span}
 const xvalues=series.flatMap(s=>s.rows.map(v=>v.lambdaNm)),xmin=Math.min(...xvalues),xmax=Math.max(...xvalues);
 const xx=x=>margins.l+(x-xmin)/(xmax-xmin)*pw,yy=y=>margins.t+(hi-y)/(hi-lo)*ph;
 ctx.strokeStyle="#294659";ctx.lineWidth=1;ctx.font="11px Arial";ctx.fillStyle="#a9c5d7";
 for(let i=0;i<=4;i++){
  const y=margins.t+i*ph/4;
  ctx.beginPath();ctx.moveTo(margins.l,y);ctx.lineTo(W-margins.r,y);ctx.stroke();
  const v=hi-(hi-lo)*i/4;
  ctx.fillText(field==="kappa"?v.toExponential(1):v.toFixed(5),6,y+3);
 }
 for(let i=0;i<=4;i++){
  const x=margins.l+i*pw/4;
  ctx.beginPath();ctx.moveTo(x,margins.t);ctx.lineTo(x,H-margins.b);ctx.stroke();
  ctx.fillText(Math.round(xmin+i*(xmax-xmin)/4),x-12,H-20);
 }
 ctx.fillText("λ (nm)",W-58,H-5);ctx.fillText(title,margins.l,16);
 for(const s of series){
  ctx.strokeStyle=s.color;ctx.lineWidth=2;ctx.beginPath();
  s.rows.forEach((v,i)=>{if(i===0)ctx.moveTo(xx(v.lambdaNm),yy(v[field]));else ctx.lineTo(xx(v.lambdaNm),yy(v[field]));});
  ctx.stroke();
 }
}
function output(){
 const status=$("m2v6-status"),errors=[],start=Number($("m2v6-start").value),end=Number($("m2v6-end").value);
 const active=[];
 for(const role of ["core","clad"]){
  const model=readMaterial(role);
  state[role]=structuredClone(model);
  provenance(role);
  try{
   const rows=sampleM2Material(model,{startNm:start,endNm:end,count:71});
   const at1550=evaluateM2Material(model,1550);
   active.push({name:role==="core"?"CORE":"CLADDING",color:role==="core"?"#73d7f8":"#ffa77a",rows,at1550,role,model});
  }catch(e){errors.push((role==="core"?"Core":"Cladding")+": "+e.message)}
 }
 const refs=[];
 for(const [i,preset] of M2_CATALOG.entries()){
  try{
   const rows=sampleM2Material(preset.model,{startNm:start,endNm:end,count:71});
   refs.push({name:preset.label,color:["#9fc5d2","#93d5ad","#ba9dd2","#f1c68a"][i],rows,
    at1550:evaluateM2Material(preset.model,1550)});
  }catch(e){errors.push(preset.label+": "+e.message)}
 }
 const series=[...refs,...active],body=$("m2v6-rows");
 body.innerHTML=series.map(s=>'<tr><td>'+safe(s.name)+'</td>'+
  '<td>'+fmt(s.at1550.n,11)+'</td><td>'+s.at1550.kappa.toExponential(4)+'</td>'+
  '<td>'+fmt(s.at1550.alphaDbKm,7)+'</td><td>'+safe(s.role?"선택 재료":"참고곡선")+'</td></tr>').join("");
 draw($("m2v6-nchart"),series,"n","Real n(λ) · Sellmeier / table");
 draw($("m2v6-kchart"),series,"kappa","Imaginary κ(λ) · alpha/k preview");
 $("m2v6-legend").innerHTML=series.map(s=>'<span><i style="background:'+s.color+'"></i>'+safe(s.name)+'</span>').join("");
 latest=makeM2PreviewDocument(state.core,state.clad);
 latest.preview={rangeNm:[start,end],rows:series.map(s=>({name:s.name,at1550:s.at1550}))};
 status.textContent=errors.length?"미리보기 일부 실패 · "+errors.join(" / "):
  "미리보기 성공 · 1550 nm 표 및 n(λ)/κ(λ) 그래프. M3/Scalar FEM에는 적용하지 않음.";
 $("m2v6-save").disabled=false;
}
function download(){
 if(!latest)return;
 const blob=new Blob([JSON.stringify(latest,null,2)],{type:"application/json"}),url=URL.createObjectURL(blob);
 const a=document.createElement("a");a.href=url;a.download="LS_Wave_M2_v0.6_material_preview.json";
 document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);
}
function mount(){
 if(!ENABLED||$("m2v6-preview"))return;
 const nav=document.querySelector("nav.eng-shortcuts");if(!nav)return;
 const tab=document.createElement("a");tab.id="m2v6-nav";tab.href="#m2v6-preview";
 tab.textContent="M2 · MATERIAL PREVIEW v0.6";nav.append(tab);
 const section=document.createElement("section");
 section.id="m2v6-preview";section.hidden=true;section.className="m2v6-preview engineering-panel";
 section.innerHTML='<header><small>M2 · MATERIAL MODEL PREVIEW · v'+M2_VERSION+
 '</small><h2>파장 의존 n(λ) · 도핑 Sellmeier · 복소 굴절률</h2>'+
 '<p>기존 SiO₂+Δn 계산은 기본값으로 유지. 아래 모델은 미리보기/JSON 전용이며 벡터/스칼라 고유값 솔버에 자동 연결되지 않습니다.</p></header>'+
 '<div class="m2v6-range"><label>파장 시작 (nm)<input id="m2v6-start" type="number" value="1300" min="400" max="1700"/></label>'+
 '<label>파장 종료 (nm)<input id="m2v6-end" type="number" value="1650" min="400" max="1700"/></label></div>'+
 '<div class="m2v6-materials">'+form("core")+form("clad")+'</div>'+
 '<div class="m2v6-actions"><button type="button" id="m2v6-run">▶ 재료 곡선 미리보기</button>'+
 '<button type="button" id="m2v6-save" disabled>재료 JSON 저장</button>'+
 '<label class="m2v6-import">재료 JSON 불러오기<input type="file" id="m2v6-open" accept=".json" hidden/></label></div>'+
 '<p id="m2v6-status" role="status">미리보기 실행 대기 · F wt%/mol% 단위 구분 필수</p>'+
 '<div id="m2v6-legend" class="m2v6-legend"></div>'+
 '<div class="m2v6-plots"><canvas id="m2v6-nchart" width="590" height="300" aria-label="n wavelength chart"></canvas>'+
 '<canvas id="m2v6-kchart" width="590" height="300" aria-label="k wavelength chart"></canvas></div>'+
 '<h3>1550 nm 재료 비교</h3><div class="m2v6-scroll"><table><thead><tr>'+
 '<th>재료 / 구성</th><th>n_real @1550</th><th>κ @1550</th><th>α (dB/km)</th><th>구분</th>'+
 '</tr></thead><tbody id="m2v6-rows"><tr><td colspan="5">계산 전</td></tr></tbody></table></div>'+
 '<details class="m2v6-provenance" open><summary>근거·단위·기능 한계</summary>'+
 '<p>순수 SiO₂: <a href="'+M2_SOURCES.silica.url+'" target="_blank" rel="noopener noreferrer">Malitson, JOSA 55 (1965), Eq. 1 (20°C)</a>. 기존 스칼라 굴절률 계산 계수와 동일합니다.</p>'+
 '<p>GeO₂: <a href="'+M2_SOURCES.geO2.url+'" target="_blank" rel="noopener noreferrer">Fleming, Applied Optics 23 (1984), Table II</a>; 중간 농도는 단순 계수 보간으로 실측값이 아닙니다.</p>'+
 '<p>F: <a href="'+M2_SOURCES.fwt.url+'" target="_blank">Fleming & Wood, Applied Optics 22 (1983), Table II</a>의 1/2 <b>wt%</b> 실험. 몰농도 대응 계수는 확인되지 않아 사용자 mol% 선택 시 B/L 6개와 출처를 필수 입력합니다.</p>'+
 '<p>P₂O₅: <a href="'+M2_SOURCES.p2o5.url+'" target="_blank">광섬유 교재 Table 3.4</a>의 9.1 mol% 참고 피팅. 중간 농도는 보간 추정이며 실측 검증 아님.</p>'+
 '<p>복소 굴절률은 ñ=n+iκ, exp(iβz−iωt) 사용. α(dB/km)=4πκ×(1000/λ[m])×10/ln(10). 흡수/전파 손실은 <b>미리보기만</b> 수행하며 고유모드 감쇠 계산 아님.</p>'+
 '<p>CSV 입력은 n(λ) 두 점 이상, Linear/PCHIP 선택. 파장 범위 밖 외삽 금지(경고). 구 M2 schemaVersion 1의 F 입력 단위 불명은 수동 확인을 요청합니다.</p></details>';
 nav.after(section);
 const show=()=>{section.hidden=false;section.scrollIntoView({block:"start",behavior:"smooth"});};
 tab.addEventListener("click",e=>{e.preventDefault();show()});
 $("m2v6-run").addEventListener("click",()=>{try{output()}catch(e){$("m2v6-status").textContent="미리보기 실패: "+e.message;}});
 $("m2v6-save").addEventListener("click",download);
 for(const role of ["core","clad"]) $("m2v6-"+role+"-kind").addEventListener("change",()=>provenance(role));
 $("m2v6-open").addEventListener("change",async e=>{
  const f=e.target.files?.[0];if(!f)return;
  if(f.size>250000){$("m2v6-status").textContent="재료 파일 제한 250KB";return}
  try{
   const doc=migrateM2Preview(JSON.parse(await f.text()));
   for(const role of ["core","clad"])writeMaterial(role,doc[role]);
   $("m2v6-status").textContent="재료 JSON 불러옴 · 계산은 다시 실행. F 단위 모호한 옛 파일은 수동 확인.";
  }catch(error){$("m2v6-status").textContent="재료 불러오기 실패: "+error.message;}
  e.target.value="";
 });
 globalThis.LSWaveM2Preview=Object.freeze({enabled:true,run:output,latest:()=>latest,show,snapshot:()=>Object.fromEntries(["core","clad"].map(r=>[r,readMaterial(r)])),loadProject:doc=>{for(const r of ["core","clad"])writeMaterial(r,doc[r]);}});
 for(const role of ["core","clad"])provenance(role);
}
if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",mount,{once:true});
else mount();
