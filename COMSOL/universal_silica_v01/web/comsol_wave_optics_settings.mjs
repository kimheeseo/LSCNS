(function waveOpticsSettings(){
 "use strict";
 const $=id=>document.getElementById(id),sidebar=document.querySelector(".layout > .sidebar");
 if(!sidebar||$("ls-wave-settings"))return;
 const pane=document.createElement("section");
 pane.id="ls-wave-settings";pane.className="ls-config-pane ls-wave-pane";pane.hidden=true;
 sidebar.append(pane);
 let active="wavePhysics",showNote="",error=false;
 const names={
  waveMaterials:["Materials","Material Assignment"],
  waveCore:["Core Material","Refractive Index · n₁"],
  waveCladding:["Cladding Material","Refractive Index · n₂"],
  wavePhysics:["Electromagnetic Waves, Frequency Domain (ewfd)","Physics Interface"],
  waveEquation:["Wave Equation, Electric 1","Wave Equation, Electric"],
  wavePEC:["Perfect Electric Conductor 1","Boundary Condition"],
  waveInitial:["Initial Values 1","Initial Values"],
  waveMesh:["Mesh 1","Mesh Sequence"],
  waveSize:["Size 1","Element Size"],
  waveStudy:["Study 1","Study Configuration"],
  waveMode:["Step 1: Mode Analysis","Mode Analysis"],
  waveResults:["Results","Results Browser"],
  waveField:["Mode Intensity · I(r)","Field Plot"],
  waveIndex:["Refractive Index · n(r)","Profile Plot"],
  waveDerived:["Derived Values","Global Evaluation"]
 };
 const nav=(id,label)=>`<button type="button" class="ls-wave-link" data-wave-nav="${id}">${label}</button>`;
 const row=(label,value)=>`<div class="ls-wave-reading"><span>${label}</span><strong>${value}</strong></div>`;
 const section=(name,body)=>`<details class="ls-wave-section" open><summary>${name}</summary><div class="ls-wave-section-content">${body}</div></details>`;
 const read=id=>$(id)?.value??"";
 const val=id=>$(id)?.textContent?.trim()||"—";
 const select=(name,id,options)=>`<label class="ls-wave-field"><span>${name}</span><select id="${id}">${options}</select></label>`;
 const input=(label,id,value,unit="")=>`<label class="ls-wave-field"><span>${label}</span><span class="ls-wave-entry"><input type="number" id="${id}" value="${value}" step="any" />${unit?`<small>${unit}</small>`:""}</span></label>`;
 const primary=(action,label)=>`<button type="button" class="ls-wave-primary" data-wave-command="${action}">${label}</button>`;
 const secondary=(action,label)=>`<button type="button" class="ls-wave-secondary" data-wave-command="${action}">${label}</button>`;
 const note=message=>`<p class="ls-wave-context">${message}</p>`;
 function results(){
  const vals=[["n_eff",val("neff")],["MFD · µm",val("mfd")],["A_eff · µm²",val("aeff")],["D · ps/(nm·km)",val("disp")]];
  return `<div class="ls-wave-results">${vals.map(x=>row(x[0],x[1])).join("")}</div>`;
 }
 function content(panel){
  const isScalar=(read("preset")!=="hcf"&&read("preset")!=="vector"&&read("profile")!=="mcf");
  const modeP=read("clad_material");
  const rawN=read("clad_index_override");
  const wavelength=read("wavelength_nm"),delta=read("delta_n_core"),elements=read("elements");
  const domain=note("Core/Cladding의 재료 구역은 현재 동심원 SMF용 반경 모델입니다. 2D 임의 도형에 대한 실제 물성 Domain Selection은 아직 계산기에 연결되지 않았습니다.");
  const noFullVector=note("현재 브라우저 기본 엔진은 1D radial scalar FEM (LP₀₁ 근사)입니다. 아래 ewfd 메뉴는 COMSOL Wave Optics의 화면 구조를 참고한 것입니다. Full-vector Maxwell FEM 또는 PEC/PML 경계의 해석 구현을 의미하지 않습니다.");
  const modeButton=primary("compute","▶ Compute · 기존 Radial FEM");
  switch(panel){
   case "waveMaterials":
    return section("Materials",`${row("Material 1","Silica core · n₁ = n₂ + Δn")}${row("Material 2",modeP==="fixed"?"Fixed cladding n₂ = "+rawN:"Silica Sellmeier n₂(λ)")}${nav("waveCore","Core material settings")}${nav("waveCladding","Cladding material settings")}`)+section("Geometric Entity Selection",domain)+note("실제 계산 입력은 기존 광섬유 설계 패널과 공유합니다.");
   case "waveCore":
    return section("Material Contents",`${row("Label","Core (n₁)")}${row("Refractive index definition","n_core = n_clad + Δn")}${input("Core − cladding index contrast","ls-wave-delta",delta,"Δn")}`)+section("Domain Selection",domain)+primary("applyDelta","Apply Δn to current solver");
   case "waveCladding":
    return section("Material Contents",`${select("Cladding index model","ls-wave-clad-mode",'<option value="silica">Silica Sellmeier</option><option value="fixed">Constant refractive index</option>')}${input("Constant n_clad","ls-wave-clad-index",rawN,"dimensionless")}`)+section("Material Assignment",domain)+primary("applyClad","Apply cladding model to current solver");
   case "wavePhysics":
    return section("Physics Interface",`${row("Label","Electromagnetic Waves, Frequency Domain")}${row("COMSOL interface tag","ewfd")}${row("Linked study","Mode Analysis · LP₀₁ scalar")}${row("Current numerical engine",isScalar?"Radial scalar FEM":"Specialized research mode")}`)+section("Physics Features",`${nav("waveEquation","Wave Equation, Electric 1")}${nav("wavePEC","Perfect Electric Conductor 1")}${nav("waveInitial","Initial Values 1")}`)+noFullVector+nav("waveMode","Configure Mode Analysis");
   case "waveEquation":
    return section("Domain Selection",row("Selection","All radial domains (core + cladding)"))+
     section("Electric Displacement Field",`${row("Refractive index","n₂(λ) + Δn (core only)")}${row("Relative permeability","μᵣ = 1 (assumed)")}${row("Polarization","Scalar LP₀₁ approximation")}`)+section("Input Parameters",`${input("Wavelength","ls-wave-wavelength",wavelength,"nm")}${input("Core index contrast","ls-wave-delta",delta,"Δn")}`)+primary("applyWaveEquation","Apply wavelength and Δn")+noFullVector;
   case "wavePEC":
    return section("Boundary Selection",`${row("COMSOL default node","Perfect Electric Conductor 1")}${row("Browser FEM boundary","Finite outer-domain numerical boundary")}`)+note("PEC는 COMSOL ewfd에서 기본 제공하는 경계조건 노드입니다. 현재의 Radial FEM이 동일한 벡터 PEC 경계식을 지원하는 것은 아니며, 이 화면에서는 변경·적용할 수 없습니다.");
   case "waveInitial":
    return section("Initial Values",`${row("Electric field components","Ex / Ey / Ez (COMSOL reference)")}${row("Radial solver","Eigenmode is found numerically")}`)+note("현재 Scalar Radial Mode Analysis에서는 전기장의 개별 Ex/Ey/Ez 초기값 편집을 지원하지 않습니다.");
   case "waveMesh":
    return section("Mesh Settings",`${row("Sequence type","User-controlled radial FEM")}${row("Element type","1D P1 (radial elements)")}${row("Radial elements",elements)}`)+section("Mesh Sequence",nav("waveSize","Size 1 · element count"))+
     `${secondary("mesh","Mesh convergence · 실제 재계산")}`+
     note("실제 COMSOL ewfd의 2D 삼각형/사각형 메시나 PML 구조 격자는 아직 생성하지 않습니다.");
   case "waveSize":{
    const outer=Number(read("outer_radius_um"));const e=Number(elements);
    return section("Element Size",`${input("Radial elements","ls-wave-elements",elements,"elements")}${row("Computational radius",`${read("outer_radius_um")} µm`)}${row("Nominal radial spacing",outer>0&&e>0?(outer/e).toPrecision(4)+" µm":"—")}`)+
     primary("applyMesh","Apply element count")+secondary("mesh","Mesh convergence study")+note("범위: 100–1800, step 5. 계산 엔진이 사용하는 요소 수에 직접 연결합니다.");
   }
   case "waveStudy":
    return section("Study Type",`${row("Study 1","Mode Analysis")}${row("Supported","LP₀₁ scalar mode + wavelength sweep")}`)+nav("waveMode","Step 1: Mode Analysis")+modeButton+noFullVector;
   case "waveMode":
    return section("Study Settings",`${input("Free-space wavelength","ls-wave-wavelength",wavelength,"nm")}${row("Mode count","1 (fundamental LP₀₁ only)")}${row("Eigenvalue output","n_eff")}`)+
     section("Solver",`${row("Discretization","Radial P1 FEM")}${row("Core radius",read("core_radius_um")+" µm")}${row("Radial mesh",elements+" elements")}`)+
     primary("applyCompute","Apply wavelength & Compute")+secondary("sweep","Wavelength sweep")+noFullVector;
   case "waveResults":
    return section("Datasets",row("Dataset 1","Last computed scalar LP₀₁ mode"))+
     section("Derived Values",results())+section("Plots",`${nav("waveField","Mode Intensity · I(r)")}${nav("waveIndex","Refractive Index · n(r)")}${nav("waveDerived","n_eff / MFD / A_eff / D")}`)+secondary("export","Export solver JSON");
   case "waveField":
    return section("Mode Intensity",`${row("Display","LP₀₁ intensity I(r), normalized")}${row("Source","Existing radial solver field output")}`)+
      primary("plotField","Show mode intensity plot")+note("기존 계산 결과를 표시합니다. Ex/Ey/Ez full-vector field map으로 표시하지 않습니다.");
   case "waveIndex":
    return section("Refractive Index",`${row("Display","n(r)")}${row("Index model",modeP==="fixed"?"Constant n₂ + Δn":"Sellmeier silica + Δn")}`)+
      primary("plotIndex","Show index profile plot");
   case "waveDerived":
    return section("Global Evaluation",results())+secondary("export","Export solver JSON")+note("Derived Values는 마지막 실제 FEM 결과를 읽습니다. 데이터 변경 시 재계산이 필요합니다.");
   default:return section("Wave Optics",nav("wavePhysics","Electromagnetic Waves, Frequency Domain"));
  }
 }
 function updateSelected(){
  document.querySelectorAll(".ls-model-tree [data-action^='wave']").forEach(n=>{
   const state=n.dataset.action===active;
   n.classList.toggle("ls-wave-node-selected",state);
   if(state)n.setAttribute("aria-current","true");else n.removeAttribute("aria-current");
  });
 }
 function render(){
  const [title,sub]=names[active]||names.wavePhysics;
  pane.innerHTML=`<div class="ls-config-heading"><div><strong>Settings</strong><small>Wave Optics / ${sub}</small></div><button type="button" id="ls-wave-back">← Fiber settings</button></div>
   <div class="ls-wave-heading"><span class="ls-wave-heading-icon">◈</span><div><h3>${title}</h3><small>Component 1 · LS Wave Optics</small></div></div>
   <div class="ls-wave-body">${content(active)}</div>
   <p id="ls-wave-status" class="ls-config-status ${error?"error":""}" role="status" aria-live="polite">${showNote||"Select a Model Builder node to edit its settings."}</p>`;
  if($("ls-wave-clad-mode"))$("ls-wave-clad-mode").value=read("clad_material");
  if($("ls-wave-clad-index"))$("ls-wave-clad-index").disabled=read("clad_material")!=="fixed";
  updateSelected();
 }
 function close(){
  if(pane.hidden)return;
  pane.hidden=true;
  document.body.classList.remove("ls-wave-optics-active");
  const hasOther=document.body.classList.contains("ls-cad-active")||($("ls-config-pane")&&!$("ls-config-pane").hidden);
  if(!hasOther){sidebar.classList.remove("ls-config-active");document.body.classList.remove("ls-configuration-open");}
 }
 function open(id){
  if(!(id in names))id="wavePhysics";
  const cfg=$("ls-config-pane"),cad=$("ls-cad-pane");
  if(cfg)cfg.hidden=true;
  if(cad)cad.hidden=true;
  active=id;
  pane.hidden=false;
  sidebar.classList.add("ls-config-active");
  document.body.classList.add("ls-configuration-open","ls-wave-optics-active");
  showNote="COMSOL Wave Optics UI reference · linked scalar solver where explicitly stated.";
  error=false;render();sidebar.scrollTo({top:0,behavior:"instant"});
 }
 function setAuthoritative(id,value){
  const target=$(id);
  if(!target)throw Error("Missing solver input: "+id);
  target.value=String(value);
  target.dispatchEvent(new Event("input",{bubbles:true}));
 }
 function action(cmd){
  try{
   if(cmd==="compute"){$("calculateBtn")?.click();showNote="Mode Analysis invoked with current numerical inputs.";error=false;render();return;}
   if(cmd==="mesh"){$("meshBtn")?.click();showNote="Existing mesh convergence analysis invoked.";error=false;render();return;}
   if(cmd==="sweep"){$("sweepBtn")?.click();showNote="Existing wavelength sweep invoked.";error=false;render();return;}
   if(cmd==="export"){$("jsonBtn")?.click();showNote="Existing solver JSON export invoked.";error=false;render();return;}
   if(cmd==="plotField"||cmd==="plotIndex"){
    const sel=cmd==="plotField"?"#modeChart":"#indexChart",target=document.querySelector(sel);
    if(!target)throw Error("Plot unavailable");
    target.scrollIntoView({block:"center",behavior:"smooth"});showNote="Showing existing computed plot in Graphics.";error=false;render();return;
   }
   if(["applyDelta","applyWaveEquation","applyCompute"].includes(cmd)){
    const n=Number($("ls-wave-delta")?.value??read("delta_n_core"));
    if(!(n>0&&n<0.1))throw Error("Δn must be between 0 and 0.1");
    const wl=Number($("ls-wave-wavelength")?.value??read("wavelength_nm"));
    if(!(wl>=250&&wl<=3600))throw Error("Wavelength must be 250–3600 nm");
    if(cmd!=="applyCompute")setAuthoritative("delta_n_core",n);
    if(cmd!=="applyDelta")setAuthoritative("wavelength_nm",wl);
    if(cmd==="applyCompute")$("calculateBtn")?.click();
    showNote=cmd==="applyCompute"?"Wavelength applied; Mode Analysis computation invoked.":"Existing FEM input field(s) updated. Run Compute to solve.";
   }else if(cmd==="applyClad"){
    const model=$("ls-wave-clad-mode").value,n=Number($("ls-wave-clad-index").value);
    if(model==="fixed"&&!(n>=1&&n<2))throw Error("Fixed cladding index must be between 1 and 2");
    setAuthoritative("clad_material",model);
    if(model==="fixed")setAuthoritative("clad_index_override",n);
    showNote="Material index model applied. Run Compute to solve.";
   }else if(cmd==="applyMesh"){
    const n=Number($("ls-wave-elements").value);
    if(!Number.isInteger(n)||n<100||n>1800||n%5!==0)throw Error("Radial elements must be an integer multiple of 5 from 100 to 1800.");
    setAuthoritative("elements",n);showNote="Radial mesh element count applied. Run Compute to solve.";
   }
   error=false;render();
  }catch(e){showNote=e.message;error=true;render();}
 }
 pane.addEventListener("click",e=>{
  const back=e.target.closest("#ls-wave-back");if(back){close();return;}
  const nav=e.target.closest("[data-wave-nav]");if(nav){active=nav.dataset.waveNav;showNote="";error=false;render();sidebar.scrollTop=0;return;}
  const command=e.target.closest("[data-wave-command]");if(command){action(command.dataset.waveCommand);return;}
 });
 pane.addEventListener("change",e=>{
  if(e.target.id==="ls-wave-clad-mode"){
   const fixed=e.target.value==="fixed";
   if($("ls-wave-clad-index"))$("ls-wave-clad-index").disabled=!fixed;
  }
 });
 document.addEventListener("ls-open-wave-optics",e=>open(e.detail?.panel));
 // Other panels have already opened by the time these handlers run.
 document.addEventListener("ls-open-cad",close);
 document.addEventListener("ls-cad-add",close);
 document.addEventListener("ls-open-settings",close);
 const observer=new MutationObserver(()=>{if(!pane.hidden&&["waveResults","waveDerived"].includes(active))render();});
 if($("neff"))observer.observe($("neff"),{childList:true,characterData:true,subtree:true});
 if($("notice"))observer.observe($("notice"),{childList:true,characterData:true,subtree:true});
 $("preset")?.addEventListener("change",close);
 $("resetBtn")?.addEventListener("click",close);
 // Initial appearance mirrors COMSOL Desktop: a selected physics node opens its Settings.
 open("wavePhysics");
})();
