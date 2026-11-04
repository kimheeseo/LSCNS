(function stageThreeWorkflow(){
 "use strict";
 const $=id=>document.getElementById(id);
 const pane=document.querySelector(".ls-study-pane"),list=pane?.querySelector(".ls-study-list");
 if(!pane||!list||$("ls-linked-workflow"))return;
 const root=document.createElement("section");
 root.id="ls-linked-workflow";root.className="ls-workflow";root.setAttribute("aria-label","Linked Geometry Materials Mesh Study Results workflow");
 root.innerHTML=`
  <div class="ls-workflow-heading"><strong>Model Workflow</strong><small>Connected · radial FEM</small></div>
  <div class="ls-workflow-steps">
   <button data-wf-step="geometry" type="button"><span>1</span> Geometry <b id="ls-wf-geometry-state">Ready</b></button>
   <button data-wf-step="materials" type="button"><span>2</span> Materials <b id="ls-wf-material-state">Ready</b></button>
   <button data-wf-step="mesh" type="button"><span>3</span> Mesh <b id="ls-wf-mesh-state">Ready</b></button>
   <button data-wf-step="study" type="button"><span>4</span> Study <b id="ls-wf-study-state">Ready</b></button>
   <button data-wf-step="results" type="button"><span>5</span> Results <b id="ls-wf-result-state">Pending</b></button>
  </div>
  <div class="ls-wf-domain-display" id="ls-wf-domains">Domain 1 · Core / Domain 2 · Cladding</div>
  <button type="button" id="ls-wf-apply-geometry">Apply Geometry to FEM</button>
  <button type="button" id="ls-wf-compute">▶ Compute linked mode</button>
  <p id="ls-wf-status" role="status" aria-live="polite"></p>`;
 list.before(root);
 const linkedIds=["preset","profile","wavelength_nm","core_radius_um","outer_radius_um","delta_n_core","elements","clad_material","clad_index_override"];
 let resultSignature=null,pendingSignature=null,runCount=0,selectedDomain="core",feedback="",feedbackError=false;
 const signature=()=>JSON.stringify(linkedIds.map(id=>$(id)?.value??null));
 const geometry=()=>window.LSWaveCADService?.status()||{supported:false,reason:"Geometry editor not ready",pending:false,domains:[]};
 function scalarMode(){const preset=$("preset")?.value,profile=$("profile")?.value;return !["hcf","vector","mcf"].includes(preset)&&profile!=="mcf";}
 function numericReady(){
  const n=Number($("delta_n_core")?.value),wl=Number($("wavelength_nm")?.value),
    core=Number($("core_radius_um")?.value),outer=Number($("outer_radius_um")?.value),mesh=Number($("elements")?.value);
  if(!Number.isFinite(n)||n<=0||n>=0.1)return "Core Δn must be between 0 and 0.1.";
  if(!(wl>=250&&wl<=3600))return "Wavelength must be 250–3600 nm.";
  if(!(core>=.5&&core<=60&&outer>=15&&outer<=150&&outer>core))return "Core/cladding geometry outside solver limits.";
  if(!Number.isInteger(mesh)||mesh<100||mesh>1800)return "Radial mesh must contain 100–1800 elements.";
  if($("clad_material")?.value==="fixed"&&!(Number($("clad_index_override")?.value)>=1&&Number($("clad_index_override")?.value)<2))return "Fixed cladding index must be 1–2.";
  return "";
 }
 function blocker(){
  const g=geometry();
  if(!scalarMode())return "Linked workflow supports scalar SMF/step-index profiles. HCF, MCF and vector research are separate.";
  if(!g.supported)return "Unsupported geometry: "+g.reason;
  if(g.pending)return "Geometry edits not applied. Use Apply Geometry to FEM first.";
  return numericReady();
 }
 function setState(id,text,state){
  const el=$(id);if(!el)return;el.textContent=text;el.dataset.state=state;
 }
 function update(){
  const g=geometry(),problem=numericReady(),scalar=scalarMode();
  setState("ls-wf-geometry-state",!g.supported?"Blocked":g.pending?"Apply":"Ready",!g.supported?"blocked":g.pending?"pending":"ready");
  setState("ls-wf-material-state",!scalar?"N/A":problem&&/Δn|cladding/i.test(problem)?"Check":"Ready",scalar?"ready":"pending");
  setState("ls-wf-mesh-state",!scalar?"N/A":problem&&/mesh/i.test(problem)?"Check":"Ready",scalar?"ready":"pending");
  setState("ls-wf-study-state",!scalar?"N/A":blocker()?"Blocked":"Ready",blocker()?"blocked":"ready");
  const fresh=!!resultSignature&&resultSignature===signature()&&scalar&&g.supported&&!g.pending;
  setState("ls-wf-result-state",fresh?"Current":resultSignature?"Stale":"Pending",fresh?"ready":"pending");
  $("ls-wf-domains").textContent=g.supported?`D1 Core (${g.coreRadiusUm.toFixed(3)} µm) · D2 Cladding (R ${g.outerRadiusUm.toFixed(3)} µm)`:"2D preview only · no FEM domain assignment";
  $("ls-wf-apply-geometry").disabled=!scalar||!g.supported||!g.pending;
  $("ls-wf-compute").disabled=!scalar||!!pendingSignature;
  $("ls-wf-status").textContent=feedback||(!g.supported?g.reason:g.pending?"Geometry edited: apply before computing.":fresh?`Result current · ${runCount||1} linked run(s)`:resultSignature?"Inputs changed. Recompute for current Results.":"Set parameters, then Compute.");
  $("ls-wf-status").classList.toggle("ls-wf-error",feedbackError||!g.supported);
  document.dispatchEvent(new CustomEvent("ls-workflow-status",{detail:{geometry:g,resultFresh:fresh,blocked:!!blocker(),selectedDomain}}));
 }
 function navigate(step){
  if(step==="geometry")document.dispatchEvent(new CustomEvent("ls-open-cad",{detail:{id:"root"}}));
  else {
   const panels={materials:"waveMaterials",mesh:"waveSize",study:"waveMode",results:"waveResults"};
   document.dispatchEvent(new CustomEvent("ls-open-wave-optics",{detail:{panel:panels[step]}}));
  }
  feedback="";feedbackError=false;update();
 }
 function run(){
  const problem=blocker();
  if(problem){feedback="Mode Analysis blocked: "+problem;feedbackError=true;pendingSignature=null;update();return false;}
  if(!$("calculateBtn")||$("calculateBtn").disabled){feedback="Solver is currently busy.";feedbackError=true;update();return false;}
  pendingSignature=signature();feedback="Computing scalar LP₀₁ mode with linked Geometry, Materials and Mesh...";feedbackError=false;update();
  $("calculateBtn").click();
  return true;
 }
 root.addEventListener("click",e=>{
  const step=e.target.closest("[data-wf-step]");if(step){navigate(step.dataset.wfStep);return;}
  if(e.target.closest("#ls-wf-compute")){run();return;}
  if(e.target.closest("#ls-wf-apply-geometry")){
   const g=geometry();
   if(!g.supported){feedback="Cannot apply: "+g.reason;feedbackError=true;update();return;}
   window.LSWaveCADService?.apply();update();return;
  }
 });
 document.addEventListener("ls-workflow-compute-request",run);
 document.addEventListener("ls-cad-state",update);
 document.addEventListener("ls-global-parameters-changed",()=>{feedback="Global Parameters changed; apply Geometry before Mode Analysis.";feedbackError=false;update();});
 document.addEventListener("ls-cad-applied",()=>{
  resultSignature=null;feedback="Geometry applied: Core and Cladding radii updated in FEM inputs.";feedbackError=false;update();
 });
 document.addEventListener("ls-cad-apply-failed",e=>{feedback=e.detail?.message||"Geometry apply failed";feedbackError=true;update();});
 document.addEventListener("ls-workflow-domain",e=>{
  if(e.detail?.domain!=="core"&&e.detail?.domain!=="cladding")return;
  selectedDomain=e.detail.domain;
  document.dispatchEvent(new CustomEvent("ls-open-wave-optics",{detail:{panel:selectedDomain==="core"?"waveCore":"waveCladding"}}));
  update();
 });
 const source=$("notice");
 if(source)new MutationObserver(()=>{
  const t=source.textContent||"";
  if(pendingSignature){
   if(source.classList.contains("error")||t.startsWith("해석 실패")){
    feedback="Mode Analysis failed: "+t;feedbackError=true;pendingSignature=null;
   }else if(t.includes("모드 계산 완료")||t.startsWith("고정 클래딩 n=")){
    const n=Number($("neff")?.textContent);
    if(Number.isFinite(n)&&n>1){
     resultSignature=pendingSignature;pendingSignature=null;runCount++;
     feedback="Mode Analysis complete · neff "+n.toFixed(9)+" · Results linked.";feedbackError=false;
    }
   }
  }else if(t.includes("모드 계산 완료")||t.startsWith("고정 클래딩 n=")){
   const n=Number($("neff")?.textContent);
   if(Number.isFinite(n)&&n>1&&scalarMode())resultSignature=signature();
  }
  update();
 }).observe(source,{subtree:true,childList:true,characterData:true,attributes:true,attributeFilter:["class"]});
 for(const id of linkedIds){
  const el=$(id);if(!el)continue;
  el.addEventListener("input",()=>{feedback="";feedbackError=false;update();});
  el.addEventListener("change",()=>{if(id==="preset")resultSignature=null;feedback="";feedbackError=false;update();});
 }
 $("resetBtn")?.addEventListener("click",()=>{resultSignature=null;feedback="";feedbackError=false;queueMicrotask(update);});
 if(source?.textContent?.includes("모드 계산 완료")&&Number($("neff")?.textContent)>1)resultSignature=signature();
 window.LSWaveWorkflowService=Object.freeze({status:()=>({blocker:blocker(),geometry:geometry(),resultFresh:!!resultSignature&&resultSignature===signature(),selectedDomain}),run,navigate});
 update();
})();
