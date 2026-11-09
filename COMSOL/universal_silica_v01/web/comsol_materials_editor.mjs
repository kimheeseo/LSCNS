/**
 * 2D Geometry material assignment editor.
 * Explicit geometry-domain preview only: custom material values are not silently
 * passed into the legacy two-domain radial FEM or full-vector Maxwell solvers.
 */
(function comsolMaterialsEditor(){
 "use strict";
 const $=id=>document.getElementById(id);
 const sidebar=document.querySelector(".layout > .sidebar");
 const tree=document.querySelector(".ls-model-tree");
 if(!sidebar||!tree||$("ls-materials-editor"))return;
 const safe=s=>String(s??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
 const materials=[];
 let activeId=null,manualPick=false,warning="",menuOpen=false,lastDomainSignature="";
 const pane=document.createElement("section");
 pane.id="ls-materials-editor";pane.className="ls-config-pane ls-materials-editor";pane.hidden=true;
 sidebar.append(pane);
 const menu=document.createElement("div");
 menu.id="ls-materials-context";menu.hidden=true;menu.setAttribute("role","menu");
 menu.innerHTML='<button type="button" data-material-command="library">Add Material from Library <small>planned</small></button><button type="button" data-material-command="browse">Browse Materials <small>planned</small></button><button type="button" data-material-command="blank">Blank Material</button><button type="button" data-material-command="layers">Layers <small>planned</small></button><button type="button" data-material-command="more">More Materials <small>planned</small></button><button type="button" data-material-command="help">Help</button>';
 document.body.append(menu);
 const showMessage=t=>{warning=t;if($("ls-material-note"))$("ls-material-note").textContent=t;};
 const geometry=()=>window.LSWaveCADService;
 function cumulativeSelections(){
  const selections=geometry()?.cumulativeSelections?.();
  return Array.isArray(selections)?selections:[];
 }
 const LEVELS=["Domain","Boundary","Point"];
 function entityLevel(material){return LEVELS.includes(material?.level)?material.level:"Domain";}
 function getEntities(level="Domain"){
  if(!LEVELS.includes(level))return [];
  try{
   const entities=geometry()?.entities()||[],output=[];
   for(const shape of entities){
    const positive=(shape.layers||[]).filter(t=>t>0).length;
    const count=level==="Domain"?positive+1:level==="Boundary"?positive+2:1;
    for(let i=0;i<count;i++){
     const id=output.length+1;
     const label=level==="Domain"
      ?positive?(i===positive?" · inner domain":" · layer "+(i+1)):" · disk"
      :level==="Boundary"?(i===0?" · outer boundary":" · inner boundary "+i)
      :" · center point";
     output.push({id,entityKey:shape.entityKey,owner:shape.owner,source:shape.source,
       label:shape.entityKey+label,layer:i,level});
    }
   }
   return output;
  }catch(error){warning="Geometry unavailable: "+error.message;return [];}
 }
 function getDomains(){return getEntities("Domain");}
 function explicitIds(material,level){
  return level==="Domain"?material.domainIds||[]:material.selectedIdsByLevel?.[level]||[];
 }
 function assignedEntities(material,entities){
  if(material.mode==="all")return entities;
  if(material.mode==="cumulative"){
   const cumulative=cumulativeSelections().find(row=>row.id===material.cumulativeId);
   const owners=new Set(cumulative?.featureIds||[]);
   return entities.filter(entity=>owners.has(entity.owner));
  }
  const ids=new Set(explicitIds(material,entityLevel(material)));
  return entities.filter(entity=>ids.has(entity.id));
 }
 function assignedDomains(material,domains){
  // Boundary/Point selections do not count as assigned material domains.
  if(entityLevel(material)!=="Domain")return [];
  return assignedEntities(material,domains);
 }
 function selectionTitle(material){
  if(material.mode==="all")return "All "+entityLevel(material).toLowerCase()+"s";
  if(material.mode==="cumulative")return cumulativeSelections().find(v=>v.id===material.cumulativeId)?.label||"Missing cumulative selection";
  return explicitIds(material,entityLevel(material)).join(", ")||"None";
 }
 function overviewBody(domains){
  const all=materials.map(m=>'<tr><td>◈ '+safe(m.label)+'</td><td>'+safe(selectionTitle(m))+'</td></tr>').join("");
  const covered=new Set(materials.flatMap(m=>assignedDomains(m,domains).map(d=>d.id)));
  const missing=domains.filter(d=>!covered.has(d.id)).map(d=>d.id);
  const missingText=missing.length?("Domains "+missing.join(", ")):"None";
  return '<details class="ls-material-section" open><summary>Material Overview</summary><div class="ls-material-section-inner">'+
   '<table class="ls-material-overview"><thead><tr><th>Material</th><th>Selection</th></tr></thead><tbody>'+
   '<tr><td><span class="ls-material-warning">⛔</span> Entities needing a material</td><td>'+safe(missingText)+'</td></tr>'+all+'</tbody></table>'+
   '<p class="ls-material-explainer">Geometry preview: '+domains.length+' selectable region(s). Overlapping circles are not Boolean-unioned and the displayed IDs are not validated COMSOL FEM domains.</p>'+
   '<button class="ls-material-primary" type="button" data-material-command="blank">＋ Blank Material</button></div></details>';
 }
 function detailsBody(material,domains){
  const level=entityLevel(material),entities=getEntities(level);
  const selection=assignedEntities(material,entities);
  const checkedIds=new Set(selection.map(e=>e.id));
  const allowManual=material.mode==="manual";
  const name=level.toLowerCase()+"s";
  const cumul=cumulativeSelections();
  const optGroups=cumul.length
   ?'<optgroup label="Cumulative Selections">'+cumul.map(row=>
     '<option data-cumulative-selection="'+safe(row.id)+'" value="cumulative:'+safe(row.id)+'" '+
      (material.mode==="cumulative"&&material.cumulativeId===row.id?"selected":"")+'>'+
      safe(row.label)+'</option>').join("")+'</optgroup>'
   :"";
  const items=entities.map(entity=>{
   return '<label class="ls-material-domain-option"><input type="checkbox" data-material-domain="'+entity.id+'" '+
    (checkedIds.has(entity.id)?"checked":"")+' '+(allowManual?"":"disabled")+'/>'+
    '<span><b>'+entity.id+'</b> '+safe(entity.label)+'</span></label>';
  }).join("");
  const nameOptions=LEVELS.map(option=>'<option value="'+option+'" '+(level===option?"selected":"")+'>'+option+'</option>').join("");
  const selectedMode=material.mode==="cumulative"?"cumulative:"+material.cumulativeId:material.mode;
  return '<div class="ls-setting-line"><label for="ls-material-label">Label</label><input id="ls-material-label" value="'+safe(material.label)+'"/></div>'+
   '<details class="ls-material-section" open><summary>Geometric Entity Selection</summary><div class="ls-material-section-inner">'+
   '<label class="ls-material-field"><span>Geometric entity level:</span><select id="ls-material-entity-level">'+nameOptions+'</select></label>'+
   '<label class="ls-material-field"><span>Selection:</span><select id="ls-material-select-mode">'+
   '<option value="all" '+(selectedMode==="all"?"selected":"")+'>All '+name+'</option>'+
   '<option value="manual" '+(selectedMode==="manual"?"selected":"")+'>Manual selection</option>'+
   optGroups+'</select></label>'+
   '<div class="ls-material-selection-row"><button type="button" id="ls-material-pick" aria-pressed="'+(manualPick&&material.mode!=="cumulative")+'" '+
   (material.mode==="cumulative"?"disabled ":"")+'title="Activate Selection in Graphics">▣</button>'+
   '<div id="ls-material-domain-list">'+(items||"<p>No "+name+" in current geometry</p>")+'</div></div>'+
   '<p class="ls-material-help">Selected '+selection.length+' of '+entities.length+' '+name+' in the Geometry preview. '+
    (material.mode==="cumulative"?"Using named selection: "+safe(selectionTitle(material))+". Changes to its contributing Geometry features update this list.":
    manualPick?"Graphics selection is active: click a circle to toggle associated region(s).":
    "Choose Manual selection to edit individual "+name+".")+'</p>'+
   (level!=="Domain"?'<p class="ls-material-help">Boundary and Point are selectable geometric entities only; they are not material-filled FEM domains.</p>':"")+
   '</div></details>'+
   '<details class="ls-material-section"><summary>Override</summary><p class="ls-material-help">Not linked to the FEM material model.</p></details>'+
   '<details class="ls-material-section"><summary>Material Properties</summary><p class="ls-material-help">Isotropic refractive index. Enter a real value (e.g. 1.45) and imaginary value (e.g. 0) for material preview.</p></details>'+
   '<details class="ls-material-section" open><summary>Material Contents</summary><div class="ls-material-section-inner">'+
   '<table class="ls-material-values"><thead><tr><th>Property</th><th>Variable</th><th>Value</th><th>Unit</th><th>Property group</th></tr></thead><tbody>'+
   '<tr><td>Refractive index, real part</td><td>n_iso</td><td><input id="ls-material-real" aria-label="Refractive index real part" value="'+safe(material.real)+'" placeholder="e.g. 1.45"/></td><td>1</td><td>Refractive index</td></tr>'+
   '<tr><td>Refractive index, imaginary part</td><td>ki_iso</td><td><input id="ls-material-imag" aria-label="Refractive index imaginary part" value="'+safe(material.imag)+'"/></td><td>1</td><td>Refractive index</td></tr>'+
   '</tbody></table><p class="ls-material-help">This Blank Material is a Geometry-domain annotation. It does not change the current radial FEM n(r) or compute a full-vector EWFD solution.</p></div></details>';
 }
 function treeRender(){
  const host=$("ls-cad-material-list");if(!host)return;
  host.replaceChildren();
  for(const material of materials){
   const button=document.createElement("button");
   button.type="button";button.dataset.lsBlankMaterial=material.id;
   button.className="ls-tree-leaf"+(material.id===activeId&&!pane.hidden?" active":"");
   button.textContent="◈ "+material.label+" ("+material.id+")";
   host.append(button);
  }
 }
 function render(){
  const domains=getDomains(),material=materials.find(m=>m.id===activeId);
  lastDomainSignature=domains.map(x=>x.id+":"+x.entityKey+":"+x.layer).join("|");
  pane.innerHTML='<div class="ls-config-heading"><div><strong>Settings</strong><small id="ls-material-subtitle">'+(material?"Material":"Materials")+'</small></div><button type="button" id="ls-material-close">← Fiber settings</button></div>'+
   '<div class="ls-material-editor-body">'+(material?detailsBody(material,domains):'<div class="ls-setting-line"><label>Label</label><input readonly value="Materials"/></div>'+overviewBody(domains))+
   '<p class="ls-config-status" id="ls-material-note" role="status">'+safe(warning||"Material assignment is a Geometry preview until a full-vector domain mesh and material solver are implemented.")+'</p></div>';
  treeRender();
  highlight();
 }
 function selectedKeys(){
  const m=materials.find(x=>x.id===activeId);
  if(!m)return [];
  return [...new Set(assignedEntities(m,getEntities(entityLevel(m))).map(x=>x.entityKey))];
 }
 function highlight(){
  if(pane.hidden)return;
  document.dispatchEvent(new CustomEvent("ls-material-highlight",{detail:{entityKeys:selectedKeys()}}));
 }
 function pickChanged(enabled){
  const material=materials.find(m=>m.id===activeId);
  manualPick=!!enabled&&material?.mode!=="cumulative";
  document.dispatchEvent(new CustomEvent("ls-material-picking",{detail:{enabled:manualPick,entityKeys:manualPick?selectedKeys():[]}}));
  if($("ls-material-pick")){$("ls-material-pick").setAttribute("aria-pressed",String(manualPick));$("ls-material-pick").classList.toggle("active",manualPick);}
 }
 function open(id=null){
  activeId=id;
  menu.hidden=true;
  const cad=$("ls-cad-pane"),wave=$("ls-wave-settings"),other=$("ls-config-pane");
  if(other)other.hidden=true;
  if(cad)cad.hidden=true;
  if(wave)wave.hidden=true;
  pane.hidden=false;
  sidebar.classList.add("ls-config-active");
  document.body.classList.remove("ls-wave-optics-active","ls-cad-active");
  document.body.classList.add("ls-configuration-open","ls-material-active");
  document.dispatchEvent(new CustomEvent("ls-open-material-editor",{detail:{id}}));
  render();
  pickChanged(!!id);
 }
 function close(){
  if(pane.hidden)return;
  pickChanged(false);
  pane.hidden=true;menu.hidden=true;activeId=null;
  document.body.classList.remove("ls-material-active","ls-configuration-open");
  sidebar.classList.remove("ls-config-active");
  const cadGraphics=$("ls-cad-graphics");if(cadGraphics)cadGraphics.hidden=true;
  treeRender();
 }
 function blank(){
  let i=1;while(materials.some(m=>m.id==="mat"+i))i++;
  const material={id:"mat"+i,label:"Material "+i,mode:"all",level:"Domain",cumulativeId:null,domainIds:[],selectedIdsByLevel:{Boundary:[],Point:[]},real:"",imag:"0"};
  materials.push(material);
  warning="Blank Material created. Assign a domain and refractive index values to configure this preview.";
  open(material.id);
 }
 function context(e){
  e.preventDefault();e.stopPropagation();
  menu.hidden=false;
  menu.style.left=Math.max(8,Math.min(e.clientX,innerWidth-260))+"px";
  menu.style.top=Math.max(8,Math.min(e.clientY,innerHeight-225))+"px";
 }
 tree.addEventListener("contextmenu",e=>{
  if(e.target.closest('summary[data-action="waveMaterials"]')||e.target.closest("[data-ls-blank-material]"))context(e);
 });
 tree.addEventListener("click",e=>{
  const material=e.target.closest("[data-ls-blank-material]");
  if(material){e.preventDefault();e.stopImmediatePropagation();open(material.dataset.lsBlankMaterial);return;}
  const summary=e.target.closest('summary[data-action="waveMaterials"]');
  if(summary&&e.clientX-summary.getBoundingClientRect().left>=22){
   e.preventDefault();e.stopImmediatePropagation();open();
  }
 },true);
 menu.addEventListener("click",e=>{
  const cmd=e.target.closest("[data-material-command]")?.dataset.materialCommand;
  if(!cmd)return;
  menu.hidden=true;
  if(cmd==="blank"){blank();return;}
  if(cmd==="help"){window.open("https://github.com/kimheeseo/LSCNS/tree/main/COMSOL/universal_silica_v01/web/","_blank","noopener");return;}
  warning="This library action is not implemented yet. Use Blank Material for Geometry preview.";
  open();
 });
 document.addEventListener("pointerdown",e=>{if(!menu.hidden&&!menu.contains(e.target))menu.hidden=true;});
 pane.addEventListener("click",e=>{
  if(e.target.closest("#ls-material-close")){close();return;}
  if(e.target.closest('[data-material-command="blank"]')){blank();return;}
  if(e.target.closest("#ls-material-pick")){pickChanged(!manualPick);render();return;}
 });
 pane.addEventListener("input",e=>{
  const material=materials.find(x=>x.id===activeId);if(!material)return;
  if(e.target.id==="ls-material-label"){material.label=e.target.value;treeRender();}
  if(e.target.id==="ls-material-real")material.real=e.target.value;
  if(e.target.id==="ls-material-imag")material.imag=e.target.value;
 });
 pane.addEventListener("change",e=>{
  const material=materials.find(x=>x.id===activeId);if(!material)return;
  if(e.target.id==="ls-material-entity-level"){
   const newLevel=e.target.value;
   if(!LEVELS.includes(newLevel))return;
   material.level=newLevel;
   // Keep separately selected IDs per entity level. Switching levels must
   // never reuse Domain IDs as Boundary or Point assignments.
   render();pickChanged(manualPick&&material.mode!=="cumulative");
   return;
  }
  if(e.target.id==="ls-material-select-mode"){
   const value=e.target.value;
   if(value.startsWith("cumulative:")){
    const id=value.slice("cumulative:".length);
    if(!cumulativeSelections().some(row=>row.id===id)){
     showMessage("Named Cumulative Selection is no longer available.");render();return;
    }
    material.mode="cumulative";material.cumulativeId=id;
    render();pickChanged(false);
   }else if(value==="all"||value==="manual"){
    material.mode=value;material.cumulativeId=null;
    if(value==="manual"&&!explicitIds(material,entityLevel(material)).length){
     const ids=getEntities(entityLevel(material)).map(x=>x.id);
     if(entityLevel(material)==="Domain")material.domainIds=ids;
     else{
      material.selectedIdsByLevel??={Boundary:[],Point:[]};
      material.selectedIdsByLevel[entityLevel(material)]=ids;
     }
    }
    render();pickChanged(manualPick);
   }
   return;
  }
  const box=e.target.closest("[data-material-domain]");
  if(box&&material.mode==="manual"){
   const level=entityLevel(material),id=Number(box.dataset.materialDomain);
   const ids=new Set(explicitIds(material,level));
   if(box.checked)ids.add(id);else ids.delete(id);
   const saved=[...ids].sort((a,b)=>a-b);
   if(level==="Domain")material.domainIds=saved;
   else{
    material.selectedIdsByLevel??={Boundary:[],Point:[]};
    material.selectedIdsByLevel[level]=saved;
   }
   render();pickChanged(manualPick);
  }
 });
 document.addEventListener("ls-material-geometry-click",e=>{
  if(!manualPick||pane.hidden)return;
  const material=materials.find(x=>x.id===activeId);if(!material||material.mode==="cumulative")return;
  const level=entityLevel(material),entities=getEntities(level);
  const matched=entities.filter(entity=>entity.entityKey===e.detail?.entityKey);
  if(!matched.length)return;
  const current=new Set(assignedEntities(material,entities).map(x=>x.id));
  const allSelected=matched.every(d=>current.has(d.id));
  for(const d of matched){if(allSelected)current.delete(d.id);else current.add(d.id);}
  material.mode="manual";material.cumulativeId=null;
  const ids=[...current].sort((a,b)=>a-b);
  if(level==="Domain")material.domainIds=ids;
  else{
   material.selectedIdsByLevel??={Boundary:[],Point:[]};
   material.selectedIdsByLevel[level]=ids;
  }
  warning="Changed "+level.toLowerCase()+" preview selection for "+e.detail.entityKey+" ("+matched.length+" entity/ies).";
  render();pickChanged(true);
 });
 document.addEventListener("ls-cad-cumulative-changed",()=>{
  // Refresh the named-selection options and the currently displayed assignments.
  if(!pane.hidden)render();
 });
 document.addEventListener("ls-cad-state",()=>{
  if(pane.hidden)return;
  const domains=getDomains(),signature=domains.map(x=>x.id+":"+x.entityKey+":"+x.layer).join("|");
  if(signature===lastDomainSignature)return;
  lastDomainSignature=signature;
  render();
 });
 document.addEventListener("ls-wizard-finished",()=>{materials.length=0;activeId=null;treeRender();});
 document.addEventListener("ls-open-cad",close);
 document.addEventListener("ls-open-wave-optics",close);
 document.addEventListener("ls-open-settings",close);
 window.LSWaveMaterialsEditor=Object.freeze({list:()=>materials.map(m=>({...m,domainIds:[...m.domainIds],selectedIdsByLevel:{...m.selectedIdsByLevel}})),domains:getDomains,entities:getEntities});
 treeRender();
})();
