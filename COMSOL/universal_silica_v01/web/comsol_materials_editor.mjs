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
 let activeId=null,manualPick=false,warning="",menuOpen=false,lastDomainSignature="",contextMaterialId=null,propertyContextName="",refractiveDetailsExpanded=false,activeRefractiveRow=null;
 const expandedPropertyGroups=new Set(["Basic Properties","Electromagnetic Models"]);
 const pane=document.createElement("section");
 pane.id="ls-materials-editor";pane.className="ls-config-pane ls-materials-editor";pane.hidden=true;
 sidebar.append(pane);
 const menu=document.createElement("div");
 menu.id="ls-materials-context";menu.hidden=true;menu.setAttribute("role","menu");
 menu.innerHTML='<button type="button" data-material-command="library">Add Material from Library <small>planned</small></button><button type="button" data-material-command="browse">Browse Materials <small>planned</small></button><button type="button" data-material-command="blank">Blank Material</button><button type="button" data-material-command="layers">Layers <small>planned</small></button><button type="button" data-material-command="more">More Materials <small>planned</small></button><button type="button" data-material-command="help">Help</button>';
 document.body.append(menu);
 const propertyMenu=document.createElement("div");
 propertyMenu.id="ls-material-property-context";propertyMenu.hidden=true;propertyMenu.setAttribute("role","menu");
 propertyMenu.innerHTML='<button type="button" id="ls-material-property-add">＋ Add to Material</button>';
 document.body.append(propertyMenu);
 const deleteDialog=document.createElement("dialog");
 deleteDialog.id="ls-material-delete-dialog";
 deleteDialog.innerHTML='<form method="dialog"><h3>Confirm Delete</h3><p>Delete the selected material and its Geometry domain assignments?</p><div class="ls-material-dialog-actions"><button type="submit" value="yes">Yes</button><button type="submit" value="no">No</button></div></form>';
 document.body.append(deleteDialog);
 deleteDialog.querySelector("form").addEventListener("submit",e=>{
  if(e.submitter?.value!=="yes")return;
  const id=contextMaterialId||activeId,at=materials.findIndex(x=>x.id===id);
  if(at<0)return;
  const label=materials[at].label;
  materials.splice(at,1);contextMaterialId=null;
  warning=label+" deleted. The former domain assignments are no longer present.";
  activeId=null;manualPick=false;render();pickChanged(false);
 });

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
  if(level==="Domain"&&geometry()?.domains)return geometry().domains().map(domain=>({id:domain.id,entityKey:domain.entityKey,owner:domain.owner,source:domain.source,label:domain.label,level:"Domain",domainKey:domain.key}));
  try{
   const entities=geometry()?.entities()||[],output=[];
   for(const shape of entities){
    const positive=(shape.layers||[]).filter(t=>t>0).length;
    const count=level==="Point"?1:positive+1; // Outer circle plus one interface per inward layer; no fictitious inner ring.
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

 // A refractive-index property is a value-bearing Geometry material property.
 // It is evaluated with the SAME Parameters 1 expression evaluator as the
 // Circle radius / Layer expressions; not just a text label.
 function evaluateIndexPart(expression,kind){
  const value=String(expression??"").trim();
  if(!value)return {valid:false,error:(kind==="real"?"Refractive index real part":"Refractive index imaginary part")+" is required."};
  try{
   const result=window.LSWaveParameterService?.calculate?.(value);
   if(!Number.isFinite(result))throw Error("Expression is not a finite number.");
   if(kind==="real"&&result<=0)throw Error("Real refractive index n must be greater than zero.");
   if(kind==="imag"&&result<0)throw Error("Imaginary refractive index ki must be zero or positive.");
   return {valid:true,value:result};
  }catch(error){return {valid:false,error:error.message};}
 }
 function resolvedIndex(material){
  const real=String(material?.real??"").trim(),imag=String(material?.imag??"").trim();
  const n=evaluateIndexPart(real,"real"),k=evaluateIndexPart(imag,"imag");
  if(!n.valid||!k.valid)return {
   valid:false,realValid:n.valid,imagValid:k.valid,
   error:[n.valid?null:n.error,k.valid?null:k.error].filter(Boolean).join(" "),
   expressionReal:real,expressionImag:imag
  };
  return {valid:true,realValid:true,imagValid:true,real:n.value,imag:k.value,
   expressionReal:real,expressionImag:imag};
 }
 function previewAssignments(){
  const domains=getDomains(),out=[];
  for(const domain of domains){
   const assigned=materials.filter(m=>assignedDomains(m,domains).some(e=>e.id===domain.id));
   const material=assigned.at(-1);
   if(!material){out.push({domainId:domain.id,materialId:null,status:"unassigned"});continue;}
   const n=resolvedIndex(material);
   out.push({domainId:domain.id,domainKey:domain.domainKey,materialId:material.id,materialLabel:material.label,
    real:n.valid?n.real:null,imag:n.valid?n.imag:null,status:assigned.length>1?"overlap":n.valid?"valid":"invalid"});
  }
  return out;
 }
 const BASIC_EM_PROPERTIES=[
  {id:"relative-permeability",name:"Relative permeability",variable:"mur_iso"},
  {id:"relative-permittivity",name:"Relative permittivity",variable:"epsilonr_iso"}
 ];
 function basicEMState(material,property){
  const expression=String(material?.basicPropertyValues?.[property.id]??"").trim();
  if(!expression)return {valid:false,error:property.name+" value is required.",expression};
  try{
   const value=window.LSWaveParameterService?.calculate?.(expression);
   if(!Number.isFinite(value))throw Error("Expression must evaluate to a finite real scalar.");
   return {valid:true,value,expression};
  }catch(error){return {valid:false,error:error.message,expression};}
 }
 const MATERIAL_PROPERTY_GROUPS=[
  {name:"Basic Properties",items:["Density","Heat capacity at constant pressure","Thermal conductivity","Relative permeability","Relative permittivity"]},
  {name:"Acoustics",items:[]},
  {name:"Electrochemistry",items:[]},
  {name:"Electromagnetic Models",items:[
   "B-H curve","B-S curve","Coercive magnetic field","Dielectric losses","E-J characteristic",
   "Effective B-H curve","H-B curve","Jiles-Atherton model parameters",
   "Linearized resistivity","Loss tangent, dissipation factor","Loss tangent, loss angle",
   "Magnetic losses","Permanent magnet shifted B-H curve","Permanent magnet shifted H-B curve",
   "Refractive index","Remanent flux density"
  ]},
  {name:"External Material Parameters",items:[]},
  {name:"Fluid Flow",items:[]},
  {name:"Gas Models",items:[]},
  {name:"Geometric Properties",items:[]},
  {name:"Magnetostrictive Models",items:[]},
  {name:"Piezoelectric Models",items:[]},
  {name:"Optical Dispersion Models",items:[]},
  {name:"Thermo-Optic Dispersion Models",items:[]}
 ];
 function propertiesBrowser(){
  return '<div id="ls-material-properties-tree" role="tree" aria-label="Material Properties">'+
   MATERIAL_PROPERTY_GROUPS.map(g=>'<details class="ls-material-property-group" data-ls-property-group="'+safe(g.name)+'" '+(expandedPropertyGroups.has(g.name)?"open":"")+'><summary>'+safe(g.name)+'</summary>'+
    '<div class="ls-material-property-children">'+(g.items.length?g.items.map(name=>{
     if(name==="Refractive index"){
      return '<details class="ls-material-refractive-node" '+(refractiveDetailsExpanded?"open":"")+'>'+
       '<summary data-ls-property="Refractive index" tabindex="0" title="Click to expand; right-click for Add to Material">◈ Refractive index</summary>'+
       '<div class="ls-material-refractive-children">'+
       '<button type="button" data-ls-refractive-focus="real">Refractive index, real part (n)</button>'+
       '<button type="button" data-ls-refractive-focus="imag">Refractive index, imaginary part (ki)</button>'+
       '</div></details>';
     }
     if(BASIC_EM_PROPERTIES.some(property=>property.name===name))
      return '<button type="button" role="treeitem" class="ls-material-property-entry" data-ls-property="'+safe(name)+'" title="Right-click → Add to Material">◇ '+safe(name)+'</button>';
     return '<button type="button" role="treeitem" class="ls-material-property-entry" data-ls-property="'+safe(name)+'" disabled title="Property editor not implemented yet">◇ '+safe(name)+'</button>';
    }).join(""):'<span class="ls-material-empty-property">No editable properties in this prototype.</span>')+'</div></details>').join("")+
   '</div><p class="ls-material-help">Right-click <b>Relative permeability</b>, <b>Relative permittivity</b>, or <b>Refractive index</b> → Add to Material. Enter values or Parameters expressions in Material Contents to receive valid Status checks.</p>';
 }
 function materialContents(material){
  const added=material.properties?.includes("refractive-index");
  if(!added)return '<p class="ls-material-help">No material properties added. Right-click Refractive index under Material Properties → Add to Material.</p>';
  const state=resolvedIndex(material);
  const statusIcon=valid=>valid
   ?'<span class="ls-material-valid" role="img" aria-label="Valid property" title="Valid expression">✓</span>'
   :'<span class="ls-material-invalid" role="img" aria-label="Invalid property" title="Invalid or missing expression">!</span>';
  const realIcon=statusIcon(state.realValid),imagIcon=statusIcon(state.imagValid);
  const num=value=>Number.isFinite(value)?Number(value.toPrecision(9)).toString():"—";
  return '<table class="ls-material-values"><thead><tr><th>Status</th><th>Property</th><th>Variable</th><th>Value / Expression</th><th>Unit</th><th>Property group</th></tr></thead><tbody>'+
   '<tr data-ls-material-index-row="real" class="'+(activeRefractiveRow==="real"?"is-active":"")+'" aria-selected="'+(activeRefractiveRow==="real")+'"><td data-ls-real-state>'+realIcon+'</td><td><button type="button" class="ls-material-index-label" data-ls-material-label="real">Refractive index, real part</button></td><td>n_iso</td>'+
   '<td><input id="ls-material-real" aria-label="Refractive index real part" value="'+safe(material.real)+'" placeholder="e.g. nbg or 1.45" /></td><td>1</td><td>Refractive index</td></tr>'+
   '<tr data-ls-material-index-row="imag" class="'+(activeRefractiveRow==="imag"?"is-active":"")+'" aria-selected="'+(activeRefractiveRow==="imag")+'"><td data-ls-imag-state>'+imagIcon+'</td><td><button type="button" class="ls-material-index-label" data-ls-material-label="imag">Refractive index, imaginary part</button></td><td>ki_iso</td>'+
   '<td><input id="ls-material-imag" aria-label="Refractive index imaginary part" value="'+safe(material.imag)+'" placeholder="0" /></td><td>1</td><td>Refractive index</td></tr>'+
   '</tbody></table>'+
   '<div id="ls-material-evaluated" class="ls-material-evaluated '+(state.valid?"is-valid":"is-invalid")+'" aria-live="polite">'+
    (state.valid?'Applied to assigned Geometry domains: <b>n_iso = '+num(state.real)+'</b>, <b>ki_iso = '+num(state.imag)+'</b>.':
     'Not applied: '+safe(state.error))+'</div>'+
   '<p class="ls-material-help">Parameters 1 expressions (e.g. nbg, nair, nbg+0.005) are evaluated numerically. The evaluated index is attached to the Geometry preview domains, but is not passed to the current radial FEM or a full-vector EWFD solver.</p>';
 }
 function activateRefractiveRow(which,focusInput=false){
  if(which!=="real"&&which!=="imag")return;
  activeRefractiveRow=which;
  for(const row of pane.querySelectorAll("[data-ls-material-index-row]")){
   const active=row.dataset.lsMaterialIndexRow===which;
   row.classList.toggle("is-active",active);
   row.setAttribute("aria-selected",String(active));
  }
  if(focusInput){
   const input=$(which==="real"?"ls-material-real":"ls-material-imag");
   const panel=input?.closest("details");
   if(panel)panel.open=true;
   input?.focus();
  }
 }
 function updateMaterialEvaluation(material){
  const resolved=resolvedIndex(material),box=$("ls-material-evaluated");
  if(box){
   box.classList.toggle("is-valid",resolved.valid);box.classList.toggle("is-invalid",!resolved.valid);
   box.textContent=resolved.valid?
    "Applied to assigned Geometry domains: n_iso = "+Number(resolved.real.toPrecision(9))+", ki_iso = "+Number(resolved.imag.toPrecision(9)):
    "Not applied: "+resolved.error;
  }
  for(const [kind,id,selector] of [
   ["real","ls-material-real","[data-ls-real-state]"],
   ["imag","ls-material-imag","[data-ls-imag-state]"]
  ]){
   const valid=kind==="real"?resolved.realValid:resolved.imagValid;
   const input=$(id);
   if(input)input.setAttribute("aria-invalid",valid?"false":"true");
   const cell=pane.querySelector(selector);
   if(cell){
    cell.innerHTML=valid
     ?'<span class="ls-material-valid" role="img" aria-label="Valid property" title="Valid expression">✓</span>'
     :'<span class="ls-material-invalid" role="img" aria-label="Invalid property" title="Invalid or missing expression">!</span>';
   }
  }
  document.dispatchEvent(new CustomEvent("ls-material-property-changed",{detail:{
   materialId:material.id,evaluated:resolved,assignedDomains:assignedDomains(material,getDomains()).map(d=>d.id)
  }}));
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
   (material.mode!=="manual"?"disabled ":"")+'title="Activate Selection in Graphics">▣</button>'+
   '<div id="ls-material-domain-list">'+(items||"<p>No "+name+" in current geometry</p>")+'</div></div>'+
   '<div class="ls-material-selection-actions"><button type="button" id="ls-material-select-all" '+(allowManual?"":"disabled")+' title="Select all (Ctrl+A in Graphics)">All</button><button type="button" id="ls-material-select-clear" '+(allowManual?"":"disabled")+' title="Clear selection (Ctrl+D in Graphics)">Clear</button></div>'+
   '<p class="ls-material-help">Selected '+selection.length+' of '+entities.length+' '+name+' in the Geometry preview. '+
    (material.mode==="cumulative"?"Using named selection: "+safe(selectionTitle(material))+". Changes to its contributing Geometry features update this list.":
    manualPick?"Manual picking active: click one domain to select or deselect it.":
    "Choose Manual selection to edit individual "+name+".")+'</p>'+
   (level!=="Domain"?'<p class="ls-material-help">Boundary and Point are selectable geometric entities only; they are not material-filled FEM domains.</p>':"")+
   '</div></details>'+
   '<details class="ls-material-section"><summary>Override</summary><p class="ls-material-help">Not linked to the FEM material model.</p></details>'+
   '<div class="ls-material-controls"><button type="button" data-material-command="blank">＋ Add Material</button>'+
   '<button type="button" data-material-command="duplicate">Duplicate</button>'+
   '<button type="button" data-material-command="delete">Delete Material</button></div>'+
   '<details class="ls-material-section" id="ls-material-properties-section"><summary>Material Properties</summary><div class="ls-material-section-inner">'+propertiesBrowser()+'</div></details>'+
   '<details class="ls-material-section" open><summary>Material Contents</summary><div class="ls-material-section-inner">'+materialContents(material)+'</div></details>';
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
  if(material)updateMaterialEvaluation(material);
 }
 function selectedKeys(){
  const m=materials.find(x=>x.id===activeId);
  if(!m)return [];
  return [...new Set(assignedEntities(m,getEntities(entityLevel(m))).map(x=>x.entityKey))];
 }
 function highlight(){
  if(pane.hidden)return;
  const material=materials.find(m=>m.id===activeId);
  document.dispatchEvent(new CustomEvent("ls-material-highlight",{detail:{
   entityKeys:selectedKeys(),
   domainIds:material&&entityLevel(material)==="Domain"?assignedEntities(material,getDomains()).map(d=>d.id):[],
   level:entityLevel(material)
  }}));
 }
 function pickChanged(enabled){
  const material=materials.find(m=>m.id===activeId);
  manualPick=!!enabled&&material?.mode==="manual";
  document.dispatchEvent(new CustomEvent("ls-material-picking",{detail:{
   enabled:manualPick,entityKeys:manualPick?selectedKeys():[],
   domainIds:material&&entityLevel(material)==="Domain"?assignedEntities(material,getDomains()).map(d=>d.id):[],
   level:entityLevel(material)
  }}));
  if($("ls-material-pick")){$("ls-material-pick").setAttribute("aria-pressed",String(manualPick));$("ls-material-pick").classList.toggle("active",manualPick);}
 }
 function open(id=null){
  if(activeId!==id)activeRefractiveRow=null;
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
  pickChanged(!!id&&materials.find(m=>m.id===activeId)?.mode==="manual");
 }
 function close(){
  if(pane.hidden)return;
  pickChanged(false);
  pane.hidden=true;menu.hidden=true;propertyMenu.hidden=true;activeId=null;contextMaterialId=null;
  document.body.classList.remove("ls-material-active","ls-configuration-open");
  sidebar.classList.remove("ls-config-active");
  const cadGraphics=$("ls-cad-graphics");if(cadGraphics)cadGraphics.hidden=true;
  treeRender();
 }
 function blank(){
  let i=1;while(materials.some(m=>m.id==="mat"+i))i++;
  const material={id:"mat"+i,label:"Material "+i,mode:"all",level:"Domain",cumulativeId:null,domainIds:[],selectedIdsByLevel:{Boundary:[],Point:[]},properties:["refractive-index"],real:"",imag:"0"};
  materials.push(material);
  warning="Blank Material created. Assign a domain and refractive index values to configure this preview.";
  open(material.id);
 }

 function duplicate(){
  const original=materials.find(m=>m.id===(contextMaterialId||activeId));
  if(!original){blank();return;}
  let i=1;while(materials.some(m=>m.id==="mat"+i))i++;
  const copy={...original,id:"mat"+i,label:"Material "+i,
   domainIds:[...original.domainIds],
   selectedIdsByLevel:{Boundary:[...(original.selectedIdsByLevel?.Boundary||[])],Point:[...(original.selectedIdsByLevel?.Point||[])]},
   properties:[...(original.properties||["refractive-index"])]};
  materials.push(copy);warning="Duplicated "+original.label+" as "+copy.label+". Select its domain assignment independently.";
  contextMaterialId=null;open(copy.id);
 }
 function showDelete(){
  const id=contextMaterialId||activeId;
  if(!materials.some(m=>m.id===id))return;
  deleteDialog.returnValue="no";deleteDialog.showModal();
 }
 function addRefractiveIndex(){
  const material=materials.find(m=>m.id===activeId);
  if(!material)return;
  material.properties??=[];
  if(!material.properties.includes("refractive-index"))material.properties.push("refractive-index");
  if(material.real===undefined)material.real="";
  if(material.imag===undefined)material.imag="0";
  warning="Refractive index available in "+material.label+". Set n_iso (e.g. nbg) and ki_iso (0).";
  propertyMenu.hidden=true;
  refractiveDetailsExpanded=true;
  render();
  const section=$("ls-material-properties-section");
  if(section)section.open=true;
  const node=pane.querySelector(".ls-material-refractive-node");
  if(node)node.open=true;
 }
 function context(e){
  e.preventDefault();e.stopPropagation();
  contextMaterialId=e.target.closest("[data-ls-blank-material]")?.dataset.lsBlankMaterial||null;
  menu.querySelectorAll('[data-material-command="delete"],[data-material-command="duplicate"]').forEach(b=>b.remove());
  if(contextMaterialId){
   menu.insertAdjacentHTML("beforeend",'<button type="button" data-material-command="duplicate">Duplicate</button><button type="button" data-material-command="delete">Delete</button>');
  }
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
  if(cmd==="blank"){contextMaterialId=null;blank();return;}
  if(cmd==="duplicate"){duplicate();return;}
  if(cmd==="delete"){showDelete();return;}
  if(cmd==="help"){window.open("https://github.com/kimheeseo/LSCNS/tree/main/COMSOL/universal_silica_v01/web/","_blank","noopener");return;}
  warning="This library action is not implemented yet. Use Blank Material for Geometry preview.";
  open();
 });
 document.addEventListener("pointerdown",e=>{
  if(!menu.hidden&&!menu.contains(e.target))menu.hidden=true;
  if(!propertyMenu.hidden&&!propertyMenu.contains(e.target))propertyMenu.hidden=true;
 });
 pane.addEventListener("contextmenu",e=>{
  const property=e.target.closest('[data-ls-property="Refractive index"]');
  if(!property)return;
  e.preventDefault();e.stopPropagation();
  propertyContextName=property.dataset.lsProperty;
  propertyMenu.style.left=Math.max(6,Math.min(e.clientX,window.innerWidth-200))+"px";
  propertyMenu.style.top=Math.max(6,Math.min(e.clientY,window.innerHeight-65))+"px";
  propertyMenu.hidden=false;
 });
 propertyMenu.addEventListener("click",e=>{
  if(e.target.closest("#ls-material-property-add")&&propertyContextName==="Refractive index")addRefractiveIndex();
 });
 pane.addEventListener("click",e=>{
  const label=e.target.closest("[data-ls-material-label]");
  if(label){activateRefractiveRow(label.dataset.lsMaterialLabel,true);return;}
  if(e.target.closest("#ls-material-close")){close();return;}
  if(e.target.closest('[data-material-command="blank"]')){contextMaterialId=null;blank();return;}
  if(e.target.closest('[data-material-command="duplicate"]')){contextMaterialId=activeId;duplicate();return;}
  if(e.target.closest('[data-material-command="delete"]')){contextMaterialId=activeId;showDelete();return;}
  const field=e.target.closest("[data-ls-refractive-focus]");
  if(field){
   const material=materials.find(m=>m.id===activeId);
   if(material&&!material.properties?.includes("refractive-index"))addRefractiveIndex();
   const input=$(field.dataset.lsRefractiveFocus==="real"?"ls-material-real":"ls-material-imag");
   const contents=input?.closest("details");
   if(contents)contents.open=true;
   input?.focus();
   input?.select();
   activateRefractiveRow(field.dataset.lsRefractiveFocus);
   return;
  }
  if(e.target.closest("#ls-material-pick")){pickChanged(!manualPick);render();return;}
  if(e.target.closest("#ls-material-select-all, #ls-material-select-clear")){
   document.dispatchEvent(new CustomEvent("ls-selection-command",{detail:{owner:"material",command:e.target.closest("#ls-material-select-all")?"all":"clear"}}));
  }
 });
 document.addEventListener("ls-selection-command",e=>{
  if(pane.hidden||e.detail?.owner!=="material")return;
  const material=materials.find(m=>m.id===activeId);
  if(!material||material.mode!=="manual")return;
  const level=entityLevel(material),command=e.detail.command;
  let ids=explicitIds(material,level);
  if(command==="all")ids=getEntities(level).map(d=>d.id);
  else if(command==="clear")ids=[];
  else if(command==="remove")ids=ids.filter(id=>id!==Number(e.detail.key));
  else return;
  if(level==="Domain")material.domainIds=ids;
  else{material.selectedIdsByLevel??={Boundary:[],Point:[]};material.selectedIdsByLevel[level]=ids;}
  warning="Updated "+level.toLowerCase()+" selection.";render();pickChanged(manualPick);
 });
 pane.addEventListener("toggle",e=>{
  if(e.target.matches?.(".ls-material-refractive-node")){
   refractiveDetailsExpanded=e.target.open;
   if(!e.target.open)return;
   const material=materials.find(m=>m.id===activeId);
   if(material&&!material.properties?.includes("refractive-index"))addRefractiveIndex();
  }
 },true);
 pane.addEventListener("focusin",e=>{
  if(e.target.id==="ls-material-real")activateRefractiveRow("real");
  if(e.target.id==="ls-material-imag")activateRefractiveRow("imag");
 });
 pane.addEventListener("input",e=>{
  const material=materials.find(x=>x.id===activeId);if(!material)return;
  if(e.target.id==="ls-material-label"){material.label=e.target.value;treeRender();}
  if(e.target.id==="ls-material-real"){material.real=e.target.value;activateRefractiveRow("real");updateMaterialEvaluation(material);}
  if(e.target.id==="ls-material-imag"){material.imag=e.target.value;activateRefractiveRow("imag");updateMaterialEvaluation(material);}
 });
 pane.addEventListener("change",e=>{
  const material=materials.find(x=>x.id===activeId);if(!material)return;
  if(e.target.id==="ls-material-entity-level"){
   const newLevel=e.target.value;
   if(!LEVELS.includes(newLevel))return;
   material.level=newLevel;
   // Keep separately selected IDs per entity level. Switching levels must
   // never reuse Domain IDs as Boundary or Point assignments.
   render();pickChanged(material.mode==="manual");
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
    // COMSOL-style Manual selection starts empty. User clicks individual
    // Geometry domains to add them. Previous manual selections are preserved
    // when switching away and back; do not auto-select the entire model.
    render();pickChanged(value==="manual");
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
  // Domain selection is per physical preview region, NOT per whole circle.
  // Boundary and Point retain feature-based selection for now.
  const targetId=Number(e.detail?.domainId);
  if(level==="Domain"&&(!Number.isInteger(targetId)||targetId<1))return;
  const matched=level==="Domain"
   ?entities.filter(entity=>entity.id===targetId)
   :entities.filter(entity=>entity.entityKey===e.detail?.entityKey);
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
  warning="Changed "+level.toLowerCase()+" preview selection for "+(e.detail?.domainId??e.detail?.entityKey)+" ("+matched.length+" entity/ies).";
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
 document.addEventListener("ls-wizard-finished",()=>{materials.length=0;activeId=null;contextMaterialId=null;treeRender();});
 document.addEventListener("ls-adjacent-editor-activated",close);
 document.addEventListener("ls-pml-editor-activated",close);
 document.addEventListener("ls-open-cad",close);
 document.addEventListener("ls-open-wave-optics",close);
 document.addEventListener("ls-open-settings",close);
 window.LSWaveMaterialsEditor=Object.freeze({list:()=>materials.map(m=>({...m,domainIds:[...m.domainIds],selectedIdsByLevel:{...m.selectedIdsByLevel}})),domains:getDomains,entities:getEntities,resolve:resolvedIndex,assignments:previewAssignments});
 treeRender();
})();
