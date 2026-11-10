/**
 * Definitions > Artificial Domains > Perfectly Matched Layer.
 * COMSOL-inspired editor for the SVG Geometry preview.
 *
 * IMPORTANT: These settings and domain selections are saved in the current
 * browser model session, but they do NOT alter the radial FEM matrix or
 * produce a vector FEM PML absorption/confinement-loss calculation.
 */
(function initializePMLEditor(){
 "use strict";
 const $=id=>document.getElementById(id);
 const tree=document.querySelector(".ls-model-tree");
 const sidebar=document.querySelector(".layout > .sidebar");
 if(!tree||!sidebar||$("ls-pml-editor"))return;
 const safe=s=>String(s??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
 const rows=[];let activeId=null,picking=false,notice="";
 const geometry=()=>window.LSWaveCADService;
 const domains=()=>{try{return geometry()?.domains?.()||[];}catch(e){notice="Geometry preview error: "+e.message;return [];}};
 const current=()=>rows.find(r=>r.id===activeId);
 const selectedDomains=(row=current())=>{
  if(!row)return [];
  const all=domains();
  if(row.selection==="all")return all;
  const keys=new Set(row.domainKeys);
  return all.filter(d=>keys.has(d.key));
 };
 const pane=document.createElement("section");
 pane.id="ls-pml-editor";pane.className="ls-config-pane ls-pml-editor";
 pane.hidden=true;sidebar.append(pane);
 function treeRender(){
  const group=$("ls-definitions-artificial-group"),list=$("ls-definitions-artificial-list");
  if(!group||!list)return;
  group.hidden=rows.length===0;list.replaceChildren();
  for(const row of rows){
   const b=document.createElement("button");b.type="button";
   b.className="ls-tree-leaf ls-pml-tree-item"+(row.id===activeId&&!pane.hidden?" active":"");
   b.dataset.pmlNode=row.id;
   b.textContent="▧ "+row.label+" ("+row.id+")";
   b.title="Perfectly Matched Layer (Geometry selection preview; solver not linked)";
   list.append(b);
  }
 }
 function broadcast(){
  const row=current(),currentDomains=selectedDomains(row);
  document.dispatchEvent(new CustomEvent("ls-pml-picker",{detail:{
   visible:!pane.hidden&&!!row,enabled:!pane.hidden&&picking&&row?.selection==="manual",
   selectedDomainIds:currentDomains.map(d=>d.id),
   selectedDomainKeys:currentDomains.map(d=>d.key)
  }}));
 }
 function render(){
  const row=current();if(!row)return;
  const all=domains(),selected=new Set(selectedDomains(row).map(d=>d.key));
  const listHtml=all.length?all.map(d=>
   '<label class="ls-pml-domain-row"><input type="checkbox" data-pml-domain-key="'+safe(d.key)+'" '+
   (selected.has(d.key)?"checked":"")+' '+(row.selection==="manual"?"":"disabled")+'/>'+
   '<span><b>'+d.id+'</b> '+safe(d.label)+'</span></label>').join(""):
   '<p class="ls-pml-note">No Geometry domains available. Build Geometry 1 first.</p>';
  pane.innerHTML=
   '<div class="ls-config-heading"><div><strong>Settings</strong><small>Perfectly Matched Layer</small></div><button type="button" id="ls-pml-close">← Close</button></div>'+
   '<div class="ls-pml-body">'+
   '<div class="ls-setting-line"><label for="ls-pml-label">Label:</label><input id="ls-pml-label" value="'+safe(row.label)+'"/></div>'+
   '<div class="ls-setting-line"><label for="ls-pml-name">Name:</label><input id="ls-pml-name" value="'+safe(row.id)+'" readonly title="PML feature tag"/></div>'+
   '<details class="ls-pml-section" open><summary>Domain Selection</summary><div class="ls-pml-section-inner">'+
   '<label for="ls-pml-selection">Selection:</label>'+
   '<select id="ls-pml-selection"><option value="manual" '+(row.selection==="manual"?"selected":"")+'>Manual</option>'+
   '<option value="all" '+(row.selection==="all"?"selected":"")+'>All domains</option></select>'+
   '<div class="ls-pml-domain-select"><button type="button" id="ls-pml-activate" aria-pressed="'+picking+'" '+
   (row.selection==="manual"?"":"disabled")+' title="Activate domain selection in Graphics">▣</button>'+
   '<div id="ls-pml-domain-list" aria-label="PML domain selection">'+listHtml+'</div></div>'+
   '<div class="ls-pml-actions"><button type="button" id="ls-pml-all" '+(row.selection==="manual"?"":"disabled")+'>All</button><button type="button" id="ls-pml-clear" '+(row.selection==="manual"?"":"disabled")+
   '>Clear</button><span id="ls-pml-count">'+selected.size+' of '+all.length+' Geometry domains selected</span></div>'+
   '</div></details>'+
   '<details class="ls-pml-section"><summary>Override</summary><p class="ls-pml-note">No numerical solver override is connected.</p></details>'+
   '<details class="ls-pml-section" open><summary>Geometry</summary><div class="ls-pml-section-inner">'+
   '<label for="ls-pml-type">Type:</label><select id="ls-pml-type">'+
    ["Cartesian","Cylindrical","Spherical"].map(v=>'<option '+(row.geometryType===v?"selected":"")+'>'+v+'</option>').join("")+'</select>'+
   '<p class="ls-pml-note">Geometry type is saved for the PML definition. Coordinate transforms are not yet applied to the FEM mesh.</p>'+
   '</div></details>'+
   '<details class="ls-pml-section" open><summary>Scaling</summary><div class="ls-pml-section-inner">'+
   '<label for="ls-pml-stretch">Coordinate stretching type:</label>'+
   '<select id="ls-pml-stretch">'+["Polynomial","Rational"].map(v=>'<option '+(row.stretch===v?"selected":"")+'>'+v+'</option>').join("")+'</select>'+
   '<label for="ls-pml-wavelength-source">Typical wavelength from:</label>'+
   '<select id="ls-pml-wavelength-source">'+["Physics interface","User defined"].map(v=>'<option '+(row.wavelengthSource===v?"selected":"")+'>'+v+'</option>').join("")+'</select>'+
   (row.wavelengthSource==="User defined"?'<label for="ls-pml-wavelength">Typical wavelength:</label>'+
    '<input id="ls-pml-wavelength" value="'+safe(row.wavelength)+'" placeholder="e.g. wl"/>':'')+
   '</div></details>'+
   '<p class="ls-pml-disclaimer">PML definition and colored domains are UI/Geometry preview only. They are not applied to the existing radial FEM or full-vector EWFD solver.</p>'+
   '<p id="ls-pml-status" role="status">'+safe(notice||"Click a domain in Graphics to assign or unassign it from "+row.label+".")+'</p>'+
   '</div>';
  treeRender();
 }
 function open(row,newlyCreated=false){
  if(!row)return;
  document.dispatchEvent(new CustomEvent("ls-pml-editor-activated",{detail:{id:row.id}}));
  activeId=row.id;picking=row.selection==="manual";
  for(const id of ["ls-materials-editor","ls-adjacent-editor","ls-cad-pane","ls-wave-settings","ls-config-pane"]){
   const el=$(id);if(el)el.hidden=true;
  }
  pane.hidden=false;sidebar.classList.add("ls-config-active");
  document.body.classList.remove("ls-cad-active","ls-material-active","ls-wave-optics-active","ls-adjacent-active");
  document.body.classList.add("ls-configuration-open","ls-pml-active");
  const graphics=$("ls-cad-graphics");if(graphics)graphics.hidden=false;
  const root=$("ls-component-definitions"),group=$("ls-definitions-artificial-group");
  if(root)root.open=true;if(group)group.open=true;
  notice=newlyCreated?"Created "+row.label+" under Definitions → Artificial Domains.":"Editing "+row.label+".";
  render();broadcast();
 }
 function close(){
  if(pane.hidden)return;
  picking=false;pane.hidden=true;broadcast();activeId=null;
  document.body.classList.remove("ls-pml-active");
  const otherMode=["ls-material-active","ls-cad-active","ls-wave-optics-active","ls-adjacent-active"]
   .some(cls=>document.body.classList.contains(cls));
  if(!otherMode){
   document.body.classList.remove("ls-configuration-open");
   sidebar.classList.remove("ls-config-active");
  }
  treeRender();
 }
 function add(){
  let n=1;while(rows.some(row=>row.id==="pml"+n))n++;
  const row={id:"pml"+n,label:"Perfectly Matched Layer "+n,selection:"manual",domainKeys:[],
   geometryType:"Cartesian",stretch:"Polynomial",wavelengthSource:"Physics interface",wavelength:"wl"};
  rows.push(row);open(row,true);
 }
 tree.addEventListener("click",e=>{
  const item=e.target.closest("[data-pml-node]");
  if(!item)return;
  e.preventDefault();e.stopImmediatePropagation();
  const row=rows.find(r=>r.id===item.dataset.pmlNode);
  if(row)open(row);
 },true);
 pane.addEventListener("input",e=>{
  const row=current();if(!row)return;
  if(e.target.id==="ls-pml-label"){
   row.label=e.target.value||"Perfectly Matched Layer";
   treeRender();notice="PML label updated in Model Builder.";
  }
  if(e.target.id==="ls-pml-wavelength")row.wavelength=e.target.value;
 });
 pane.addEventListener("change",e=>{
  const row=current();if(!row)return;
  if(e.target.id==="ls-pml-selection"){
   row.selection=e.target.value;picking=row.selection==="manual";
   render();broadcast();return;
  }
  if(e.target.id==="ls-pml-type"){row.geometryType=e.target.value;return;}
  if(e.target.id==="ls-pml-stretch"){row.stretch=e.target.value;return;}
  if(e.target.id==="ls-pml-wavelength-source"){
   row.wavelengthSource=e.target.value;render();broadcast();return;
  }
  const checkbox=e.target.closest("[data-pml-domain-key]");
  if(checkbox&&row.selection==="manual"){
   const keys=new Set(row.domainKeys),key=checkbox.dataset.pmlDomainKey;
   if(checkbox.checked)keys.add(key);else keys.delete(key);
   row.domainKeys=[...keys];notice="Selected "+keys.size+" PML input domain(s).";
   render();broadcast();
  }
 });
 pane.addEventListener("click",e=>{
  if(e.target.closest("#ls-pml-close")){close();return;}
  const row=current();if(!row)return;
  if(e.target.closest("#ls-pml-activate")){
   picking=!picking;
   notice=picking?"Click Graphics domains to select or deselect.":"Graphics selection paused.";
   render();broadcast();
  }
  if(e.target.closest("#ls-pml-clear")&&row.selection==="manual"){
   row.domainKeys=[];notice="Cleared PML domain selection.";render();broadcast();
  }
  if(e.target.closest("#ls-pml-all")&&row.selection==="manual"){
   row.domainKeys=domains().map(d=>d.key);notice="Selected all PML input domains.";render();broadcast();
  }
 });
 document.addEventListener("ls-selection-command",e=>{
  const row=current();
  if(pane.hidden||!row||row.selection!=="manual"||e.detail?.owner!=="pml")return;
  const command=e.detail.command;
  if(command==="all")row.domainKeys=domains().map(d=>d.key);
  else if(command==="clear")row.domainKeys=[];
  else if(command==="remove")row.domainKeys=row.domainKeys.filter(key=>key!==e.detail.key);
  else return;
  notice="Updated PML domain selection.";render();broadcast();
 });
 document.addEventListener("ls-open-pml-editor",e=>{
  if(e.detail?.create)add();
  else if(e.detail?.id){const row=rows.find(r=>r.id===e.detail.id);if(row)open(row);}
 });
 document.addEventListener("ls-pml-geometry-click",e=>{
  const row=current();if(pane.hidden||!picking||row?.selection!=="manual")return;
  const domain=domains().find(d=>d.key===e.detail?.key);
  if(!domain)return;
  const keys=new Set(row.domainKeys);
  if(keys.has(domain.key))keys.delete(domain.key);else keys.add(domain.key);
  row.domainKeys=[...keys];
  notice="Selected "+keys.size+" PML input domain(s).";
  render();broadcast();
 });
 for(const eventName of ["ls-open-cad","ls-open-material-editor","ls-open-adjacent-editor",
    "ls-open-wave-optics","ls-open-settings"]){
  document.addEventListener(eventName,()=>{if(!pane.hidden)close();},true);
 }
 document.addEventListener("ls-wizard-finished",()=>{if(!pane.hidden)close();rows.length=0;treeRender();});
 document.addEventListener("ls-cad-state",()=>{
  if(pane.hidden)return;
  const row=current();if(!row)return;
  const currentKeys=new Set(domains().map(d=>d.key));
  if(row.domainKeys.some(key=>!currentKeys.has(key))){
   row.domainKeys=row.domainKeys.filter(key=>currentKeys.has(key));
   notice="Selection updated to reflect changed Geometry.";
   render();broadcast();
  }
 });
 window.LSWavePMLService=Object.freeze({
  list:()=>rows.map(row=>({...row,domainKeys:[...row.domainKeys]})),
  domains:()=>domains().map(d=>({...d})),
  selected:()=>selectedDomains().map(d=>({...d}))
 });
 treeRender();
})();
