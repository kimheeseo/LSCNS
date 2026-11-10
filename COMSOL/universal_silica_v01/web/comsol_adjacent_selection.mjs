/**
 * COMSOL-style Definitions > Selections > Adjacent editor.
 *
 * This implementation allows persistent named selections of real SVG edges,
 * geometry domains, or center points. It does NOT claim to perform COMSOL's
 * FEM topological adjacency/mesh boundary inference: Output Entities is a
 * preview and selected edges are input boundaries.
 */
(function initializeAdjacentSelections(){
 "use strict";
 const $=id=>document.getElementById(id);
 const tree=document.querySelector(".ls-model-tree");
 const sidebar=document.querySelector(".layout > .sidebar");
 if(!tree||!sidebar||$("ls-adjacent-editor"))return;
 const escapeHTML=value=>String(value??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
 const rows=[];let activeId=null,enabled=false;let notice="";
 const allFor=level=>{
  try{return window.LSWaveCADService?.adjacentEntities?.(level)||[];}
  catch(error){notice="Geometry unavailable: "+error.message;return [];}
 };
 const current=()=>rows.find(row=>row.id===activeId);
 const selectedKeys=(row=current())=>new Set(row?.selectedByLevel?.[row.level]||[]);
 const selectedEntities=(row=current())=>{
  if(!row)return [];
  const keys=selectedKeys(row);
  return allFor(row.level).filter(e=>keys.has(e.key));
 };
 const pane=document.createElement("section");
 pane.id="ls-adjacent-editor";pane.className="ls-config-pane ls-adjacent-editor";
 pane.hidden=true;sidebar.append(pane);
 function renderTree(){
  const group=$("ls-definitions-selections-group");
  const list=$("ls-definitions-selections-list");
  if(!group||!list)return;
  group.hidden=!rows.length;
  list.replaceChildren();
  for(const row of rows){
   const button=document.createElement("button");
   button.type="button";button.className="ls-tree-leaf"+(row.id===activeId&&!pane.hidden?" active":"");
   button.dataset.adjacentNode=row.id;
   button.textContent="◫ "+row.label;
   button.title="Adjacent preview boundary selection: "+row.label;
   list.append(button);
  }
 }
 function transmit(){
  const row=current();
  document.dispatchEvent(new CustomEvent("ls-adjacent-picker",{detail:{
   visible:!pane.hidden,
   enabled:enabled&&!!row,
   level:row?.level||"Boundary",
   selectedKeys:[...selectedKeys(row)]
  }}));
 }
 function render(){
  const row=current();if(!row)return;
  const selected=selectedEntities(row);
  const candidates=allFor(row.level);
  // Resolve stored keys only against current geometry; keep the stored names
  // when the model changes, but remove stale visual edge IDs from selection.
  const keyset=new Set(candidates.map(v=>v.key));
  row.selectedByLevel[row.level]=(row.selectedByLevel[row.level]||[]).filter(key=>keyset.has(key));
  const entries=selectedEntities(row);
  const selectedMarkup=entries.length
   ?entries.map(ent=>'<label class="ls-adjacent-row"><input type="checkbox" data-adjacent-list-key="'+escapeHTML(ent.key)+'" checked/><span><b>'+ent.id+'</b> '+escapeHTML(ent.label)+'</span></label>').join("")
   :'<p class="ls-adjacent-empty">No input entities selected. Click individual '+row.level.toLowerCase()+(row.level==="Boundary"?" rings or lines":" regions")+' in Graphics.</p>';
  pane.innerHTML=
   '<div class="ls-config-heading"><div><strong>Settings</strong><small>Adjacent</small></div><button type="button" id="ls-adjacent-close">← Close</button></div>'+
   '<div class="ls-adjacent-body">'+
   '<div class="ls-setting-line"><label for="ls-adjacent-label">Label:</label><input id="ls-adjacent-label" value="'+escapeHTML(row.label)+'" autocomplete="off"/></div>'+
   '<details class="ls-adjacent-section" open><summary>Input Entities</summary><div class="ls-adjacent-section-content">'+
   '<label class="ls-adjacent-field" for="ls-adjacent-level">Geometric entity level:</label>'+
   '<select id="ls-adjacent-level"><option value="Domain" '+(row.level==="Domain"?"selected":"")+'>Domain</option>'+
   '<option value="Boundary" '+(row.level==="Boundary"?"selected":"")+'>Boundary</option>'+
   '<option value="Point" '+(row.level==="Point"?"selected":"")+'>Point</option></select>'+
   '<label class="ls-adjacent-field">Input selections:</label>'+
   '<div id="ls-adjacent-list" aria-label="Selected input entities">'+selectedMarkup+'</div>'+
   '<div class="ls-adjacent-actions"><button type="button" id="ls-adjacent-activate" aria-pressed="'+enabled+'">'+(enabled?"<span>▣</span> Activate Selection · On":"<span>□</span> Activate Selection")+'</button>'+
   '<button type="button" id="ls-adjacent-clear">Clear</button><span>'+entries.length+' of '+candidates.length+' selected</span></div>'+
   '</div></details>'+
   '<details class="ls-adjacent-section" open><summary>Output Entities</summary><div class="ls-adjacent-section-content">'+
   '<label class="ls-adjacent-field" for="ls-adjacent-output">Geometric entity level:</label>'+
   '<select id="ls-adjacent-output">'+["Adjacent domains","Adjacent boundaries","Adjacent points"].map(label=>
   '<option value="'+label+'" '+(row.output===label?"selected":"")+'>'+label+'</option>').join("")+'</select>'+
   '<label class="ls-adjacent-check"><input type="checkbox" id="ls-adjacent-exterior" '+(row.exterior?"checked":"")+'/> Exterior boundaries</label>'+
   '<label class="ls-adjacent-check"><input type="checkbox" id="ls-adjacent-interior" '+(row.interior?"checked":"")+'/> Interior boundaries</label>'+
   '<p class="ls-adjacent-note">This is a Geometry labeling/selection preview. Blue outlines are the selected input entities; COMSOL-style topological adjacency output is not yet calculated.</p>'+
   '</div></details>'+
   '<p id="ls-adjacent-status" role="status">'+escapeHTML(notice||"Click a "+row.level.toLowerCase()+" in Graphics to add or remove it from "+row.label+".")+'</p>'+
   '</div>';
  renderTree();
 }
 function open(row,create=false){
  if(!row)return;
  // All other editors clean up their own state. This tool then claims the
  // sidebar and graphics so the order of event listeners cannot hide it.
  document.dispatchEvent(new CustomEvent("ls-adjacent-editor-activated",{detail:{id:row.id}}));
  activeId=row.id;enabled=true;
  $("ls-materials-editor")&&( $("ls-materials-editor").hidden=true);
  $("ls-cad-pane")&&($("ls-cad-pane").hidden=true);
  $("ls-wave-settings")&&($("ls-wave-settings").hidden=true);
  $("ls-config-pane")&&($("ls-config-pane").hidden=true);
  pane.hidden=false;
  sidebar.classList.add("ls-config-active");
  document.body.classList.remove("ls-cad-active","ls-material-active","ls-wave-optics-active");
  document.body.classList.add("ls-configuration-open","ls-adjacent-active");
  const graphics=$("ls-cad-graphics");if(graphics)graphics.hidden=false;
  const group=$("ls-component-definitions");if(group)group.open=true;
  const selections=$("ls-definitions-selections-group");if(selections)selections.open=true;
  notice=create?"Created "+row.label+" under Component 1 → Definitions → Selections.":"Editing "+row.label;
  render();transmit();
 }
 function close(){
  if(pane.hidden)return;
  enabled=false;pane.hidden=true;transmit();
  activeId=null;
  document.body.classList.remove("ls-adjacent-active","ls-configuration-open");
  sidebar.classList.remove("ls-config-active");
  renderTree();
 }
 function newAdjacent(){
  let i=1;while(rows.some(row=>row.id==="adj"+i))i++;
  const row={id:"adj"+i,label:"Adjacent "+i,level:"Domain",output:"Adjacent boundaries",
   selectedByLevel:{Domain:[],Boundary:[],Point:[]},
   exterior:true,interior:false};
  rows.push(row);
  open(row,true);
 }
 tree.addEventListener("click",e=>{
  const button=e.target.closest("[data-adjacent-node]");
  if(!button)return;
  e.stopImmediatePropagation();e.preventDefault();
  const row=rows.find(x=>x.id===button.dataset.adjacentNode);
  if(row)open(row);
 },true);
 pane.addEventListener("input",e=>{
  const row=current();if(!row)return;
  if(e.target.id==="ls-adjacent-label"){
   row.label=e.target.value;
   renderTree();
   notice="Name is saved in Model Builder and will be kept while this model is open.";
  }
 });
 pane.addEventListener("change",e=>{
  const row=current();if(!row)return;
  if(e.target.id==="ls-adjacent-level"){
   row.level=e.target.value;
   notice="Choose "+row.level.toLowerCase()+" entities directly in Graphics.";
   render();transmit();return;
  }
  if(e.target.id==="ls-adjacent-output"){row.output=e.target.value;notice="Output entity level: "+row.output+" (preview; topology computation not implemented).";render();transmit();return;}
  if(e.target.id==="ls-adjacent-exterior"){row.exterior=e.target.checked;return;}
  if(e.target.id==="ls-adjacent-interior"){row.interior=e.target.checked;return;}
  if(e.target.matches("[data-adjacent-list-key]")){
   if(!e.target.checked)row.selectedByLevel[row.level]=row.selectedByLevel[row.level].filter(key=>key!==e.target.dataset.adjacentListKey);
   render();transmit();
  }
 });
 pane.addEventListener("click",e=>{
  if(e.target.closest("#ls-adjacent-close")){close();return;}
  const row=current();if(!row)return;
  if(e.target.closest("#ls-adjacent-activate")){enabled=!enabled;notice=enabled?"Click Graphics "+row.level.toLowerCase()+" entities to toggle inputs.":"Graphics selection paused.";render();transmit();}
  if(e.target.closest("#ls-adjacent-clear")){row.selectedByLevel[row.level]=[];notice="Cleared selected "+row.level.toLowerCase()+" inputs.";render();transmit();}
 });
 document.addEventListener("ls-open-adjacent-editor",e=>{
  if(e.detail?.create)newAdjacent();
  else if(e.detail?.id){const row=rows.find(r=>r.id===e.detail.id);if(row)open(row);}
 });
 document.addEventListener("ls-adjacent-geometry-click",e=>{
  const row=current();
  if(!enabled||!row||e.detail?.level!==row.level)return;
  const item=allFor(row.level).find(ent=>ent.key===e.detail?.key);
  if(!item)return;
  const keys=new Set(row.selectedByLevel[row.level]||[]);
  if(keys.has(item.key))keys.delete(item.key);else keys.add(item.key);
  row.selectedByLevel[row.level]=[...keys];
  notice="Selected "+row.level.toLowerCase()+": "+[...keys].length+" item(s) · "+row.label;
  render();transmit();
 });
 for(const eventName of ["ls-open-cad","ls-open-material-editor","ls-open-wave-optics","ls-open-settings"]){
  document.addEventListener(eventName,()=>{if(!pane.hidden)close();},true);
 }
 document.addEventListener("ls-wizard-finished",()=>{
  if(!pane.hidden)close();
  rows.length=0;renderTree();
 });
 document.addEventListener("ls-cad-state",()=>{
  if(pane.hidden)return;
  // Renders only if a selection becomes invalid, not on every CAD redraw,
  // preventing a recursive render → Graphics → ls-cad-state loop.
  const row=current(),keys=selectedKeys(row),valid=new Set(allFor(row.level).map(e=>e.key));
  if([...keys].some(key=>!valid.has(key))){render();transmit();}
 });
 window.LSWaveAdjacentSelection=Object.freeze({
  list:()=>rows.map(row=>({...row,selectedByLevel:{
   Domain:[...row.selectedByLevel.Domain],
   Boundary:[...row.selectedByLevel.Boundary],
   Point:[...row.selectedByLevel.Point]
  }})),
  selected:()=>selectedEntities().map(ent=>({...ent}))
 });
 renderTree();
})();
