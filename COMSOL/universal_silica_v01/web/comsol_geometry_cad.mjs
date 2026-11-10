(function geometryCad(){
 "use strict";
 const $=id=>document.getElementById(id),side=document.querySelector(".layout > .sidebar"),main=document.querySelector("main.workspace"),svc=window.LSWaveParameterService;
 if(!side||!main||!svc||$("ls-cad-pane"))return;
 const esc=s=>String(s??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
 const unit=n=>Number(n.toPrecision(10)).toString();
 const circle=(id,label,radius)=>({id,kind:"circle",label,enabled:true,type:"Solid",radius,sector:"360",base:"Center",x:"0[um]",y:"0[um]",rotation:"0",layer:"0[um]",layerExtra:[],selection:false,color:"None"});
 const defaults=()=>[circle("c1","Circle 1","rb"),circle("c2","Circle 2","d/2")];
 let nodes=defaults(),selected="c2",limit=null,zoom=1,gridOn=true,clipboard=null,rootLabel="Geometry 1",lengthUnit="um",panX=0,panY=0,viewExtent=null;
 let angleUnit="Degrees",constraintMode="Off",repairTolerance="Automatic",buildOnNew=true,buildOnLeave=true,scaleWhenUnitsChange=false,selectingRotateInputs=false,activeCumulativeId=null,materialPickMode=false;
 const undoStack=[],redoStack=[],cumulativeSelections=[];
 let materialHighlightKeys=new Set(),materialHighlightDomainIds=new Set(),materialEntityLevel="Domain";
 let adjacentView=false,adjacentGridOn=false,adjacentPickMode=false,adjacentLevel="Boundary",adjacentSelectedKeys=new Set();
 function snap(){return {nodes:structuredClone(nodes),selected,limit,zoom,panX,panY,rootLabel,lengthUnit,gridOn,viewExtent};}
 function updateHistoryButtons(){
  if($("ls-cad-undo"))$("ls-cad-undo").disabled=!undoStack.length;
  if($("ls-cad-redo"))$("ls-cad-redo").disabled=!redoStack.length;
  document.dispatchEvent(new CustomEvent("ls-cad-history",{detail:{canUndo:!!undoStack.length,canRedo:!!redoStack.length}}));
 }
 function checkpoint(){undoStack.push(snap());if(undoStack.length>60)undoStack.shift();redoStack.length=0;updateHistoryButtons();}
 function restore(state){({nodes,selected,limit,zoom,panX,panY,rootLabel,lengthUnit,gridOn,viewExtent}=structuredClone(state));const node=$("ls-cad-root-name");if(node)node.textContent=rootLabel;loadFields();updateHistoryButtons();}
 function historyMove(back){const from=back?undoStack:redoStack,to=back?redoStack:undoStack;if(!from.length)return;to.push(snap());restore(from.pop());say(back?"Undo geometry edit.":"Redo geometry edit.");}
 const pane=document.createElement("section");pane.id="ls-cad-pane";pane.className="ls-config-pane ls-cad-pane";pane.hidden=true;
 pane.innerHTML=`
 <div class="ls-config-heading"><div><strong>Settings</strong><small id="ls-cad-subtitle">Circle</small></div><button type="button" id="ls-cad-close">← Fiber settings</button></div>
 <div class="ls-cad-tools"><button id="ls-cad-build-selected" type="button">▣ Build Selected</button><button id="ls-cad-build-all" type="button">▦ Build All Objects</button><button id="ls-cad-undo" type="button" disabled title="Ctrl+Z">↶ Undo</button><button id="ls-cad-redo" type="button" disabled title="Ctrl+Shift+Z">↷ Redo</button></div>
 <div class="ls-config-body">
  <div id="ls-cad-root-settings" hidden>
   <div class="ls-setting-line"><label for="ls-cad-root-label">Label</label><input id="ls-cad-root-label" value="Geometry 1"/></div>
   <details class="ls-cad-section" open><summary>Units</summary>
    <label class="ls-cad-check"><input id="ls-cad-scale-unit" type="checkbox"/> Scale values when changing units</label>
    <div class="ls-cad-field"><label for="ls-cad-root-unit">Length unit</label><select id="ls-cad-root-unit"><option value="m">m</option><option value="mm">mm</option><option value="um">µm</option><option value="nm">nm</option></select><span></span></div>
    <div class="ls-cad-field"><label for="ls-cad-angle-unit">Angular unit</label><select id="ls-cad-angle-unit"><option>Degrees</option><option>Radians</option></select><span></span></div>
    </details>
   <details class="ls-cad-section" open><summary>Constraints and Dimensions</summary>
    <div class="ls-cad-field"><label for="ls-cad-constraints">Use constraints and dimensions</label><select id="ls-cad-constraints"><option>Off</option><option>On</option></select><span></span></div>
    <p class="ls-cad-sequence-note">Constraint solver is not yet implemented; the selector is for UI parity.</p>
   </details>
   <details class="ls-cad-section" open><summary>Advanced</summary>
    <div class="ls-cad-field"><label for="ls-cad-repair">Default repair tolerance</label><select id="ls-cad-repair"><option>Automatic</option><option>Manual</option></select><span></span></div>
    <label class="ls-cad-check"><input id="ls-cad-build-new" type="checkbox" checked/> Build new operations automatically</label>
    <label class="ls-cad-check"><input id="ls-cad-build-leave" type="checkbox" checked/> Build automatically when leaving geometry</label>
   </details>
   <details class="ls-cad-section" open><summary>Geometry Sequence</summary><div id="ls-cad-root-sequence" class="ls-cad-sequence"></div><p class="ls-cad-sequence-note">Select Circle/Rotate to edit. Build All regenerates the 2D preview.</p></details>
   <details class="ls-cad-section" open><summary>Finalization</summary><p class="ls-cad-sequence-note">Form Union (UI reference). General-purpose 2D Boolean domains are not computed by the radial FEM.</p></details>
  </div>
  <div id="ls-cad-feature-settings">
  <div class="ls-setting-line"><label for="ls-cad-label">Label</label><input id="ls-cad-label" /></div>
  <div id="ls-cad-circle">
   <details class="ls-cad-section" open><summary>Object Type</summary><div class="ls-cad-field"><label for="ls-cad-type">Type</label><select id="ls-cad-type"><option>Solid</option><option>Curve</option></select><span></span></div></details>
   <details class="ls-cad-section" open><summary>Size and Shape</summary>
    <div class="ls-cad-field"><label for="ls-cad-radius">Radius</label><input id="ls-cad-radius"/><span>m</span></div>
    <div class="ls-cad-field"><label for="ls-cad-sector">Sector angle</label><input id="ls-cad-sector"/><span>deg</span></div></details>
   <details class="ls-cad-section" open><summary>Position</summary>
    <div class="ls-cad-field"><label for="ls-cad-base">Base</label><select id="ls-cad-base"><option>Center</option></select><span></span></div>
    <div class="ls-cad-field"><label for="ls-cad-x">x</label><input id="ls-cad-x"/><span>m</span></div>
    <div class="ls-cad-field"><label for="ls-cad-y">y</label><input id="ls-cad-y"/><span>m</span></div></details>
   <details class="ls-cad-section" open><summary>Rotation Angle</summary><div class="ls-cad-field"><label for="ls-cad-rotation">Rotation</label><input id="ls-cad-rotation"/><span>deg</span></div></details>
   <details class="ls-cad-section" id="ls-cad-layers-section"><summary>Layers</summary>
     <div class="ls-cad-layer"><span>Layer name</span><span>Thickness (m)</span><label for="ls-cad-layer">Layer 1</label><input id="ls-cad-layer" spellcheck="false" aria-label="Layer 1 thickness"/><div id="ls-cad-extra-layers"></div></div>
     <div class="ls-cad-layer-tools"><button id="ls-cad-add-layer" type="button">＋ Add Layer</button><button id="ls-cad-remove-layer" type="button">－ Remove Last Layer</button></div>
     <p class="ls-cad-sequence-note">Layer thickness is drawn inward from the circle boundary. Multiple layers create concentric boundaries in the Graphics preview; not linked to the radial FEM.</p>
   </details>
  </div>
  <div id="ls-cad-rotate" hidden>
   <details class="ls-cad-section" open><summary>Input</summary><label class="ls-cad-block-label">Input objects:</label><div class="ls-cad-selection-row"><button type="button" id="ls-cad-activate-selection" aria-label="Activate Selection" aria-pressed="false" title="Activate Selection in Graphics">▣</button><div id="ls-cad-inputs" class="ls-cad-inputs"></div></div><p id="ls-cad-selection-hint" class="ls-cad-selection-hint">Activate Selection to highlight/select c2 in Graphics.</p><label class="ls-cad-check"><input id="ls-cad-keep" type="checkbox" checked/> Keep input objects</label></details>
   <details class="ls-cad-section" open id="ls-cad-rotate-angle-section"><summary>Rotation</summary><div class="ls-cad-field"><label for="ls-cad-angle">Angle</label><div class="ls-cad-angle-entry"><input id="ls-cad-angle" spellcheck="false"/><span>deg</span><button type="button" id="ls-cad-range-open" title="Range — Start, Step, Stop" aria-label="Range">▥</button></div><span></span></div><small>Use an angle, range(start,step,stop), or comma-separated ranges. Click Range to open the step dialog.</small></details>
   <details class="ls-cad-section" open id="ls-cad-rotate-center-section"><summary>Center of Rotation</summary><div class="ls-cad-field"><label for="ls-cad-cx">x</label><input id="ls-cad-cx"/><span>m</span></div><div class="ls-cad-field"><label for="ls-cad-cy">y</label><input id="ls-cad-cy"/><span>m</span></div></details>
   <details class="ls-cad-section" id="ls-cad-mirror-point-section" hidden open><summary>Point on Line of Reflection</summary><div class="ls-cad-field"><label for="ls-cad-mirror-cx">x</label><input id="ls-cad-mirror-cx" value="0[um]"/><span>m</span></div><div class="ls-cad-field"><label for="ls-cad-mirror-cy">y</label><input id="ls-cad-mirror-cy" value="0[um]"/><span>m</span></div></details>
   <details class="ls-cad-section" id="ls-cad-mirror-normal-section" hidden open><summary>Normal Vector to Line of Reflection</summary><div class="ls-cad-field"><label for="ls-cad-mirror-nx">x</label><input id="ls-cad-mirror-nx" value="1"/><span></span></div><div class="ls-cad-field"><label for="ls-cad-mirror-ny">y</label><input id="ls-cad-mirror-ny" value="0"/><span></span></div><p class="ls-cad-sequence-note">Mirror reflects selected circles across this line. Geometry preview only.</p></details>
  </div>
  <details class="ls-cad-section" open><summary>Selections of Resulting Entities</summary>
   <label class="ls-cad-check"><input id="ls-cad-selection" type="checkbox"/> Resulting objects selection</label>
   <div class="ls-cad-field"><label for="ls-cad-physics">Show in physics</label><select id="ls-cad-physics" disabled><option>Domain selection</option><option>Boundary selection</option></select><span></span></div>
   <div class="ls-cad-field"><label for="ls-cad-color">Color</label><select id="ls-cad-color"><option>None</option><option>Blue</option><option>Orange</option><option>Green</option></select><span></span></div>
   <div class="ls-cad-subsection">Cumulative selection</div><div class="ls-cad-field ls-cad-cumulative-row"><label for="ls-cad-contribute">Contribute to</label><select id="ls-cad-contribute"><option value="">None</option></select><button id="ls-cad-cumulative-new" type="button" title="New cumulative selection">New</button></div></details>
  <div class="ls-cad-mini"><svg id="ls-cad-mini-svg" viewBox="0 0 460 230" role="img" aria-label="Small geometry preview"></svg></div>
  </div>
  <div id="ls-cad-cumulative-settings" hidden>
    <div class="ls-setting-line"><label for="ls-cad-cumulative-label">Label</label><input id="ls-cad-cumulative-label" readonly /></div>
    <details class="ls-cad-section" open><summary>Cumulative Selection</summary><div class="ls-cad-block-label" id="ls-cad-cumulative-summary"></div><p class="ls-cad-sequence-note">This named selection references Geometry results for display and organization. It does not create a full-vector FEM material domain.</p></details>
  </div>
  <div class="ls-config-actions"><button type="button" id="ls-cad-add-circle">＋ Circle</button><button type="button" id="ls-cad-add-rotate">↻ Rotate</button><button type="button" id="ls-cad-add-mirror">⇋ Mirror</button><button type="button" id="ls-cad-reset">Reset</button><button type="button" id="ls-cad-export">Export JSON</button></div>
  <label class="ls-cad-check"><input id="ls-cad-fix-nbg" type="checkbox"/> Use nbg as fixed cladding index</label>
  <button class="ls-config-primary" type="button" id="ls-cad-apply">Apply supported geometry to Radial FEM</button>
  <p class="ls-config-help">Circle and Rotate are interactive 2D geometry features. The 1D Radial FEM only accepts the original two concentric 360° Solid circles without extra layers, added circles or Rotate features. Arbitrary layouts are preview-only.</p>
 </div><p class="ls-config-status" id="ls-cad-status" role="status"></p>`;
 side.append(pane);
 const graphics=document.createElement("section");
 graphics.id="ls-cad-graphics";graphics.hidden=true;
 graphics.innerHTML=`<div class="ls-cad-graphics-bar"><b>Graphics · Geometry 1</b><span id="ls-cad-build-label">Build All</span><div><button id="ls-cad-plus" type="button" title="Zoom in">＋</button><button id="ls-cad-minus" type="button" title="Zoom out">－</button><button id="ls-cad-fit" type="button" title="Fit drawing">⌖</button><button id="ls-cad-grid" type="button" title="Toggle grid">▦</button></div></div><div class="ls-cad-graphics-select"><label for="ls-cad-object-select">Graphics selection</label><select id="ls-cad-object-select" aria-label="Select Geometry feature in Graphics"></select><span>빈 영역 드래그: 이동 · 휠: 확대/축소</span></div><svg id="ls-cad-view" viewBox="0 0 640 500" aria-label="Geometry 2D graphics"></svg><p id="ls-cad-legend">2D Geometry preview · not a FEM field plot</p>`;
 main.prepend(graphics);
 const context=document.createElement("div");
 context.id="ls-cad-menu";context.className="ls-cad-menu";context.hidden=true;context.setAttribute("role","menu");
 document.body.append(context);
 const confirmDialog=document.createElement("dialog");
 confirmDialog.id="ls-cad-delete-dialog";
 confirmDialog.setAttribute("aria-labelledby","ls-cad-delete-title");
 confirmDialog.innerHTML='<form method="dialog"><h3 id="ls-cad-delete-title">Confirm Delete</h3><p id="ls-cad-delete-message"></p><div class="ls-cad-dialog-actions"><button type="submit" value="yes" class="ls-cad-confirm-yes">Yes</button><button type="submit" value="no" class="ls-cad-confirm-no">No</button></div></form>';
 document.body.append(confirmDialog);
 const cumulativeDialog=document.createElement("dialog");
 cumulativeDialog.id="ls-cad-cumulative-dialog";
 cumulativeDialog.innerHTML='<form method="dialog"><h3>New Cumulative Selection</h3><label for="ls-cad-cumulative-name">Label</label><input id="ls-cad-cumulative-name" value="Cumulative Selection 1" maxlength="80" required/><p id="ls-cad-cumulative-error" hidden role="alert"></p><div class="ls-cad-dialog-actions"><button type="submit" value="ok" id="ls-cad-cumulative-ok">OK</button><button type="submit" value="cancel" id="ls-cad-cumulative-cancel">Cancel</button></div></form>';
 document.body.append(cumulativeDialog);
 const rangeDialog=document.createElement("dialog");
 rangeDialog.id="ls-cad-range-dialog";
 rangeDialog.setAttribute("aria-labelledby","ls-cad-range-title");
 rangeDialog.innerHTML=`<form id="ls-cad-range-form" method="dialog" novalidate>
   <div class="ls-cad-range-titlebar"><strong id="ls-cad-range-title">Range</strong><button type="button" id="ls-cad-range-close" title="Close">×</button></div>
   <div class="ls-cad-range-body">
    <label for="ls-cad-range-method">Entry method:</label><select id="ls-cad-range-method"><option value="Step">Step</option></select>
    <label for="ls-cad-range-start">Start:</label><input id="ls-cad-range-start" autocomplete="off" spellcheck="false" required/>
    <label for="ls-cad-range-step">Step:</label><input id="ls-cad-range-step" autocomplete="off" spellcheck="false" required/>
    <label for="ls-cad-range-stop">Stop:</label><input id="ls-cad-range-stop" autocomplete="off" spellcheck="false" required/>
    <label for="ls-cad-range-function">Function to apply to all values:</label><select id="ls-cad-range-function"><option value="none">None</option><option value="negate">Negate</option></select>
    <p id="ls-cad-range-error" role="alert" hidden></p>
   </div>
   <div class="ls-cad-range-actions"><button type="submit" value="replace" id="ls-cad-range-replace">Replace</button><button type="submit" value="add" id="ls-cad-range-add">Add</button><button type="button" id="ls-cad-range-cancel">Cancel</button></div>
  </form>`;
 document.body.append(rangeDialog);

 confirmDialog.addEventListener("close",()=>{
  if(confirmDialog.returnValue!=="yes")return;
  const id=confirmDialog.dataset.deleteId;
  if(!nodes.some(n=>n.id===id))return;
  checkpoint();const i=index(id);nodes.splice(i,1);
  selected=nodes.length?nodes[Math.min(i,nodes.length-1)].id:"root";
  limit=null;loadFields();say("Deleted "+id+" from Geometry 1. Undo can restore the object.");
 });
 const cur=()=>nodes.find(n=>n.id===selected),index=id=>nodes.findIndex(n=>n.id===id);
 const physSignature=()=>{
  const structural=nodes.map(n=>n.kind==="circle"?[n.id,n.kind,n.enabled,n.type,n.radius,n.sector,n.x,n.y,n.layer,n.layerExtra||[]]:[n.id,n.kind,n.enabled,n.inputIds,n.keep,n.angle,n.cx,n.cy]);
  // Resolve parameter expressions as well as literal expressions. An edit of d/rb
  // changes the geometry even if Circle 1/2 still display the same expressions.
  const evaluated=nodes.filter(n=>n.kind==="circle").map(n=>{
   try{return [n.id,...[n.radius,n.x,n.y,n.sector,n.layer,...(n.layerExtra||[])].map(x=>svc.calculate(x))];}
   catch(e){return [n.id,"parameter error",e.message];}
  });
  return JSON.stringify([structural,evaluated]);
 };
 let lastAppliedSignature=physSignature();
 function status(){
  const base={kind:"radial",featureCount:nodes.length,pending:physSignature()!==lastAppliedSignature,supported:false,domains:[]};
  try{
   const extra=nodes.find(n=>n.enabled&&(!["c1","c2"].includes(n.id)||n.kind!=="circle"));
   if(extra)throw Error(extra.label+" is preview-only; disable it before radial FEM.");
   const c1=nodes.find(n=>n.id==="c1"),c2=nodes.find(n=>n.id==="c2");
   if(!c1?.enabled||!c2?.enabled)throw Error("Original Circle 1 and Circle 2 must be enabled.");
   const check=n=>{
    if(n.type!=="Solid"||Math.abs(evalNumber(n.sector)-360)>1e-9||getLayers(n).some(t=>t>1e-12))throw Error(n.label+" must be a 360° Solid with no layers.");
    if(Math.abs(evalNumber(n.x))>1e-12||Math.abs(evalNumber(n.y))>1e-12)throw Error("Radial FEM requires circles centered at (0,0).");
    return evalNumber(n.radius)*1e6;
   };
   const outer=check(c1),core=check(c2);
   if(!(outer>=15&&outer<=150&&core>=.5&&core<=60&&outer>core))throw Error("Core and outer radii exceed supported radial FEM limits.");
   base.supported=true;base.outerRadiusUm=outer;base.coreRadiusUm=core;
   base.domains=[{id:1,name:"Core",feature:"c2"},{id:2,name:"Cladding",feature:"c1"}];
   return base;
  }catch(e){base.reason=String(e.message||e);return base;}
 }
 function emitCadState(){document.dispatchEvent(new CustomEvent("ls-cad-state",{detail:status()}));}
 window.LSWaveCADService=Object.freeze({
  status,apply:()=>apply(),
  entities:()=>compute().map(({entityKey,owner,source,r,x,y,layers,derived})=>({entityKey,owner,source,r,x,y,layers:[...layers],derived})),
  // Live named Geometry selections, including the features that contribute
  // their resulting entities. Materials may reference these by stable ID.
  domains:()=>materialDomains().map(d=>({...d,holeCircles:d.holeCircles.map(h=>({...h}))})),
  adjacentEntities:(level="Boundary")=>adjacentEntities(level).map(d=>({...d})),
  cumulativeSelections:()=>cumulativeSelections.map(row=>({
   id:row.id,label:row.label,featureIds:[...row.featureIds]
  }))
 });
 const evalNumber=s=>svc.calculate(s);
 function evalLength(expr){
  const value=evalNumber(expr);
  const bare=/^[+-]?(?:(?:\d+(?:\.\d*)?)|(?:\.\d+))(?:[eE][+-]?\d+)?$/.test(String(expr).trim());
  return bare?value*({m:1,mm:1e-3,um:1e-6,nm:1e-9}[lengthUnit]||1):value;
 }
 function formatLength(m){
  const abs=Math.abs(m);
  if(abs>=1e-3)return unit(m*1e3)+" mm";
  if(abs>=1e-6)return unit(m*1e6)+" µm";
  if(abs>=1e-9)return unit(m*1e9)+" nm";
  return m.toExponential(4)+" m";
 }
 function visibleExtent(){
  const shapes=compute();
  return Math.max(1e-12,...shapes.map(shape=>shape.r+Math.hypot(shape.x,shape.y)));
 }
 function lockViewIfNeeded(){
  const extent=visibleExtent();
  if(viewExtent===null||extent/viewExtent<0.12||extent/viewExtent>8){
   // Recover automatically from legacy 1 m → micrometer expression changes.
   // Keep a stable scale for normal radius and layer edits.
   viewExtent=extent;panX=0;panY=0;zoom=1;
  }
 }
 function shapeSummary(shape){
  let consumed=0;const layerInfo=(shape.layers||[]).filter(x=>x>0).map((t,i)=>{
   consumed+=t;return "L"+(i+1)+" "+formatLength(t)+" → inner "+formatLength(shape.r-consumed);
  });
  return shape.owner+" R="+formatLength(shape.r)+(layerInfo.length?" | "+layerInfo.join(" | "):"");
 }
 function getLayers(n){
  const all=[n.layer??"0[um]",...(n.layerExtra||[])].map(x=>evalLength(x||"0[um]"));
  if(all.some(v=>!Number.isFinite(v)||v<0))throw Error(n.label+": layer thickness must be finite and nonnegative.");
  return all;
 }
 function renderExtraLayers(n){
  const holder=$("ls-cad-extra-layers");if(!holder)return;
  holder.replaceChildren();
  (n.layerExtra||[]).forEach((thickness,i)=>{
   const label=document.createElement("label"),input=document.createElement("input");
   label.textContent="Layer "+(i+2);input.value=thickness;input.dataset.cadLayerExtra=String(i);
   input.setAttribute("aria-label","Layer "+(i+2)+" thickness");
   input.spellcheck=false;
   holder.append(label,input);
  });
  $("ls-cad-remove-layer").disabled=!(n.layerExtra?.length);
 }
 const say=(msg,error=false)=>{const e=$("ls-cad-status");e.textContent=msg;e.classList.toggle("error",error);};
 function fresh(kind,at){
  let i=1,pre=kind==="circle"?"c":kind==="mirror"?"mir":"rot";while(nodes.some(n=>n.id===pre+i))i++;
  if(kind==="circle")return circle(pre+i,"Circle "+i,nodes.length===0?"rb+wlr":"2.5[um]");
  const preceding=nodes.slice(0,at).filter(x=>x.kind==="circle"&&x.enabled);
  if(kind==="mirror")return {id:pre+i,kind:"mirror",label:"Mirror "+i,enabled:true,inputIds:[],inputRefs:[],keep:false,cx:"0[um]",cy:"0[um]",nx:"1",ny:"0",selection:false,color:"None"};
  return {id:pre+i,kind:"rotate",label:"Rotate "+i,enabled:true,inputIds:preceding.length?[preceding.at(-1).id]:[],inputRefs:null,keep:true,angle:"range(0,60,300)",cx:"0[um]",cy:"0[um]",selection:false,color:"None"};
 }

 function renderCumulativeTree(){
  const list=$("ls-cad-cumulative-list"),group=$("ls-cad-cumulative-group");
  if(!list||!group)return;
  group.hidden=!cumulativeSelections.length;
  list.replaceChildren();
  for(const row of cumulativeSelections){
   const button=document.createElement("button");
   button.type="button";button.dataset.cadCumulativeId=row.id;
   button.className="ls-tree-leaf"+(activeCumulativeId===row.id?" active":"");
   button.textContent="▧ "+row.label;
   list.append(button);
  }
 }
 function updateCumulativeOptions(node){
  const dropdown=$("ls-cad-contribute");if(!dropdown)return;
  dropdown.replaceChildren(new Option("None",""));
  for(const row of cumulativeSelections)dropdown.add(new Option(row.label,row.id));
  dropdown.value=node.cumulativeId||"";
 }
 function cumulativeHighlight(key){
  const row=cumulativeSelections.find(v=>v.id===activeCumulativeId);
  return row?.featureIds.some(id=>key.owner===id)||false;
 }
 function drawTree(){
  const list=$("ls-geometry-node-list");if(!list)return;
  renderCumulativeTree();
  list.innerHTML=nodes.map(n=>`<button type="button" data-cad-node="${esc(n.id)}" class="ls-tree-leaf ls-cad-node ${n.id===selected?"active":""} ${n.enabled?"":"ls-cad-node-disabled"}" title="Click: settings · right-click: context menu"><span class="ls-cad-symbol">${n.kind==="circle"?"◉":n.kind==="mirror"?"⇋":"↻"}</span><span>${esc(n.label)} <small>(${esc(n.id)})</small></span></button>`).join("");
 }
 function loadFields(){
  const cumulative=activeCumulativeId?cumulativeSelections.find(v=>v.id===activeCumulativeId):null;
  $("ls-cad-cumulative-settings").hidden=!cumulative;
  if(cumulative){
   $("ls-cad-root-settings").hidden=true;$("ls-cad-feature-settings").hidden=true;$("ls-cad-build-selected").disabled=true;
   $("ls-cad-subtitle").textContent="Cumulative Selection";
   $("ls-cad-cumulative-label").value=cumulative.label;
   $("ls-cad-cumulative-summary").textContent="Contributing Geometry features: "+(cumulative.featureIds.map(id=>nodes.find(n=>n.id===id)?.label||id).join(", ")||"None");
   drawTree();draw();return;
  }
  const n=cur();
   $("ls-cad-root-settings").hidden=!!n;$("ls-cad-feature-settings").hidden=!n;
   $("ls-cad-build-selected").disabled=!n;
   if(!n){
    $("ls-cad-subtitle").textContent="Geometry";
    $("ls-cad-root-label").value=rootLabel;
    $("ls-cad-root-unit").value=lengthUnit;
    $("ls-cad-angle-unit").value=angleUnit;
    $("ls-cad-constraints").value=constraintMode;
    $("ls-cad-repair").value=repairTolerance;
    $("ls-cad-scale-unit").checked=scaleWhenUnitsChange;
    $("ls-cad-build-new").checked=buildOnNew;
    $("ls-cad-build-leave").checked=buildOnLeave;
    $("ls-cad-root-sequence").textContent=nodes.map((x,i)=>`${i+1}. ${x.label} (${x.id}) · ${x.enabled?"enabled":"disabled"}`).join("\n");
    drawTree();draw();return;
   }
  if(!["rotate","mirror"].includes(n.kind))selectingRotateInputs=false;
   $("ls-cad-subtitle").textContent=n.kind==="circle"?"Circle":n.kind==="mirror"?"Mirror":"Rotate";$("ls-cad-label").value=n.label;
  $("ls-cad-circle").hidden=n.kind!=="circle";$("ls-cad-rotate").hidden=!["rotate","mirror"].includes(n.kind);
  if(n.kind==="circle"){
   for(const k of ["type","radius","sector","base","x","y","rotation","layer"])$("ls-cad-"+k).value=n[k];
   renderExtraLayers(n);
  }else{
   $("ls-cad-angle").value=n.angle;$("ls-cad-cx").value=n.cx;$("ls-cad-cy").value=n.cy;$("ls-cad-keep").checked=n.keep;
    $("ls-cad-activate-selection").setAttribute("aria-pressed",String(selectingRotateInputs));
    $("ls-cad-activate-selection").classList.toggle("active",selectingRotateInputs);
    $("ls-cad-selection-hint").textContent=selectingRotateInputs?"Graphics selection active. Click an existing circle to add or remove it.":"Activate Selection to highlight/select c2 in Graphics.";
   $("ls-cad-rotate-angle-section").hidden=n.kind!=="rotate";
   $("ls-cad-rotate-center-section").hidden=n.kind!=="rotate";
   $("ls-cad-mirror-point-section").hidden=n.kind!=="mirror";
   $("ls-cad-mirror-normal-section").hidden=n.kind!=="mirror";
   if(n.kind==="mirror"){for(const k of ["cx","cy","nx","ny"])$("ls-cad-mirror-"+k).value=n[k];}
   const list=$("ls-cad-inputs");list.replaceChildren();
   const options=availableInputs(n);
   if(!options.length)list.textContent="Create a preceding Circle or Rotate output, then activate selection.";
   const checked=selectedInputKeys(n,options);
   for(const shape of options){
    const label=document.createElement("label"),cb=document.createElement("input");
    cb.type="checkbox";cb.dataset.cadInput=shape.source;cb.dataset.cadInputRef=shape.entityKey;
    cb.checked=checked.has(shape.entityKey);
    label.append(cb,document.createTextNode(" "+entityLabel(shape)));list.append(label);
   }
  }
  $("ls-cad-selection").checked=n.selection;$("ls-cad-physics").disabled=!n.selection;$("ls-cad-color").value=n.color;
   updateCumulativeOptions(n);
  drawTree();draw();
 }

 function splitAngleExpressions(source){
  const out=[];let depth=0,start=0;
  const text=String(source||"").trim();
  for(let i=0;i<text.length;i++){
   if(text[i]==="(")depth++;
   else if(text[i]===")"){depth--;if(depth<0)throw Error("Unexpected ) in Rotate angle");}
   else if(text[i]===","&&depth===0){out.push(text.slice(start,i).trim());start=i+1;}
  }
  if(depth!==0)throw Error("Unbalanced Rotate angle parentheses");
  out.push(text.slice(start).trim());
  if(out.some(x=>!x))throw Error("Enter an angle or range(start,step,stop)");
  return out;
 }
 function angleList(source){
  const parts=splitAngleExpressions(source),angles=[];
  for(const part of parts){
   const match=/^range\s*\((.*)\)$/i.exec(part);
   if(!match){
    const angle=evalNumber(part);
    if(!Number.isFinite(angle))throw Error("Rotation angle must be finite");
    angles.push(angle);
   }else{
    const terms=match[1].split(",").map(x=>x.trim());
    if(terms.length!==3||terms.some(x=>!x))throw Error("range requires start, step, stop");
    const [start,step,stop]=terms.map(evalNumber);
    if(![start,step,stop].every(Number.isFinite)||step===0||(stop-start)*step<0)throw Error("Invalid Rotate range direction or zero step");
    const count=Math.floor((stop-start)/step+1e-10)+1;
    if(count<1||count>72)throw Error("Rotate range must have 1–72 angles");
    for(let i=0;i<count;i++)angles.push(start+i*step);
   }
   if(angles.length>72)throw Error("Rotate study supports at most 72 angles");
  }
  return angles;
 }
 function showRangeDialog(){
  if(cur()?.kind!=="rotate")return;
  const raw=cur().angle.trim();
  const match=/^range\s*\(\s*([^,]+)\s*,\s*([^,]+)\s*,\s*([^,)]+)\s*\)$/i.exec(raw);
  $("ls-cad-range-start").value=match?match[1].trim():"0";
  $("ls-cad-range-step").value=match?match[2].trim():"60";
  $("ls-cad-range-stop").value=match?match[3].trim():"300";
  $("ls-cad-range-function").value="none";
  $("ls-cad-range-error").textContent="";
  $("ls-cad-range-error").hidden=true;
  rangeDialog.showModal();
  $("ls-cad-range-start").focus();
 }

 function compute(upTo=null){
  const parameters=svc.evaluate();
  if(parameters.errors.size)throw Error([...parameters.errors.values()][0]);
  const shapes=[];
  const max=upTo===null?(limit===null?nodes.length-1:Math.min(limit,nodes.length-1)):Math.min(upTo,nodes.length-1);
  for(let i=0;i<=max;i++){
   const n=nodes[i];if(!n?.enabled)continue;
   if(n.kind==="circle"){
    const r=evalLength(n.radius),x=evalLength(n.x),y=evalLength(n.y),sector=evalNumber(n.sector),deg=evalNumber(n.rotation),layers=getLayers(n);
    if(!(r>0&&r<=10000)||!(sector>0&&sector<=360))throw Error(n.label+": invalid radius or sector");
    if(layers.reduce((sum,t)=>sum+t,0)>=r-1e-15)throw Error(n.label+": total layer thickness must be smaller than radius");
    shapes.push({id:n.id,entityKey:n.id,source:n.id,owner:n.id,kind:"circle",r,x,y,sector,deg,solid:n.type==="Solid",layers,selection:n.selection,color:n.color,derived:false});
   }else if(n.kind==="rotate"||n.kind==="mirror"){
    const available=[...shapes];
    const input=Array.isArray(n.inputRefs)?available.filter(shape=>n.inputRefs.includes(shape.entityKey)):available.filter(shape=>n.inputIds.includes(shape.owner));
    if(!input.length)continue;
    const px=evalLength(n.cx),py=evalLength(n.cy);
    if(!n.keep){
     const keys=new Set(input.map(shape=>shape.entityKey));
     for(let j=shapes.length-1;j>=0;j--)if(keys.has(shapes[j].entityKey))shapes.splice(j,1);
    }
    if(n.kind==="rotate"){
     const angles=angleList(n.angle);
     for(const shape of input)for(const [angleIndex,deg] of angles.entries()){
      const theta=deg*Math.PI/180,dx=shape.x-px,dy=shape.y-py;
      shapes.push({...shape,id:n.id,entityKey:n.id+"/"+shape.entityKey+"/"+angleIndex,owner:n.id,kind:"rotate",x:px+dx*Math.cos(theta)-dy*Math.sin(theta),y:py+dx*Math.sin(theta)+dy*Math.cos(theta),deg:shape.deg+deg,selection:n.selection,color:n.color,derived:true});
     }
    }else{
     const nx=evalNumber(n.nx),ny=evalNumber(n.ny),norm=nx*nx+ny*ny;
     if(!Number.isFinite(norm)||norm<=1e-18)throw Error(n.label+": the reflection normal vector must not be zero");
     for(const [entityIndex,shape] of input.entries()){
      const projection=((shape.x-px)*nx+(shape.y-py)*ny)/norm;
      shapes.push({...shape,id:n.id,entityKey:n.id+"/"+shape.entityKey+"/"+entityIndex,owner:n.id,kind:"mirror",x:shape.x-2*projection*nx,y:shape.y-2*projection*ny,deg:-shape.deg,selection:n.selection,color:n.color,derived:true});
     }
    }
   }
   if(shapes.length>450)throw Error("Too many objects for preview (max 450)");
  }
  return shapes;
 }

 /**
  * 2D preview-domain partition. A Circle with inward layers is split into
  * radial shells (four quadrant regions per shell, matching the four geometry
  * partition lines) plus a center disk. Smaller circular features then form
  * individually addressable regions. This is geometry selection, not a
  * COMSOL Boolean Union or FEM domain/mesh.
  */
 function materialDomains(){
  const shapes=compute(),domains=[];
  const add=(shape,part,label,outerR,innerR,startAngle=null,endAngle=null)=>{
   const key=shape.entityKey+"/"+part;
   const d={id:domains.length+1,key,entityKey:shape.entityKey,owner:shape.owner,
     source:shape.source,label,part,cx:shape.x,cy:shape.y,outerR,innerR,
     startAngle,endAngle,holeCircles:[]};
   domains.push(d);return d;
  };
  for(const shape of shapes){
   if(shape.sector<359.999||!shape.solid){
    add(shape,"disk",""+shape.entityKey+" · object",shape.r,0);
    continue;
   }
   let r=shape.r,active=0;
   for(const [idx,t] of (shape.layers||[]).entries()){
    if(t<=0)continue;
    const next=r-t;
    if(next<=0)break;
    for(let quadrant=0;quadrant<4;quadrant++)
     add(shape,"layer"+idx+"-q"+quadrant,shape.entityKey+" · layer "+(idx+1)+" sector "+(quadrant+1),r,next,quadrant*90,(quadrant+1)*90);
    r=next;active++;
   }
   add(shape,active?"core":"disk",shape.entityKey+(active?" · inner domain":" · disk"),r,0);
  }
  // Carve the smaller circular features out of a surrounding core region in
  // the SVG preview, so the user can select either the core or an air hole.
  for(const domain of domains){
   if(domain.innerR>0||domain.startAngle!==null)continue;
   const contained=domains.filter(other=>other!==domain&&other.innerR===0&&other.startAngle===null&&
    other.outerR<domain.outerR-1e-12&&
    Math.hypot(other.cx-domain.cx,other.cy-domain.cy)+other.outerR<=domain.outerR+1e-12);
   domain.holeCircles=contained.map(d=>({cx:d.cx,cy:d.cy,r:d.outerR,key:d.key}));
  }
  return domains;
 }

 // Stable, individually selectable boundaries from current preview geometry.
 // These are rendered SVG geometrical edges; no FEM/Boolean topology claimed.
 function adjacentEntities(level="Boundary"){
  if(level==="Domain")return materialDomains().map(d=>({
   id:d.id,key:d.key,entityKey:d.entityKey,owner:d.owner,kind:"domain",
   label:d.label,cx:d.cx,cy:d.cy,outerR:d.outerR,innerR:d.innerR,
   startAngle:d.startAngle,endAngle:d.endAngle,holeCircles:d.holeCircles
  }));
  const shapes=compute(),items=[];
  const add=(shape,part,kind,r,x1=null,y1=null,x2=null,y2=null)=>{
   const key=shape.entityKey+"/"+part;
   items.push({id:items.length+1,key,entityKey:shape.entityKey,owner:shape.owner,
    label:shape.entityKey+" · "+part,kind,cx:shape.x,cy:shape.y,r,
    x1,y1,x2,y2});
  };
  for(const shape of shapes){
   if(level==="Point"){add(shape,"center","point",0);continue;}
   add(shape,"outer-boundary","circle",shape.r);
   if(shape.sector<359.999)continue;
   let radius=shape.r;
   for(const [i,t] of (shape.layers||[]).entries()){
    if(!(t>0))continue;
    const inner=radius-t;
    if(inner<=0)break;
    add(shape,"layer-"+(i+1)+"-inner","circle",inner);
    for(let k=0;k<4;k++){
     const theta=k*Math.PI/2;
     add(shape,"layer-"+(i+1)+"-radial-"+k,"line",0,
      shape.x+inner*Math.cos(theta),shape.y+inner*Math.sin(theta),
      shape.x+radius*Math.cos(theta),shape.y+radius*Math.sin(theta));
    }
    radius=inner;
   }
  }
  return items;
 }
 function availableInputs(n){
  const idx=index(n?.id);
  return idx<0?[]:compute(idx-1);
 }
 function selectedInputKeys(n,candidates=availableInputs(n)){
  return Array.isArray(n.inputRefs)?new Set(n.inputRefs):new Set(candidates.filter(x=>n.inputIds.includes(x.owner)).map(x=>x.entityKey));
 }
 function entityLabel(shape){
  if(!shape.derived)return shape.source+" · "+(nodes.find(n=>n.id===shape.source)?.label||"Circle");
  const degrees=((shape.deg%360)+360)%360;
  return shape.entityKey+" ("+unit(degrees)+"°)";
 }
 function drawShape(x,y,r,sector,deg){
  if(sector>=359.999)return `<circle cx="${x}" cy="${y}" r="${r}"/>`;
  const a=(deg-90)*Math.PI/180,b=a+sector*Math.PI/180,x1=x+r*Math.cos(a),y1=y+r*Math.sin(a),x2=x+r*Math.cos(b),y2=y+r*Math.sin(b);
  return `<path d="M${x} ${y} L${x1} ${y1} A${r} ${r} 0 ${sector>180?1:0} 1 ${x2} ${y2} Z"/>`;
 }


 function materialDomainPath(domain,ox,oy,scale){
  const x=ox+domain.cx*scale,y=oy-domain.cy*scale,outer=domain.outerR*scale;
  if(!(outer>0))return "";
  const circle=(cx,cy,r)=>`M ${cx+r} ${cy} A ${r} ${r} 0 1 0 ${cx-r} ${cy} A ${r} ${r} 0 1 0 ${cx+r} ${cy} Z`;
  if(domain.startAngle!==null){
   const inner=domain.innerR*scale;
   const polar=(r,deg)=>{
    const t=deg*Math.PI/180;
    return [x+r*Math.cos(t),y-r*Math.sin(t)];
   };
   const [ax,ay]=polar(outer,domain.startAngle),[bx,by]=polar(outer,domain.endAngle);
   if(inner<=1e-9)return `M ${ax} ${ay} A ${outer} ${outer} 0 0 0 ${bx} ${by} L ${x} ${y} Z`;
   const [cx,cy]=polar(inner,domain.endAngle),[dx,dy]=polar(inner,domain.startAngle);
   return `M ${ax} ${ay} A ${outer} ${outer} 0 0 0 ${bx} ${by} L ${cx} ${cy} A ${inner} ${inner} 0 0 1 ${dx} ${dy} Z`;
  }
  let path=circle(x,y,outer);
  for(const hole of domain.holeCircles||[]){
   const cx=ox+hole.cx*scale,cy=oy-hole.cy*scale,r=hole.r*scale;
   if(r>0)path+=" "+circle(cx,cy,r);
  }
  return path;
 }

 // Use physical click coordinates to identify exactly one domain, independent
 // of SVG fill transparency or overlap among rotated air-hole objects.
 function materialDomainAtScreen(event){
  const viewport=$("ls-cad-view"),matrix=viewport?.getScreenCTM();
  if(!matrix)return null;
  const point=viewport.createSVGPoint();point.x=event.clientX;point.y=event.clientY;
  const p=point.matrixTransform(matrix.inverse());
  const plot={left:55,top:31,right:618,bottom:452};
  const shapes=compute(),actualExtent=Math.max(1e-12,...shapes.map(shape=>shape.r+Math.hypot(shape.x,shape.y)));
  const referenceExtent=viewExtent??actualExtent;
  const scale=Math.min(plot.right-plot.left,plot.bottom-plot.top)*.42*zoom/referenceExtent;
  if(!(scale>0)||p.x<plot.left||p.x>plot.right||p.y<plot.top||p.y>plot.bottom)return null;
  const ox=(plot.left+plot.right)/2+panX,oy=(plot.top+plot.bottom)/2+panY;
  const x=(p.x-ox)/scale,y=(oy-p.y)/scale;
  let best=null,bestArea=Infinity;
  for(const domain of materialDomains()){
   const dx=x-domain.cx,dy=y-domain.cy,r=Math.hypot(dx,dy);
   if(r>domain.outerR||r<domain.innerR)continue;
   if(domain.startAngle!==null){
    const angle=(Math.atan2(dy,dx)*180/Math.PI+360)%360;
    if(angle<domain.startAngle||angle>=domain.endAngle)continue;
   }
   if((domain.holeCircles||[]).some(hole=>Math.hypot(x-hole.cx,y-hole.cy)<hole.r))continue;
   const area=(domain.outerR**2-domain.innerR**2)*(domain.startAngle===null?1:(domain.endAngle-domain.startAngle)/360);
   // Prefer the smallest visible region. Coincident circles pick the
   // last-built object; the other remains selectable from the ID list.
   if(area<=bestArea*(1+1e-9)){best=domain;bestArea=area;}
  }
  return best;
 }
 function materialDomainsSVG(ox,oy,scale,mini){
  if(mini||!document.body.classList.contains("ls-material-active")||materialEntityLevel!=="Domain")return "";
  const picking=materialPickMode;
  const items=materialDomains().map(domain=>{
   const chosen=materialHighlightDomainIds.has(domain.id),d=materialDomainPath(domain,ox,oy,scale);
   const name="Domain "+domain.id+" — "+domain.label;
   return `<path class="ls-cad-material-region${chosen?" is-selected":""}" data-cad-domain-id="${domain.id}" data-cad-domain-key="${esc(domain.key)}" data-domain-owner="${esc(domain.owner)}" aria-label="${esc(name)}" d="${d}" fill-rule="evenodd" fill="${chosen?"#9b94e7":"transparent"}" fill-opacity="${chosen?".63":"0"}" stroke="${chosen?"#5149cb":"none"}" stroke-width="${chosen?"1.35":"0"}" pointer-events="${picking?"visiblePainted":"none"}"><title>${esc(name)}</title></path>`;
  }).join("");
  return `<g class="ls-cad-material-domain-overlay" data-domain-picker="${picking?"active":"inactive"}">${items}</g>`;
 }

 function adjacentEntitiesSVG(ox,oy,scale,mini){
  if(mini||!adjacentView)return "";
  const entities=adjacentEntities(adjacentLevel);
  return `<g class="ls-cad-adjacent-overlay" data-adjacent-level="${adjacentLevel}" pointer-events="${adjacentPickMode?"all":"none"}">${entities.map(entity=>{
   const chosen=adjacentSelectedKeys.has(entity.key);
   const name=esc(adjacentLevel+" "+entity.id+": "+entity.label);
   let geometry;
   if(entity.kind==="domain"){
    geometry=`<path class="ls-cad-adjacent-mark" d="${materialDomainPath(entity,ox,oy,scale)}" fill="none" fill-rule="evenodd" stroke="${chosen?"#164bfa":"transparent"}" stroke-width="1.8" pointer-events="none"/>`;
   }else if(entity.kind==="point"){
    const x=ox+entity.cx*scale,y=oy-entity.cy*scale;
    geometry=`<circle class="ls-cad-adjacent-mark" cx="${x}" cy="${y}" r="3" fill="${chosen?"#2e61e3":"transparent"}" stroke="${chosen?"#234cda":"#a7b2bc"}" stroke-width="2"/>`;
   }else if(entity.kind==="line"){
    geometry=`<line class="ls-cad-adjacent-mark" x1="${ox+entity.x1*scale}" y1="${oy-entity.y1*scale}" x2="${ox+entity.x2*scale}" y2="${oy-entity.y2*scale}" stroke="${chosen?"#195cf4":"#626f78"}" stroke-opacity="${chosen?"1":".16"}" stroke-width="1.8"/>`;
   }else{
    geometry=`<circle class="ls-cad-adjacent-mark" cx="${ox+entity.cx*scale}" cy="${oy-entity.cy*scale}" r="${entity.r*scale}" fill="none" stroke="${chosen?"#195cf4":"transparent"}" stroke-width="1.8"/>`;
   }
   const hit=entity.kind==="point"
    ?`<circle cx="${ox+entity.cx*scale}" cy="${oy-entity.cy*scale}" r="10" fill="transparent" pointer-events="all"/>`
    :entity.kind==="domain" ? `<path d="${materialDomainPath(entity,ox,oy,scale)}" fill="transparent" fill-rule="evenodd" pointer-events="all"/>`
    :entity.kind==="line"?`<line x1="${ox+entity.x1*scale}" y1="${oy-entity.y1*scale}" x2="${ox+entity.x2*scale}" y2="${oy-entity.y2*scale}" stroke="transparent" stroke-width="12" pointer-events="stroke"/>`
    :`<circle cx="${ox+entity.cx*scale}" cy="${oy-entity.cy*scale}" r="${entity.r*scale}" fill="none" stroke="transparent" stroke-width="12" pointer-events="stroke"/>`;
   return `<g class="ls-cad-adjacent-entity${chosen?" is-adjacent-selected":""}" data-adjacent-key="${esc(entity.key)}" data-adjacent-entity-id="${entity.id}" aria-label="${name}"><title>${name}</title>${geometry}${adjacentPickMode?hit:""}</g>`;
  }).join("")}</g>`;
 }

 // Reliable geometric click fallback for mouse/touch events. SVG pointer
 // hit-testing differs across browser engines for transparent stroke-only
 // circles. Pick the nearest *actual* boundary in SVG user units instead.
 function nearestAdjacentAtScreen(event){
  if(!adjacentPickMode||adjacentLevel==="Domain")return null;
  const svgNode=$("ls-cad-view");
  if(!svgNode?.getScreenCTM())return null;
  const mouse=svgNode.createSVGPoint();
  mouse.x=event.clientX;mouse.y=event.clientY;
  const p=mouse.matrixTransform(svgNode.getScreenCTM().inverse());
  const plot={left:55,top:31,right:640-22,bottom:500-48};
  const shapes=compute();
  const actual=Math.max(1e-12,...shapes.map(shape=>shape.r+Math.hypot(shape.x,shape.y)));
  const reference=viewExtent??actual;
  const scale=Math.min(plot.right-plot.left,plot.bottom-plot.top)*.42*zoom/reference;
  const ox=(plot.left+plot.right)/2+panX,oy=(plot.top+plot.bottom)/2+panY;
  const pointLine=(x,y,a,b,c,d)=>{
   const vx=c-a,vy=d-b,len=vx*vx+vy*vy;
   const t=len>0?Math.max(0,Math.min(1,((x-a)*vx+(y-b)*vy)/len)):0;
   return Math.hypot(x-a-t*vx,y-b-t*vy);
  };
  let nearest=null,best=Infinity;
  for(const entity of adjacentEntities(adjacentLevel)){
   let distance=Infinity;
   if(entity.kind==="circle"){
    const cx=ox+entity.cx*scale,cy=oy-entity.cy*scale;
    distance=Math.abs(Math.hypot(p.x-cx,p.y-cy)-entity.r*scale);
   }else if(entity.kind==="line"){
    distance=pointLine(p.x,p.y,ox+entity.x1*scale,oy-entity.y1*scale,
      ox+entity.x2*scale,oy-entity.y2*scale);
   }else if(entity.kind==="point"){
    distance=Math.hypot(p.x-(ox+entity.cx*scale),p.y-(oy-entity.cy*scale));
   }
   if(distance<best){best=distance;nearest=entity;}
  }
  return best<=9?nearest:null;
 }
 function niceTick(value){
  if(!(value>0&&Number.isFinite(value)))return 1;
  const exponent=Math.floor(Math.log10(value));
  let chosen=value;
  let delta=Infinity;
  for(const exp of [exponent-1,exponent,exponent+1]){
   const power=Math.pow(10,exp);
   for(const factor of [1,2,2.5,5,10]){
    const candidate=power*factor;
    const error=Math.abs(Math.log(candidate/value));
    if(error<delta){delta=error;chosen=candidate;}
   }
  }
  return chosen;
 }
 function svg(w,h,mini=false){
  const shapes=compute();
  const adjacentScene=adjacentView&&!mini,showGrid=adjacentScene?adjacentGridOn:gridOn;
  const plot=mini?{left:26,top:22,right:w-13,bottom:h-23}:{left:55,top:31,right:w-22,bottom:h-48};
  const pw=plot.right-plot.left,ph=plot.bottom-plot.top;
  const actualExtent=Math.max(1e-12,...shapes.map(shape=>shape.r+Math.hypot(shape.x,shape.y)));
  const referenceExtent=viewExtent??actualExtent;
  const scale=Math.min(pw,ph)*(mini?.40:.42)*zoom/referenceExtent;
  const ox=(plot.left+plot.right)/2+(mini?0:panX),oy=(plot.top+plot.bottom)/2+(mini?0:panY);
  const radiusDisplay=Math.min(pw,ph)/(scale*2);
  const units=radiusDisplay<1e-3?{mult:1e6,symbol:"µm",power:"−6"}:radiusDisplay<1?{mult:1e3,symbol:"mm",power:"−3"}:{mult:1,symbol:"m",power:"0"};
  const spacing=niceTick(ph/scale/10);
  const decimal=value=>{const n=Number(value.toPrecision(6));return Object.is(n,-0)?"0":String(n);};
  let grid="",xTicks="",yTicks="";
  const clipId=mini?"ls-cad-plot-clip-small":"ls-cad-plot-clip-main";
  const lowerX=(plot.left-ox)/scale,upperX=(plot.right-ox)/scale;
  const lowerY=(oy-plot.bottom)/scale,upperY=(oy-plot.top)/scale;
  if(showGrid||!mini){
   // Grid and tick positions are computed in meters, not in arbitrary pixel increments.
   const firstX=Math.ceil(lowerX/spacing),lastX=Math.floor(upperX/spacing);
   const firstY=Math.ceil(lowerY/spacing),lastY=Math.floor(upperY/spacing);
   for(let i=firstX;i<=lastX&&i<firstX+70;i++){
    const position=i*spacing,x=ox+position*scale;
    if(showGrid)grid+=`<path d="M ${x} ${plot.top} V ${plot.bottom}" stroke="${i===0?"#b5c4d0":"#deebf3"}" stroke-width="${i===0?1.2:.85}"/>`;
    if(!mini)xTicks+=`<path d="M ${x} ${plot.bottom} v 5" stroke="#465361" stroke-width="1"/><text data-axis="x" data-coordinate-m="${position}" x="${x}" y="${plot.bottom+19}" text-anchor="middle" font-size="12" fill="#354a5a">${decimal(position*units.mult)}</text>`;
   }
   for(let i=firstY;i<=lastY&&i<firstY+70;i++){
    const position=i*spacing,y=oy-position*scale;
    if(showGrid)grid+=`<path d="M ${plot.left} ${y} H ${plot.right}" stroke="${i===0?"#b5c4d0":"#deebf3"}" stroke-width="${i===0?1.2:.85}"/>`;
    if(!mini)yTicks+=`<path d="M ${plot.left-5} ${y} h 5" stroke="#465361" stroke-width="1"/><text data-axis="y" data-coordinate-m="${position}" x="${plot.left-8}" y="${y+4}" text-anchor="end" font-size="12" fill="#354a5a">${decimal(position*units.mult)}</text>`;
   }
  }
  const isSelectedInput=shape=>!adjacentScene&&selectingRotateInputs&&["rotate","mirror"].includes(cur()?.kind)&&selectedInputKeys(cur()).has(shape.entityKey);
   // Draw the highlighted source object last so zero-degree rotated copies do not hide it.
   const orderedShapes=selectingRotateInputs?[...shapes].sort((a,b)=>Number(isSelectedInput(a))-Number(isSelectedInput(b))):shapes;
   const items=orderedShapes.map(shape=>{
   const x=ox+shape.x*scale,y=oy-shape.y*scale,r=shape.r*scale;
   const stroke=adjacentScene?"#222":shape.derived?"#7a8793":"#41484f";
   const colors={Blue:"#9cc8ef",Orange:"#efb584",Green:"#a5d7b1"};
   const fill=adjacentScene&&shape.solid?"#ccc":shape.solid?(shape.selection&&colors[shape.color]?colors[shape.color]:"#ced0d1"):"none";
   let layersMarkup="",covered=0,outer=shape.r;
   const interior=[];
   if(shape.sector>=359.999){
    for(const [layerIndex,thickness] of (shape.layers||[]).entries()){
     if(thickness<=0)continue;
     covered+=thickness;const inner=shape.r-covered;
     if(inner<=0)break;
     const innerPx=inner*scale;
     interior.push(inner);
     layersMarkup+=`<circle class="ls-cad-layer-region" cx="${x}" cy="${y}" r="${innerPx}" fill="${adjacentScene?"#ccc":layerIndex%2?"#e4e7e9":"#b9c2c9"}" fill-opacity=".98" stroke="none" pointer-events="none"/>`;
     layersMarkup+=`<circle class="ls-cad-layer-boundary" cx="${x}" cy="${y}" r="${innerPx}" fill="none" stroke="#454d53" stroke-width="${mini?1:1.6}" pointer-events="none"/>`;
     // Radial spokes match COMSOL's displayed geometric partition guides.
     for(const deg of [0,90,180,270]){
      const a=deg*Math.PI/180,cs=Math.cos(a),sn=Math.sin(a);
      layersMarkup+=`<line class="ls-cad-layer-boundary" x1="${x+innerPx*cs}" y1="${y+innerPx*sn}" x2="${x+outer*scale*cs}" y2="${y+outer*scale*sn}" stroke="#525b64" stroke-width="${mini?1:1.3}" pointer-events="none"/>`;
     }
     outer=inner;
    }
   }
   const details=`${shape.owner} • Radius (outer) ${unit(shape.r*1e6)} µm${interior.length?" • Radius (inner) "+unit(interior.at(-1)*1e6)+" µm":""}`;
   const label=mini||adjacentScene?"":`<text x="${x}" y="${y-r-9}" font-size="11" text-anchor="middle" fill="#536d80">${esc(shape.owner)}${shape.derived?"*":""}</text>`;
   return `<g class="ls-cad-rendered-shape${isSelectedInput(shape)?" is-input-selected":""}${!adjacentScene&&materialEntityLevel!=="Domain"&&materialHighlightKeys.has(shape.entityKey)?" is-material-selected":""}${!adjacentScene&&cumulativeHighlight(shape)?" is-cumulative-selected":""}" data-cad-shape="${esc(shape.owner)}" data-cad-entity="${esc(shape.entityKey)}" data-cad-source="${esc(shape.source)}" data-outer-radius-m="${shape.r}" data-inner-radius-m="${interior.length?interior.at(-1):shape.r}" data-layer-count="${interior.length}" stroke="${stroke}" stroke-width="${mini?1.2:1.65}" ${shape.derived&&!adjacentScene?'stroke-dasharray="5 3"':""}><title>${esc(details)}</title><g class="ls-cad-rendered-surface" fill="${fill}" fill-opacity="${adjacentScene?"1":shape.derived?".55":".95"}">${drawShape(x,y,r,shape.sector,shape.deg)}</g>${layersMarkup}</g>${label}`;
  }).join("");
  const materialRegions=materialDomainsSVG(ox,oy,scale,mini);
  const adjacentRegions=adjacentEntitiesSVG(ox,oy,scale,mini);
  const labels=mini?"":`<g class="ls-cad-axis-labels">${xTicks}${yTicks}<text data-axis-units="true" x="${plot.left+5}" y="${plot.top+15}" font-size="12" fill="#2d455b">×10${units.power==="0"?"⁰":units.power==="−6"?"⁻⁶":"⁻³"} m</text><text x="${plot.right-6}" y="${h-7}" text-anchor="end" font-size="12" fill="#2d455b">x (${units.symbol})</text><text x="9" y="${plot.top+9}" font-size="12" fill="#2d455b">y</text></g>`;
  return `<rect width="${w}" height="${h}" fill="#fcfdfe"/><defs><clipPath id="${clipId}"><rect x="${plot.left}" y="${plot.top}" width="${pw}" height="${ph}"/></clipPath></defs><rect x="${plot.left}" y="${plot.top}" width="${pw}" height="${ph}" fill="#fff" stroke="#8495a4" stroke-width="1.2"/><g clip-path="url(#${clipId})">${grid}${items}${materialRegions}${adjacentRegions}</g>${labels}`;
 }
 function draw(){
  try{
   $("ls-cad-view").innerHTML=svg(640,500);
   $("ls-cad-mini-svg").innerHTML=svg(460,230,true);
    const pick=$("ls-cad-object-select");
    if(pick){
     pick.innerHTML='<option value="root">Geometry 1 · All</option>'+nodes.map(n=>`<option value="${esc(n.id)}">${esc(n.label)} (${esc(n.id)})</option>`).join("");
     pick.value=nodes.some(n=>n.id===selected)?selected:"root";
     const materialMode=document.body.classList.contains("ls-material-active");
     pick.disabled=materialMode||adjacentView;
     const graphicsLabel=graphics.querySelector('label[for="ls-cad-object-select"]');
     if(graphicsLabel)graphicsLabel.textContent=adjacentView?"Adjacent · "+adjacentLevel:materialMode?"Material Domain selection":"Graphics selection";
     const graphicsHint=graphics.querySelector(".ls-cad-graphics-select span");
     if(graphicsHint)graphicsHint.textContent=adjacentView
      ?adjacentPickMode?"영역 위에 마우스: 빨간 경계 · 클릭: 파란 선택 · 다시 클릭: 해제":"선택 일시 정지"
      :materialMode&&materialEntityLevel==="Domain"
      ?materialPickMode?"Manual selection active · click each filled area to select or deselect.":"Domain picking paused."
      :"빈 영역 드래그: 이동 · 휠: 확대/축소";
    }
   $("ls-cad-build-label").textContent=limit===null?"Build All":limit<0?"Build preceding (empty)":"Build through "+nodes[limit]?.label;
   const shapes=compute();
   const selectedShape=shapes.find(shape=>shape.owner===selected)||shapes.at(-1);
   $("ls-cad-legend").textContent=`${shapes.length} object(s) · ${selectedShape?shapeSummary(selectedShape):"No geometry"} · Zoom ${Math.round(zoom*100)}% · units ${lengthUnit==="um"?"µm":lengthUnit} · ${viewExtent===null?"Auto-fit (first Build locks scale)":"Fixed scale (use Fit ⌖ to reframe)"} · Geometry preview only`;
   if($("ls-cad-status").textContent.startsWith("Preview error:"))say("Graphics updated.");
   emitCadState();
  }catch(e){
   // Preserve the last valid view while a user is typing a partial expression.
   $("ls-cad-legend").textContent="Geometry expression error: "+e.message;
   if(!pane.hidden)say("Preview error: "+e.message,true);
  }
 }
 function open(id){
  activeCumulativeId=null;
  if(id&&id!==selected)selectingRotateInputs=false;
  if(id==="root")selected="root";else if(id&&nodes.some(n=>n.id===id))selected=id;
  if(selected!=="root"&&!cur())selected=nodes.at(-1)?.id;
  $("ls-config-pane").hidden=true;
  pane.hidden=false;graphics.hidden=false;
  side.classList.add("ls-config-active");document.body.classList.add("ls-configuration-open","ls-cad-active");
  loadFields();say("Edit Geometry features; 2D Graphics updates immediately. Radial FEM only changes after Apply.");
 }
 function openCumulative(id){
  const row=cumulativeSelections.find(r=>r.id===id);if(!row)return;
  activeCumulativeId=id;selectingRotateInputs=false;
  $("ls-config-pane").hidden=true;
  pane.hidden=false;graphics.hidden=false;side.classList.add("ls-config-active");
  document.body.classList.add("ls-configuration-open","ls-cad-active");
  loadFields();say("Cumulative Selection "+row.label+": "+row.featureIds.length+" contributing Geometry feature(s).");
 }
 function hideContext(){context.hidden=true;}
 function close(){
  selectingRotateInputs=false;activeCumulativeId=null;
  if(rangeDialog.open)rangeDialog.close("cancel");
  pane.hidden=true;graphics.hidden=true;side.classList.remove("ls-config-active");document.body.classList.remove("ls-configuration-open","ls-cad-active");hideContext();
 }
 function add(kind,pos=nodes.length){
  selectingRotateInputs=false;
  checkpoint();const n=fresh(kind,pos);nodes.splice(pos,0,n);selected=n.id;limit=null;open();say(`${n.label} added. Configure and Build All.`);
 }
 function modify(cmd,arg){
  const idx=index(selected),n=cur();
   if(cmd==="undo"||cmd==="redo"){historyMove(cmd==="undo");return;}
  switch(cmd){
   case "circle":case "rotate":case "mirror":add(cmd,arg==="before"?Math.max(idx,0):arg==="after"?idx+1:nodes.length);return;
   case "build-all":checkpoint();limit=null;try{lockViewIfNeeded();draw();say("Built All: "+compute().map(shapeSummary).join("; "));}catch(e){draw();say("Build failed: "+e.message,true);}return;
   case "build-selected":if(!n)return;checkpoint();limit=idx;try{lockViewIfNeeded();draw();say("Built Selected: "+compute().filter(shape=>shape.owner===n.id).map(shapeSummary).join("; "));}catch(e){draw();say("Build failed: "+e.message,true);}return;
   case "build-preceding":if(!n)return;checkpoint();limit=idx-1;try{lockViewIfNeeded();draw();say("Built Preceding complete.");}catch(e){draw();say("Build failed: "+e.message,true);}return;
   case "up":case "down":{const i=idx+(cmd==="up"?-1:1);if(i<0||i>=nodes.length)return;checkpoint();[nodes[idx],nodes[i]]=[nodes[i],nodes[idx]];limit=null;loadFields();say("Geometry order updated.");return;}
   case "duplicate":{if(!n)return;checkpoint();const copy=structuredClone(n),newid=fresh(n.kind,idx+1);copy.id=newid.id;copy.label=newid.label;nodes.splice(idx+1,0,copy);selected=copy.id;limit=null;loadFields();say(`Duplicated as ${copy.label}.`);return;}
   case "delete":{
    if(!n)return;
    confirmDialog.dataset.deleteId=n.id;
    $("ls-cad-delete-message").textContent='Are you sure you want to delete "'+n.label+' ('+n.id+')" ?';
    confirmDialog.returnValue="no";
    if(typeof confirmDialog.showModal==="function")confirmDialog.showModal();
    else if(window.confirm("Delete "+n.label+"?")){confirmDialog.returnValue="yes";confirmDialog.dispatchEvent(new Event("close"));}
    return;
   }
   case "toggle":{if(!n)return;checkpoint();n.enabled=!n.enabled;limit=null;loadFields();say(`${n.label} ${n.enabled?"enabled":"disabled"}.`);return;}
   case "rename":if(!n)return;open();$("ls-cad-label").focus();$("ls-cad-label").select();return;
   case "settings":open();return;
   case "copy":if(!n)return;clipboard=structuredClone(n);say(n.label+" copied in Geometry editor.");return;
   case "paste":if(!clipboard)return;{
    checkpoint();const obj=structuredClone(clipboard),item=fresh(obj.kind,Math.max(idx+1,0));
    obj.id=item.id;obj.label=item.label;
    nodes.splice(Math.max(idx+1,0),0,obj);selected=obj.id;limit=null;loadFields();
    say("Pasted "+obj.label+" into Geometry sequence.");return;
   }
   case "copy-code":{
    if(!n)return;const body=JSON.stringify(n,null,2);
    if(navigator.clipboard?.writeText)navigator.clipboard.writeText(body).then(()=>say("Feature JSON copied.")).catch(()=>say("Clipboard unavailable; use Export JSON.",true));
    else say("Clipboard unavailable; use Export JSON.",true);
    return;
   }
   case "properties":open();say(`${n.kind} node · ${n.id} · ${n.enabled?"enabled":"disabled"}`);return;
   case "help":say("Circle and Rotate support preview only for arbitrary geometry; General 2D FEM is not available.");return;
  }
 }
 function apply(){
  try{
   if(nodes.some(n=>n.enabled&&n.kind==="rotate"))throw Error("Disable Rotate operations before applying to Radial FEM.");
   if(nodes.some(n=>n.enabled&&n.kind==="circle"&&!["c1","c2"].includes(n.id)))throw Error("Extra circles are preview-only. Disable them to run two-circle Radial FEM.");
   const c1=nodes.find(n=>n.id==="c1"),c2=nodes.find(n=>n.id==="c2");
   if(!c1?.enabled||!c2?.enabled)throw Error("Circle 1 and Circle 2 must be enabled.");
   function r(c){
    const x=evalNumber(c.x),y=evalNumber(c.y),a=evalNumber(c.sector),layers=getLayers(c),radius=evalNumber(c.radius);
    if(c.type!=="Solid"||a!==360)throw Error("Radial FEM requires 360° solid circles.");
    if(Math.abs(x)>1e-12||Math.abs(y)>1e-12)throw Error("Radial FEM requires concentric centers at (0,0).");
    if(layers.some(t=>t>1e-12))throw Error("Extra layers are not supported by two-circle Radial FEM.");
    return radius*1e6;
   }
   const outer=r(c1),inner=r(c2),wl=evalNumber("wl")*1e9;
   // Nine-row COMSOL reference defaults do not contain dn_core.
   // Preserve the established radial FEM material input, but honor an explicitly
   // added/edited dn_core parameter when present.
   const dn=svc.rows().some(p=>p.name==="dn_core")?evalNumber("dn_core"):Number($("delta_n_core")?.value);
   if(!(outer>=15&&outer<=150&&inner>=.5&&inner<=60&&outer>inner&&wl>=250&&wl<=3600&&dn>0&&dn<.1))throw Error("Solver limits: outer 15–150 µm; core 0.5–60 µm; wl 250–3600 nm; 0<dn<0.1.");
   const bg=evalNumber("nbg");
   if($("ls-cad-fix-nbg").checked&&!(bg>=1&&bg<2))throw Error("nbg must be between 1 and 2.");
   if(document.body.classList.contains("hcf-selection")||document.body.classList.contains("vector-selection")||$("profile").value==="mcf")throw Error("Select standard SMF / step-index model first.");
   for(const [id,value] of [["outer_radius_um",outer],["core_radius_um",inner],["wavelength_nm",wl],["delta_n_core",dn]]){
    $(id).value=unit(value);$(id).dispatchEvent(new Event("input",{bubbles:true}));
   }
   if($("ls-cad-fix-nbg").checked){$("clad_material").value="fixed";$("clad_material").dispatchEvent(new Event("input",{bubbles:true}));$("clad_index_override").value=unit(bg);$("clad_index_override").dispatchEvent(new Event("input",{bubbles:true}));}
   lastAppliedSignature=physSignature();emitCadState();
   document.dispatchEvent(new CustomEvent("ls-cad-applied",{detail:{outerRadiusUm:outer,coreRadiusUm:inner}}));
   say("Applied to existing Radial FEM inputs. Domain 1 = Core, Domain 2 = Cladding. Run Compute.");return true;
  }catch(e){const message="Cannot apply: "+e.message;say(message,true);document.dispatchEvent(new CustomEvent("ls-cad-apply-failed",{detail:{message}}));return false;}
 }
 const menuAction=(label,cmd,props={})=>({label,cmd,...props});
 function menuItems(scope){
  if(scope==="root")return [
   menuAction("Build All","build-all",{shortcut:"F8"}),menuAction("Import","disabled",{disabled:true}),{sep:true},
   menuAction("Circle","circle"),...["Ellipse","Rectangle","Square","Circular Arc","Cubic Bézier","Interpolation Curve","Line Segment","Parametric Curve","Point","Polygon","Quadratic Bézier"].map(t=>menuAction(t,"disabled",{disabled:true})),
   {sep:true},menuAction("Transforms","submenu",{children:[menuAction("Rotate","rotate"),menuAction("Move","disabled",{disabled:true}),menuAction("Mirror","mirror")]}),
   ...["Parts","Programming","Booleans and Partitions","Conversions","Chamfer","Fillet","Measure","Insert Sequence"].map(t=>menuAction(t,"disabled",{disabled:true})),{sep:true},menuAction("Copy as Code to Clipboard","copy-code",{disabled:true}),menuAction("Rename","rename",{disabled:true}),menuAction("Settings","settings"),menuAction("Properties","properties"),menuAction("Help","help")
  ];
  return [
   menuAction("Build Selected","build-selected",{shortcut:"F7"}),menuAction("Build Preceding","build-preceding",{shortcut:"F6"}),
   menuAction("Add Before","submenu",{children:[menuAction("Circle","circle",{pos:"before"}),menuAction("Rotate","rotate",{pos:"before"}),menuAction("Mirror","mirror",{pos:"before"})]}),
   menuAction("Add After","submenu",{children:[menuAction("Circle","circle",{pos:"after"}),menuAction("Rotate","rotate",{pos:"after"}),menuAction("Mirror","mirror",{pos:"after"})]}),
   menuAction("Copy as Code to Clipboard","copy-code"),{sep:true},
   menuAction("Move Up","up",{shortcut:"Ctrl+Up"}),menuAction("Move Down","down",{shortcut:"Ctrl+Down"}),menuAction("Copy","copy"),menuAction("Paste","paste",{disabled:!clipboard,shortcut:"Ctrl+V"}),menuAction("Duplicate","duplicate",{shortcut:"Ctrl+Shift+D"}),
   menuAction("Delete","delete",{shortcut:"Del"}),menuAction(cur()?.enabled?"Disable":"Enable","toggle",{shortcut:"F3"}),menuAction("Rename","rename",{shortcut:"F2"}),
   {sep:true},menuAction("Settings","settings"),menuAction("Properties","properties"),menuAction("Help","help")
  ];
 }
 function menuHTML(items){
  return `<ul>${items.map(i=>i.sep?'<li class="ls-cad-menu-sep"></li>':
  `<li class="${i.disabled?"disabled":""} ${i.children?"has-children":""}"><button type="button" ${i.disabled?"disabled":""} data-cmd="${esc(i.cmd)}" ${i.pos?`data-pos="${i.pos}"`:""}>${esc(i.label)}${i.shortcut?`<kbd>${esc(i.shortcut)}</kbd>`:""}${i.children?'<span class="arrow">▸</span>':""}</button>${i.children?menuHTML(i.children):""}</li>`).join("")}</ul>`;
 }
 function showMenu(event,scope,id){
  event.preventDefault();event.stopPropagation();
  if(scope==="root"){selected="root";if(!pane.hidden)loadFields();}
  else if(id&&nodes.some(n=>n.id===id)){selected=id;if(!pane.hidden)loadFields();}
  context.innerHTML=menuHTML(menuItems(scope));context.hidden=false;
  const w=300,h=context.getBoundingClientRect().height;
  context.style.left=Math.max(6,Math.min(event.clientX,window.innerWidth-w-8))+"px";
  context.style.top=Math.max(6,Math.min(event.clientY,window.innerHeight-h-8))+"px";
 }
 document.querySelector(".ls-model-tree")?.addEventListener("contextmenu",e=>{
  const n=e.target.closest("[data-cad-node]"),g=e.target.closest('summary[data-action="geometry"]');
  if(n)showMenu(e,"node",n.dataset.cadNode);else if(g)showMenu(e,"root");
 });
 context.addEventListener("click",e=>{
  const b=e.target.closest("button[data-cmd]");if(!b||b.disabled||b.parentElement?.classList.contains("has-children"))return;
  context.hidden=true;modify(b.dataset.cmd,b.dataset.pos);
 });
 document.addEventListener("pointerdown",e=>{if(!context.hidden&&!context.contains(e.target))context.hidden=true;});
 document.addEventListener("ls-open-wave-optics",()=>{if(!pane.hidden)close();},true);
 document.addEventListener("ls-open-material-editor",()=>{
  if(!pane.hidden)close();
  graphics.hidden=false;materialPickMode=false;materialHighlightKeys.clear();materialHighlightDomainIds.clear();materialEntityLevel="Domain";
  draw();
 });
 document.addEventListener("ls-material-picking",e=>{
  materialPickMode=!!e.detail?.enabled;
  materialEntityLevel=e.detail?.level||materialEntityLevel;
  materialHighlightKeys=new Set(e.detail?.entityKeys||[]);
  materialHighlightDomainIds=new Set(e.detail?.domainIds||[]);
  if(!materialPickMode)materialHighlightKeys.clear();
  graphics.hidden=!document.body.classList.contains("ls-material-active")&&pane.hidden;
  draw();
 });
 document.addEventListener("ls-material-highlight",e=>{
  materialEntityLevel=e.detail?.level||materialEntityLevel;
  materialHighlightKeys=new Set(e.detail?.entityKeys||[]);
  materialHighlightDomainIds=new Set(e.detail?.domainIds||[]);
  if(!pane.hidden||!graphics.hidden)draw();
 });

 document.addEventListener("ls-adjacent-picker",e=>{
  adjacentView=!!e.detail?.visible;
  adjacentPickMode=!!e.detail?.enabled;
  adjacentLevel=e.detail?.level||"Boundary";
  adjacentSelectedKeys=new Set(e.detail?.selectedKeys||[]);
  if(adjacentPickMode){
   materialPickMode=false;selectingRotateInputs=false;
   graphics.hidden=false;
  }else if(document.body.classList.contains("ls-adjacent-active")) graphics.hidden=false;
  draw();
 });
 document.addEventListener("ls-adjacent-editor-activated",()=>{
  if(!pane.hidden)close();
  graphics.hidden=false;
  materialPickMode=false;materialHighlightKeys.clear();materialHighlightDomainIds.clear();
  draw();
 });
 document.addEventListener("ls-open-cad",e=>open(e.detail?.id));
 document.addEventListener("ls-cad-add",e=>add(e.detail?.kind==="rotate"?"rotate":"circle"));
 document.addEventListener("ls-open-settings",e=>{if(e.detail?.kind==="parameters"&&!pane.hidden)close();},true);
 $("ls-cad-close").addEventListener("click",close);
 $("ls-cad-root-label").addEventListener("input",e=>{checkpoint();rootLabel=e.target.value;const tag=$("ls-cad-root-name");if(tag)tag.textContent=rootLabel;draw();});
 $("ls-cad-root-unit").addEventListener("change",e=>{checkpoint();lengthUnit=e.target.value;draw();});
 $("ls-cad-angle-unit").addEventListener("change",e=>{angleUnit=e.target.value;});
 $("ls-cad-constraints").addEventListener("change",e=>{constraintMode=e.target.value;});
 $("ls-cad-repair").addEventListener("change",e=>{repairTolerance=e.target.value;});
 $("ls-cad-scale-unit").addEventListener("change",e=>{scaleWhenUnitsChange=e.target.checked;});
 $("ls-cad-build-new").addEventListener("change",e=>{buildOnNew=e.target.checked;});
 $("ls-cad-build-leave").addEventListener("change",e=>{buildOnLeave=e.target.checked;});
 document.addEventListener("ls-wizard-finished",e=>{
  if(e.detail?.physics!=="ewfd"||e.detail?.study!=="mode")return;
  checkpoint();nodes=[];cumulativeSelections.length=0;activeCumulativeId=null;selected="root";limit=null;zoom=1;panX=0;panY=0;viewExtent=null;rootLabel="Geometry 1";lengthUnit="m";
  const label=$("ls-cad-root-name");if(label)label.textContent=rootLabel;
  drawTree();draw();say("New 2D geometry created. Right-click Geometry 1 → Circle to add your first feature.");
 });
 $("ls-cad-undo").addEventListener("click",()=>modify("undo"));
 $("ls-cad-redo").addEventListener("click",()=>modify("redo"));
 document.addEventListener("ls-cad-command",e=>modify(e.detail?.command));
 $("ls-cad-build-selected").addEventListener("click",()=>modify("build-selected"));
 $("ls-cad-build-all").addEventListener("click",()=>modify("build-all"));
 $("ls-cad-add-circle").addEventListener("click",()=>add("circle"));
 $("ls-cad-add-rotate").addEventListener("click",()=>add("rotate"));
 $("ls-cad-add-mirror").addEventListener("click",()=>add("mirror"));
 $("ls-cad-reset").addEventListener("click",()=>{checkpoint();nodes=defaults();selected="c2";limit=null;zoom=1;viewExtent=null;loadFields();say("Geometry restored to original Circle 1 and 2.");});
 $("ls-cad-apply").addEventListener("click",apply);
 $("ls-cad-export").addEventListener("click",()=>{
  const doc={format:"LS Optical Geometry CAD v2",parameters:svc.rows(),geometry:nodes};
  const blob=new Blob([JSON.stringify(doc,null,2)],{type:"application/json"}),url=URL.createObjectURL(blob),a=document.createElement("a");
  a.href=url;a.download="LS_Circle_Rotate_Geometry.json";a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
 });
 for(const key of ["type","radius","sector","base","x","y","rotation","layer","angle","cx","cy"]){
  $("ls-cad-"+key).addEventListener("input",e=>{const n=cur();if(!n||!(key in n))return;checkpoint();n[key]=e.target.value;draw();});
 }
 for(const key of ["cx","cy","nx","ny"]){
  $("ls-cad-mirror-"+key).addEventListener("input",e=>{
   const n=cur();if(n?.kind!=="mirror")return;
   checkpoint();n[key]=e.target.value;draw();
  });
 }
 $("ls-cad-label").addEventListener("input",e=>{if(!cur())return;checkpoint();cur().label=e.target.value;drawTree();});

 $("ls-cad-add-layer").addEventListener("click",()=>{
  const n=cur();if(!n||n.kind!=="circle")return;
  if((n.layerExtra||[]).length>=15){say("Maximum 16 layers supported in preview.",true);return;}
  checkpoint();n.layerExtra??=[];n.layerExtra.push("0[um]");renderExtraLayers(n);draw();
 });
 $("ls-cad-remove-layer").addEventListener("click",()=>{const n=cur();if(!n?.layerExtra?.length)return;checkpoint();n.layerExtra.pop();renderExtraLayers(n);draw();});
 $("ls-cad-extra-layers").addEventListener("input",e=>{
  const i=Number(e.target.dataset.cadLayerExtra),n=cur();
  if(!n||!Number.isInteger(i)||i<0||i>=(n.layerExtra?.length||0))return;
  checkpoint();n.layerExtra[i]=e.target.value;draw();
 });
 document.addEventListener("ls-global-parameters-changed",()=>{
  if(!pane.hidden)draw();
  else emitCadState();
 });
 
 $("ls-cad-activate-selection").addEventListener("click",()=>{
  if(!["rotate","mirror"].includes(cur()?.kind))return;
  selectingRotateInputs=!selectingRotateInputs;
  loadFields();
  say(selectingRotateInputs?"Activate Selection: input circles are highlighted in blue. Click a shape to add/remove its input.":"Graphics input selection deactivated.");
 });
 $("ls-cad-range-open").addEventListener("click",showRangeDialog);
 $("ls-cad-range-cancel").addEventListener("click",()=>rangeDialog.close("cancel"));
 $("ls-cad-range-close").addEventListener("click",()=>rangeDialog.close("cancel"));
 $("ls-cad-range-form").addEventListener("submit",e=>{
  e.preventDefault();
  const error=$("ls-cad-range-error");error.hidden=true;error.textContent="";
  try{
   const node=cur();
   if(node?.kind!=="rotate")throw Error("Select a Rotate feature first.");
   if($("ls-cad-range-method").value!=="Step")throw Error("Only Step ranges are implemented.");
   const fields=["start","step","stop"].map(key=>$("ls-cad-range-"+key).value.trim());
   if(fields.some(x=>!x))throw Error("Start, Step and Stop are required.");
   const base="range("+fields.join(",")+")";
   const points=angleList(base);
   const nextRange=$("ls-cad-range-function").value==="negate"?points.map(x=>String(-x)).join(","):base;
   const mode=e.submitter?.value==="add"?"add":"replace";
   const next=mode==="add"&&node.angle.trim()?node.angle.trim()+","+nextRange:nextRange;
   const values=angleList(next);
   if(values.length>72)throw Error("At most 72 angles are supported.");
   checkpoint();
   node.angle=next;
   $("ls-cad-angle").value=next;
   draw();
   rangeDialog.close(mode);
   say("Rotate angle "+(mode==="add"?"added":"replaced")+": "+values.length+" position(s) • "+next);
  }catch(ex){
   error.textContent=ex.message;error.hidden=false;
   $("ls-cad-range-start").focus();
  }
 });


 $("ls-cad-contribute").addEventListener("change",e=>{
  const n=cur();if(!n)return;
  checkpoint();n.cumulativeId=e.target.value||null;
  if(n.cumulativeId)n.selection=true;
  for(const row of cumulativeSelections){
   row.featureIds=row.featureIds.filter(id=>id!==n.id);
   if(row.id===n.cumulativeId)row.featureIds.push(n.id);
  }
  loadFields();
  say(n.cumulativeId?"Contribute to "+cumulativeSelections.find(v=>v.id===n.cumulativeId)?.label:"Cumulative contribution removed.");
  document.dispatchEvent(new CustomEvent("ls-cad-cumulative-changed",{detail:{selections:window.LSWaveCADService.cumulativeSelections()}}));
 });
 $("ls-cad-cumulative-new").addEventListener("click",()=>{
  if(!cur())return;
  let number=1;
  while(cumulativeSelections.some(v=>v.label==="Cumulative Selection "+number))number++;
  $("ls-cad-cumulative-name").value="Cumulative Selection "+number;
  $("ls-cad-cumulative-error").hidden=true;
  $("ls-cad-cumulative-error").textContent="";
  cumulativeDialog.showModal();
  $("ls-cad-cumulative-name").focus();
 });
 cumulativeDialog.addEventListener("close",()=>{
  if(cumulativeDialog.returnValue!=="ok")return;
  const value=$("ls-cad-cumulative-name").value.trim(),node=cur();
  if(!value||!node)return;
  const newId="csel"+(cumulativeSelections.length+1);
  const existing=cumulativeSelections.find(v=>v.label.toLowerCase()===value.toLowerCase());
  if(existing){say("Cumulative selection already exists: "+value,true);return;}
  for(const row of cumulativeSelections)row.featureIds=row.featureIds.filter(id=>id!==node.id);
  cumulativeSelections.push({id:newId,label:value,featureIds:[node.id]});
  node.cumulativeId=newId;node.selection=true;
  loadFields();
  say("Created "+value+" under Geometry 1 → Cumulative Selections.");
  document.dispatchEvent(new CustomEvent("ls-cad-cumulative-created",{detail:{id:newId,label:value,featureId:node.id}}));
  document.dispatchEvent(new CustomEvent("ls-cad-cumulative-changed",{detail:{selections:window.LSWaveCADService.cumulativeSelections()}}));
 });
 cumulativeDialog.querySelector("form").addEventListener("submit",e=>{
  const name=$("ls-cad-cumulative-name").value.trim();
  if(e.submitter?.value!=="ok")return;
  if(!name){e.preventDefault();$("ls-cad-cumulative-error").hidden=false;$("ls-cad-cumulative-error").textContent="Enter a label";return;}
  if(cumulativeSelections.some(v=>v.label.toLowerCase()===name.toLowerCase())){e.preventDefault();$("ls-cad-cumulative-error").hidden=false;$("ls-cad-cumulative-error").textContent="Selection name already exists.";}
 });
 document.querySelector(".ls-model-tree")?.addEventListener("click",e=>{
  const selection=e.target.closest("[data-cad-cumulative-id]");
  if(selection)openCumulative(selection.dataset.cadCumulativeId);
 });

 $("ls-cad-keep").addEventListener("change",e=>{checkpoint();cur().keep=e.target.checked;draw();});
 $("ls-cad-inputs").addEventListener("change",e=>{
  const box=e.target.closest("[data-cad-input-ref]"),node=cur();
  if(!box||!node||!["rotate","mirror"].includes(node.kind))return;
  checkpoint();
  const ids=selectedInputKeys(node);
  if(box.checked)ids.add(box.dataset.cadInputRef);else ids.delete(box.dataset.cadInputRef);
  node.inputRefs=[...ids];
  node.inputIds=availableInputs(node).filter(s=>ids.has(s.entityKey)&&!s.derived).map(s=>s.source);
  draw();
 });
 $("ls-cad-selection").addEventListener("change",e=>{checkpoint();cur().selection=e.target.checked;$("ls-cad-physics").disabled=!e.target.checked;draw();});
 $("ls-cad-color").addEventListener("change",e=>{checkpoint();cur().color=e.target.value;draw();});
 $("ls-cad-plus").addEventListener("click",()=>{zoom=Math.min(zoom*1.25,8);draw();});
 $("ls-cad-minus").addEventListener("click",()=>{zoom=Math.max(zoom/1.25,.25);draw();});
 $("ls-cad-fit").addEventListener("click",()=>{try{viewExtent=visibleExtent();zoom=1;panX=0;panY=0;draw();say("Fit view: physical scale updated.");}catch(e){say("Fit failed: "+e.message,true);}});
 $("ls-cad-grid").addEventListener("click",()=>{if(adjacentView)adjacentGridOn=!adjacentGridOn;else gridOn=!gridOn;draw();});
 $("ls-cad-view").addEventListener("click",e=>{
  if(adjacentPickMode){
   const target=e.target.closest("[data-adjacent-key]");
   // Prefer the nearest actual ring/line to avoid a coincident SVG hit target
   // or a transparent fill capturing the pointer over a different circle.
   const picked=adjacentLevel==="Domain"?materialDomainAtScreen(e):nearestAdjacentAtScreen(e);
   const key=picked?.key||target?.dataset.adjacentKey;
   if(key)document.dispatchEvent(new CustomEvent("ls-adjacent-geometry-click",{detail:{
    key,id:picked?.id??Number(target?.dataset.adjacentEntityId),level:adjacentLevel
   }}));
   return;
  }
  if(adjacentView)return;
  if(materialPickMode&&materialEntityLevel==="Domain"){
   const domain=materialDomainAtScreen(e);
   if(domain)document.dispatchEvent(new CustomEvent("ls-material-geometry-click",{detail:{
    level:"Domain",domainId:domain.id,domainKey:domain.key,owner:domain.owner
   }}));
   // A Domain click must never toggle an entire Geometry Shape.
   return;
  }
  const shape=e.target.closest("[data-cad-shape]");
  if(!shape)return;
  const node=cur();
  if(materialPickMode){
   document.dispatchEvent(new CustomEvent("ls-material-geometry-click",{detail:{entityKey:shape.dataset.cadEntity,owner:shape.dataset.cadShape,level:materialEntityLevel}}));
   return;
  }
  if(selectingRotateInputs&&["rotate","mirror"].includes(node?.kind)){
   const key=shape.dataset.cadEntity;
   const available=availableInputs(node);
   if(!available.some(s=>s.entityKey===key)){say("Select an individual preceding Geometry object (c2/rot1 output).",true);return;}
   checkpoint();
   const selectedRefs=selectedInputKeys(node,available);
   selectedRefs.has(key)?selectedRefs.delete(key):selectedRefs.add(key);
   node.inputRefs=[...selectedRefs];
   node.inputIds=available.filter(s=>selectedRefs.has(s.entityKey)&&!s.derived).map(s=>s.source);
   loadFields();
   say(key+(selectedRefs.has(key)?" selected as ":" removed from ")+node.label+" input objects.");
   return;
  }
  selected=shape.dataset.cadShape;loadFields();
 });
 $("ls-cad-object-select").addEventListener("change",e=>open(e.target.value));
 const viewport=$("ls-cad-view");let gesture=null,hovered=null;
 viewport.addEventListener("pointermove",e=>{
  let next=null;
  if(adjacentView){
   if(adjacentPickMode){
    const entity=adjacentLevel==="Domain"?materialDomainAtScreen(e):nearestAdjacentAtScreen(e);
    if(entity)next=[...viewport.querySelectorAll("[data-adjacent-key]")].find(node=>node.dataset.adjacentKey===entity.key)||null;
   }
  }else next=e.target.closest?.('g[data-cad-shape]')||null;
  if(next===hovered)return;
  hovered?.classList.remove("is-hovered");
  hovered=next;
  hovered?.classList.add("is-hovered");
 });
 viewport.addEventListener("pointerleave",()=>{hovered?.classList.remove("is-hovered");hovered=null;});
 viewport.addEventListener("pointerdown",e=>{
  if(e.button!==0||selectingRotateInputs||materialPickMode||e.target.closest("[data-cad-shape]"))return;
  if(adjacentView&&adjacentPickMode&&(adjacentLevel==="Domain"?materialDomainAtScreen(e):nearestAdjacentAtScreen(e)))return;
  gesture={id:e.pointerId,x:e.clientX,y:e.clientY,panX,panY};
  viewport.setPointerCapture(e.pointerId);viewport.classList.add("is-panning");
 });
 viewport.addEventListener("pointermove",e=>{
  if(!gesture||e.pointerId!==gesture.id)return;
  const b=viewport.getBoundingClientRect();
  if(!b.width||!b.height)return;
  panX=gesture.panX+(e.clientX-gesture.x)*640/b.width;
  panY=gesture.panY+(e.clientY-gesture.y)*500/b.height;
  draw();
 });
 const stop=e=>{if(gesture?.id===e.pointerId){gesture=null;viewport.classList.remove("is-panning");}};
 viewport.addEventListener("pointerup",stop);viewport.addEventListener("pointercancel",stop);
 viewport.addEventListener("wheel",e=>{e.preventDefault();zoom=Math.max(.25,Math.min(8,zoom*(e.deltaY<0?1.15:1/1.15)));draw();},{passive:false});
 $("ls-cad-view").addEventListener("contextmenu",e=>{const s=e.target.closest("[data-cad-shape]");if(s)showMenu(e,"node",s.dataset.cadShape);});
 document.addEventListener("keydown",e=>{
  if(rangeDialog.open)return;
   if(e.key==="Escape"){if(!context.hidden){context.hidden=true;e.stopImmediatePropagation();}else if(!pane.hidden)close();}
  if(pane.hidden||["INPUT","TEXTAREA","SELECT"].includes(document.activeElement?.tagName))return;
   if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==="z"){e.preventDefault();modify(e.shiftKey?"redo":"undo");return;}
   if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==="y"){e.preventDefault();modify("redo");return;}
   if((e.ctrlKey||e.metaKey)&&e.shiftKey&&e.key.toLowerCase()==="d"){e.preventDefault();modify("duplicate");return;}
   if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==="c"){e.preventDefault();modify("copy");return;}
   if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==="v"){e.preventDefault();modify("paste");return;}
  const actions={F6:"build-preceding",F7:"build-selected",F8:"build-all",F2:"rename",F3:"toggle",Delete:"delete"};
  if(actions[e.key]){e.preventDefault();modify(actions[e.key]);}
 });
 $("preset")?.addEventListener("change",()=>{nodes=defaults();selected="c2";limit=null;zoom=1;panX=0;panY=0;viewExtent=null;undoStack.length=0;redoStack.length=0;lastAppliedSignature=physSignature();updateHistoryButtons();drawTree();emitCadState();if(!pane.hidden)close();});
 $("resetBtn")?.addEventListener("click",()=>queueMicrotask(()=>{nodes=defaults();selected="c2";limit=null;zoom=1;panX=0;panY=0;viewExtent=null;undoStack.length=0;redoStack.length=0;lastAppliedSignature=physSignature();updateHistoryButtons();drawTree();emitCadState();if(!pane.hidden)close();}));
 drawTree();updateHistoryButtons();
})();
