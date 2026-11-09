(function geometryCad(){
 "use strict";
 const $=id=>document.getElementById(id),side=document.querySelector(".layout > .sidebar"),main=document.querySelector("main.workspace"),svc=window.LSWaveParameterService;
 if(!side||!main||!svc||$("ls-cad-pane"))return;
 const esc=s=>String(s??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
 const unit=n=>Number(n.toPrecision(10)).toString();
 const circle=(id,label,radius)=>({id,kind:"circle",label,enabled:true,type:"Solid",radius,sector:"360",base:"Center",x:"0[um]",y:"0[um]",rotation:"0",layer:"0[um]",selection:false,color:"None"});
 const defaults=()=>[circle("c1","Circle 1","rb"),circle("c2","Circle 2","d/2")];
 let nodes=defaults(),selected="c2",limit=null,zoom=1,gridOn=true,clipboard=null,rootLabel="Geometry 1",lengthUnit="um",panX=0,panY=0;
 let angleUnit="Degrees",constraintMode="Off",repairTolerance="Automatic",buildOnNew=true,buildOnLeave=true,scaleWhenUnitsChange=false;
 const undoStack=[],redoStack=[];
 function snap(){return {nodes:structuredClone(nodes),selected,limit,zoom,panX,panY,rootLabel,lengthUnit,gridOn};}
 function updateHistoryButtons(){
  if($("ls-cad-undo"))$("ls-cad-undo").disabled=!undoStack.length;
  if($("ls-cad-redo"))$("ls-cad-redo").disabled=!redoStack.length;
  document.dispatchEvent(new CustomEvent("ls-cad-history",{detail:{canUndo:!!undoStack.length,canRedo:!!redoStack.length}}));
 }
 function checkpoint(){undoStack.push(snap());if(undoStack.length>60)undoStack.shift();redoStack.length=0;updateHistoryButtons();}
 function restore(state){({nodes,selected,limit,zoom,panX,panY,rootLabel,lengthUnit,gridOn}=structuredClone(state));const node=$("ls-cad-root-name");if(node)node.textContent=rootLabel;loadFields();updateHistoryButtons();}
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
   <details class="ls-cad-section" open><summary>Layers</summary><div class="ls-cad-layer"><span>Layer name</span><span>Thickness (m)</span><label for="ls-cad-layer">Layer 1</label><input id="ls-cad-layer"/></div></details>
  </div>
  <div id="ls-cad-rotate" hidden>
   <details class="ls-cad-section" open><summary>Input</summary><label class="ls-cad-block-label">Input objects</label><div id="ls-cad-inputs" class="ls-cad-inputs"></div><label class="ls-cad-check"><input id="ls-cad-keep" type="checkbox" checked/> Keep input objects</label></details>
   <details class="ls-cad-section" open><summary>Rotation</summary><div class="ls-cad-field"><label for="ls-cad-angle">Angle</label><input id="ls-cad-angle"/><span>deg</span></div><small>range(0,60,300) generates 0°, 60°, ..., 300°</small></details>
   <details class="ls-cad-section" open><summary>Center of Rotation</summary><div class="ls-cad-field"><label for="ls-cad-cx">x</label><input id="ls-cad-cx"/><span>m</span></div><div class="ls-cad-field"><label for="ls-cad-cy">y</label><input id="ls-cad-cy"/><span>m</span></div></details>
  </div>
  <details class="ls-cad-section" open><summary>Selections of Resulting Entities</summary>
   <label class="ls-cad-check"><input id="ls-cad-selection" type="checkbox"/> Resulting objects selection</label>
   <div class="ls-cad-field"><label for="ls-cad-physics">Show in physics</label><select id="ls-cad-physics" disabled><option>Domain selection</option><option>Boundary selection</option></select><span></span></div>
   <div class="ls-cad-field"><label for="ls-cad-color">Color</label><select id="ls-cad-color"><option>None</option><option>Blue</option><option>Orange</option><option>Green</option></select><span></span></div>
   <div class="ls-cad-field"><label>Contribute to</label><select disabled><option>None</option></select><span></span></div></details>
  <div class="ls-cad-mini"><svg id="ls-cad-mini-svg" viewBox="0 0 460 230" role="img" aria-label="Small geometry preview"></svg></div>
  </div>
  <div class="ls-config-actions"><button type="button" id="ls-cad-add-circle">＋ Circle</button><button type="button" id="ls-cad-add-rotate">↻ Rotate</button><button type="button" id="ls-cad-reset">Reset</button><button type="button" id="ls-cad-export">Export JSON</button></div>
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
 const cur=()=>nodes.find(n=>n.id===selected),index=id=>nodes.findIndex(n=>n.id===id);
 const physSignature=()=>{
  const structural=nodes.map(n=>n.kind==="circle"?[n.id,n.kind,n.enabled,n.type,n.radius,n.sector,n.x,n.y,n.layer]:[n.id,n.kind,n.enabled,n.inputIds,n.keep,n.angle,n.cx,n.cy]);
  // Resolve parameter expressions as well as literal expressions. An edit of d/rb
  // changes the geometry even if Circle 1/2 still display the same expressions.
  const evaluated=nodes.filter(n=>n.kind==="circle").map(n=>{
   try{return [n.id,...[n.radius,n.x,n.y,n.sector,n.layer].map(x=>svc.calculate(x))];}
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
    if(n.type!=="Solid"||Math.abs(evalNumber(n.sector)-360)>1e-9||evalNumber(n.layer)>1e-12)throw Error(n.label+" must be a 360° Solid with no layers.");
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
 window.LSWaveCADService=Object.freeze({status,apply:()=>apply()});
 const evalNumber=s=>svc.calculate(s);
 const say=(msg,error=false)=>{const e=$("ls-cad-status");e.textContent=msg;e.classList.toggle("error",error);};
 function fresh(kind,at){
  let i=1,pre=kind==="circle"?"c":"rot";while(nodes.some(n=>n.id===pre+i))i++;
  if(kind==="circle")return circle(pre+i,"Circle "+i,"1[m]");
  const preceding=nodes.slice(0,at).filter(x=>x.kind==="circle"&&x.enabled);
  return {id:pre+i,kind:"rotate",label:"Rotate "+i,enabled:true,inputIds:preceding.length?[preceding.at(-1).id]:[],keep:true,angle:"range(0,60,300)",cx:"0[um]",cy:"0[um]",selection:false,color:"None"};
 }
 function drawTree(){
  const list=$("ls-geometry-node-list");if(!list)return;
  list.innerHTML=nodes.map(n=>`<button type="button" data-cad-node="${esc(n.id)}" class="ls-tree-leaf ls-cad-node ${n.id===selected?"active":""} ${n.enabled?"":"ls-cad-node-disabled"}" title="Click: settings · right-click: context menu"><span class="ls-cad-symbol">${n.kind==="circle"?"◉":"↻"}</span><span>${esc(n.label)} <small>(${esc(n.id)})</small></span></button>`).join("");
 }
 function loadFields(){
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
  $("ls-cad-subtitle").textContent=n.kind==="circle"?"Circle":"Rotate";$("ls-cad-label").value=n.label;
  $("ls-cad-circle").hidden=n.kind!=="circle";$("ls-cad-rotate").hidden=n.kind!=="rotate";
  if(n.kind==="circle"){
   for(const k of ["type","radius","sector","base","x","y","rotation","layer"])$("ls-cad-"+k).value=n[k];
  }else{
   $("ls-cad-angle").value=n.angle;$("ls-cad-cx").value=n.cx;$("ls-cad-cy").value=n.cy;$("ls-cad-keep").checked=n.keep;
   const list=$("ls-cad-inputs");list.replaceChildren();
   const options=nodes.slice(0,index(n.id)).filter(o=>o.kind==="circle"&&o.enabled);
   if(!options.length)list.textContent="Move Rotate below an enabled Circle.";
   for(const input of options){
    const label=document.createElement("label"),cb=document.createElement("input");
    cb.type="checkbox";cb.dataset.cadInput=input.id;cb.checked=n.inputIds.includes(input.id);
    label.append(cb,document.createTextNode(` ${input.label} (${input.id})`));list.append(label);
   }
  }
  $("ls-cad-selection").checked=n.selection;$("ls-cad-physics").disabled=!n.selection;$("ls-cad-color").value=n.color;
  drawTree();draw();
 }
 function angleList(s){
  const text=String(s).trim();
  if(!/^range\s*\(/i.test(text))return [evalNumber(text)];
  const match=/^range\s*\((.*)\)$/i.exec(text);
  if(!match)throw Error("Use range(start,step,end)");
  const p=match[1].split(",").map(v=>v.trim());if(p.length!==3)throw Error("range requires start, step, end");
  const [a,b,c]=p.map(evalNumber);
  if(!b||(c-a)*b<0)throw Error("Invalid Rotate range direction or step");
  const length=Math.floor((c-a)/b+1e-10)+1;
  if(length<1||length>72)throw Error("Rotate range must have 1–72 angles");
  return Array.from({length},(_,i)=>a+i*b);
 }
 function compute(){
  const p=svc.evaluate();if(p.errors.size)throw Error([...p.errors.values()][0]);
  const shapes=[],original=new Map();
  const max=limit===null?nodes.length-1:Math.min(limit,nodes.length-1);
  for(let i=0;i<=max;i++){
   const n=nodes[i];if(!n.enabled)continue;
   if(n.kind==="circle"){
    const r=evalNumber(n.radius),x=evalNumber(n.x),y=evalNumber(n.y),sector=evalNumber(n.sector),deg=evalNumber(n.rotation),thick=evalNumber(n.layer);
    if(!(r>0&&r<=10000)||!(sector>0&&sector<=360)||thick<0)throw Error(n.label+": invalid radius, sector or layer");
    const s={id:n.id,source:n.id,owner:n.id,r,x,y,sector,deg,solid:n.type==="Solid",selection:n.selection,color:n.color,derived:false};
    shapes.push(s);original.set(n.id,s);
   }else{
    if(!n.inputIds.length)throw Error(n.label+": select input circles");
    const input=n.inputIds.map(id=>original.get(id));
    if(input.some(x=>!x))throw Error(n.label+": input must be a preceding enabled Circle");
    const cx=evalNumber(n.cx),cy=evalNumber(n.cy),angles=angleList(n.angle);
    if(!n.keep)for(let j=shapes.length-1;j>=0;j--)if(n.inputIds.includes(shapes[j].source)&&!shapes[j].derived)shapes.splice(j,1);
    for(const s of input)for(const deg of angles){
     const a=deg*Math.PI/180,dx=s.x-cx,dy=s.y-cy;
     shapes.push({...s,id:n.id,owner:n.id,x:cx+dx*Math.cos(a)-dy*Math.sin(a),y:cy+dx*Math.sin(a)+dy*Math.cos(a),deg:s.deg+deg,selection:n.selection,color:n.color,derived:true});
    }
   }
   if(shapes.length>450)throw Error("Too many objects for preview (max 450)");
  }
  return shapes;
 }
 function drawShape(x,y,r,sector,deg){
  if(sector>=359.999)return `<circle cx="${x}" cy="${y}" r="${r}"/>`;
  const a=(deg-90)*Math.PI/180,b=a+sector*Math.PI/180,x1=x+r*Math.cos(a),y1=y+r*Math.sin(a),x2=x+r*Math.cos(b),y2=y+r*Math.sin(b);
  return `<path d="M${x} ${y} L${x1} ${y1} A${r} ${r} 0 ${sector>180?1:0} 1 ${x2} ${y2} Z"/>`;
 }
 function svg(w,h,mini=false){
  const shapes=compute();
  const extent=Math.max(1e-6,...shapes.map(s=>s.r+Math.hypot(s.x,s.y)));
  const scale=Math.min(w,h)*(mini?.42:.41)*zoom/extent;
  let grid="";
  if(gridOn){for(let x=0;x<w;x+=30)grid+=`<path d="M ${x} 0 V ${h}" stroke="#e6ecf1" stroke-width=".8"/>`;for(let y=0;y<h;y+=30)grid+=`<path d="M 0 ${y} H ${w}" stroke="#e6ecf1" stroke-width=".8"/>`;}
  const items=shapes.map(s=>{
   const x=w/2+(mini?0:panX)+s.x*scale,y=h/2+(mini?0:panY)-s.y*scale,r=s.r*scale;
   const focus=selected===s.owner,stroke=focus?"#1f77ae":s.derived?"#7a8793":"#3f505f";
   const colors={Blue:"#9cc8ef",Orange:"#efb584",Green:"#a5d7b1"};
   const fill=s.solid?(s.selection&&colors[s.color]?colors[s.color]:(focus?"#c1d9eb":"#d0d3d7")):"none";
   const label=mini?"":`<text x="${x}" y="${y-r-7}" font-size="11" text-anchor="middle" fill="${stroke}">${esc(s.owner)}${s.derived?"*":""}</text>`;
   return `<g data-cad-shape="${esc(s.owner)}" fill="${fill}" fill-opacity="${s.derived?".43":".65"}" stroke="${stroke}" stroke-width="${focus?2.2:1.5}" ${s.derived?'stroke-dasharray="5 3"':""}>${drawShape(x,y,r,s.sector,s.deg)}</g>${label}`;
  }).join("");
  const ox=w/2+(mini?0:panX),oy=h/2+(mini?0:panY);
  return `<rect width="${w}" height="${h}" fill="#fcfdfe"/>${grid}<path d="M0 ${oy} H${w} M${ox} 0 V${h}" stroke="#82939f" stroke-width="1.2"/>${items}<text x="12" y="21" font-size="11" fill="#5d7f94">${esc(rootLabel)} · 2D</text>`;
 }
 function draw(){
  try{
   $("ls-cad-view").innerHTML=svg(640,500);
   $("ls-cad-mini-svg").innerHTML=svg(460,230,true);
    const pick=$("ls-cad-object-select");
    if(pick){pick.innerHTML='<option value="root">Geometry 1 · All</option>'+nodes.map(n=>`<option value="${esc(n.id)}">${esc(n.label)} (${esc(n.id)})</option>`).join("");pick.value=nodes.some(n=>n.id===selected)?selected:"root";}
   $("ls-cad-build-label").textContent=limit===null?"Build All":limit<0?"Build preceding (empty)":"Build through "+nodes[limit]?.label;
   $("ls-cad-legend").textContent=`${compute().length} object(s) in preview · Zoom ${Math.round(zoom*100)}% · ${lengthUnit==="um"?"µm":lengthUnit} · Not FEM field data`;emitCadState();
  }catch(e){$("ls-cad-view").innerHTML="";$("ls-cad-mini-svg").innerHTML="";$("ls-cad-legend").textContent=e.message;if(!pane.hidden)say("Preview: "+e.message,true);}
 }
 function open(id){
  if(id==="root")selected="root";else if(id&&nodes.some(n=>n.id===id))selected=id;
  if(selected!=="root"&&!cur())selected=nodes.at(-1)?.id;
  $("ls-config-pane").hidden=true;
  pane.hidden=false;graphics.hidden=false;
  side.classList.add("ls-config-active");document.body.classList.add("ls-configuration-open","ls-cad-active");
  loadFields();say("Edit Geometry features; 2D Graphics updates immediately. Radial FEM only changes after Apply.");
 }
 function hideContext(){context.hidden=true;}
 function close(){
  pane.hidden=true;graphics.hidden=true;side.classList.remove("ls-config-active");document.body.classList.remove("ls-configuration-open","ls-cad-active");hideContext();
 }
 function add(kind,pos=nodes.length){
  checkpoint();const n=fresh(kind,pos);nodes.splice(pos,0,n);selected=n.id;limit=null;open();say(`${n.label} added. Configure and Build All.`);
 }
 function modify(cmd,arg){
  const idx=index(selected),n=cur();
   if(cmd==="undo"||cmd==="redo"){historyMove(cmd==="undo");return;}
  switch(cmd){
   case "circle":case "rotate":add(cmd,arg==="before"?Math.max(idx,0):arg==="after"?idx+1:nodes.length);return;
   case "build-all":checkpoint();limit=null;draw();say("Built all Geometry features in the 2D Graphics preview.");return;
   case "build-selected":if(!n)return;checkpoint();limit=idx;draw();say("Build Selected finished (preview).");return;
   case "build-preceding":if(!n)return;checkpoint();limit=idx-1;draw();say("Preceding Geometry features built (preview).");return;
   case "up":case "down":{const i=idx+(cmd==="up"?-1:1);if(i<0||i>=nodes.length)return;checkpoint();[nodes[idx],nodes[i]]=[nodes[i],nodes[idx]];limit=null;loadFields();say("Geometry order updated.");return;}
   case "duplicate":{if(!n)return;checkpoint();const copy=structuredClone(n),newid=fresh(n.kind,idx+1);copy.id=newid.id;copy.label=newid.label;nodes.splice(idx+1,0,copy);selected=copy.id;limit=null;loadFields();say(`Duplicated as ${copy.label}.`);return;}
   case "delete":{if(!n)return;checkpoint();nodes.splice(idx,1);if(!nodes.length)nodes=defaults();selected=nodes[Math.min(idx,nodes.length-1)].id;limit=null;loadFields();return;}
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
    const x=evalNumber(c.x),y=evalNumber(c.y),a=evalNumber(c.sector),t=evalNumber(c.layer),radius=evalNumber(c.radius);
    if(c.type!=="Solid"||a!==360)throw Error("Radial FEM requires 360° solid circles.");
    if(Math.abs(x)>1e-12||Math.abs(y)>1e-12)throw Error("Radial FEM requires concentric centers at (0,0).");
    if(t>1e-12)throw Error("Extra layers are not supported by two-circle Radial FEM.");
    return radius*1e6;
   }
   const outer=r(c1),inner=r(c2),wl=evalNumber("wl")*1e9,dn=evalNumber("dn_core");
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
   {sep:true},menuAction("Transforms","submenu",{children:[menuAction("Rotate","rotate"),menuAction("Move","disabled",{disabled:true}),menuAction("Mirror","disabled",{disabled:true})]}),
   ...["Parts","Programming","Booleans and Partitions","Conversions","Chamfer","Fillet","Measure","Insert Sequence"].map(t=>menuAction(t,"disabled",{disabled:true})),{sep:true},menuAction("Copy as Code to Clipboard","copy-code",{disabled:true}),menuAction("Rename","rename",{disabled:true}),menuAction("Settings","settings"),menuAction("Properties","properties"),menuAction("Help","help")
  ];
  return [
   menuAction("Build Selected","build-selected",{shortcut:"F7"}),menuAction("Build Preceding","build-preceding",{shortcut:"F6"}),
   menuAction("Add Before","submenu",{children:[menuAction("Circle","circle",{pos:"before"}),menuAction("Rotate","rotate",{pos:"before"})]}),
   menuAction("Add After","submenu",{children:[menuAction("Circle","circle",{pos:"after"}),menuAction("Rotate","rotate",{pos:"after"})]}),
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
  if(id&&nodes.some(n=>n.id===id)){selected=id;if(!pane.hidden)loadFields();}
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
  checkpoint();nodes=[];selected="root";limit=null;zoom=1;panX=0;panY=0;rootLabel="Geometry 1";lengthUnit="m";
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
 $("ls-cad-reset").addEventListener("click",()=>{checkpoint();nodes=defaults();selected="c2";limit=null;zoom=1;loadFields();say("Geometry restored to original Circle 1 and 2.");});
 $("ls-cad-apply").addEventListener("click",apply);
 $("ls-cad-export").addEventListener("click",()=>{
  const doc={format:"LS Optical Geometry CAD v2",parameters:svc.rows(),geometry:nodes};
  const blob=new Blob([JSON.stringify(doc,null,2)],{type:"application/json"}),url=URL.createObjectURL(blob),a=document.createElement("a");
  a.href=url;a.download="LS_Circle_Rotate_Geometry.json";a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
 });
 for(const key of ["type","radius","sector","base","x","y","rotation","layer","angle","cx","cy"]){
  $("ls-cad-"+key).addEventListener("input",e=>{const n=cur();if(!n||!(key in n))return;checkpoint();n[key]=e.target.value;draw();});
 }
 $("ls-cad-label").addEventListener("input",e=>{if(!cur())return;checkpoint();cur().label=e.target.value;drawTree();});
 $("ls-cad-keep").addEventListener("change",e=>{checkpoint();cur().keep=e.target.checked;draw();});
 $("ls-cad-inputs").addEventListener("change",e=>{
  const b=e.target.closest("[data-cad-input]");if(!b)return;
  checkpoint();const ids=new Set(cur().inputIds);b.checked?ids.add(b.dataset.cadInput):ids.delete(b.dataset.cadInput);
  cur().inputIds=[...ids];draw();
 });
 $("ls-cad-selection").addEventListener("change",e=>{checkpoint();cur().selection=e.target.checked;$("ls-cad-physics").disabled=!e.target.checked;draw();});
 $("ls-cad-color").addEventListener("change",e=>{checkpoint();cur().color=e.target.value;draw();});
 $("ls-cad-plus").addEventListener("click",()=>{zoom=Math.min(zoom*1.25,8);draw();});
 $("ls-cad-minus").addEventListener("click",()=>{zoom=Math.max(zoom/1.25,.25);draw();});
 $("ls-cad-fit").addEventListener("click",()=>{zoom=1;panX=0;panY=0;draw();});
 $("ls-cad-grid").addEventListener("click",()=>{gridOn=!gridOn;draw();});
 $("ls-cad-view").addEventListener("click",e=>{const s=e.target.closest("[data-cad-shape]");if(s){selected=s.dataset.cadShape;loadFields();}});
 $("ls-cad-object-select").addEventListener("change",e=>open(e.target.value));
 const viewport=$("ls-cad-view");let gesture=null;
 viewport.addEventListener("pointerdown",e=>{
  if(e.button!==0||e.target.closest("[data-cad-shape]"))return;
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
 $("preset")?.addEventListener("change",()=>{nodes=defaults();selected="c2";limit=null;zoom=1;panX=0;panY=0;undoStack.length=0;redoStack.length=0;lastAppliedSignature=physSignature();updateHistoryButtons();drawTree();emitCadState();if(!pane.hidden)close();});
 $("resetBtn")?.addEventListener("click",()=>queueMicrotask(()=>{nodes=defaults();selected="c2";limit=null;zoom=1;panX=0;panY=0;undoStack.length=0;redoStack.length=0;lastAppliedSignature=physSignature();updateHistoryButtons();drawTree();emitCadState();if(!pane.hidden)close();}));
 drawTree();updateHistoryButtons();
})();
