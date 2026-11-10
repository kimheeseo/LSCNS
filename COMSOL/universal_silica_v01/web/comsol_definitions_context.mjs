/**
 * Component 1 > Definitions right-click menu.
 * COMSOL-like surface: only implemented Adjacent is actionable.
 * Unimplemented COMSOL features remain visible and disabled.
 */
(function initDefinitionsContext(){
 "use strict";
 if(document.getElementById("ls-definitions-context"))return;
 const tree=document.querySelector(".ls-model-tree");
 if(!tree)return;
 const groups=[
  [{label:"Variables",icon:"a≔"},{label:"Functions",children:["Analytic","Interpolation","Piecewise","More Functions"]},
   {label:"Selections",children:["Explicit","Complement","Adjacent","Ball/Disk","Box","Cylinder","Union","Intersection","Difference","Colors"]},
   {label:"Probes",children:["Domain Probe","Boundary Probe","Point Probe"]},
   {label:"Nonlocal Couplings",children:["Integration","Average","Maximum","Minimum"]},
   {label:"Pairs",children:["Identity Pair","Contact Pair"]},
   {label:"Coordinate Systems",children:["Boundary System","Cylindrical System"]}],
  [{label:"Moving Mesh",children:["Moving Mesh"]},{label:"Deformed Geometry",children:["Deformed Geometry"]},
   {label:"Topology Optimization",children:["Topology Optimization"]},{label:"Shape Optimization",children:["Shape Optimization"]}],
  [{label:"Perfectly Matched Layer",icon:"▥"},{label:"Shared Properties",children:["View","Selections"]},
   {label:"View",icon:"▧"}],
  [{label:"Show More Options..."},{label:"Node Group"},{label:"Group by Type",children:["Default","Name","Physics"]},
   {label:"Update Probes"}],
  [{label:"Help",hint:"F1"}]
 ];
 const safe=s=>String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
 const menu=document.createElement("div");
 menu.id="ls-definitions-context";menu.hidden=true;menu.setAttribute("role","menu");
 menu.setAttribute("aria-label","Definitions context menu");
 let count=0;
 menu.innerHTML=groups.map(group=>'<div class="ls-def-context-group">'+group.map(item=>{
  const hasChildren=!!item.children,identifier="ls-def-context-"+count++;
  if(!hasChildren){
   const isPML=item.label==="Perfectly Matched Layer";
   return '<div class="ls-def-context-item">'+
    '<button type="button" role="menuitem"'+(isPML?' data-definitions-command="pml" title="Add Perfectly Matched Layer"':' disabled aria-disabled="true" title="'+safe(item.label)+' — not implemented"')+'>'+
    '<span class="ls-def-context-glyph">'+safe(item.icon||"▧")+'</span><span>'+safe(item.label)+'</span>'+
    (item.hint?'<kbd>'+safe(item.hint)+'</kbd>':'')+'</button></div>';
  }
  return '<div class="ls-def-context-item ls-def-context-submenu" data-def-context-parent="'+safe(item.label)+'">'+
   '<button type="button" role="menuitem" aria-haspopup="true" aria-expanded="false" aria-controls="'+identifier+'">'+
   '<span class="ls-def-context-glyph">'+(item.label==="Selections"?"◈":"▧")+'</span><span>'+safe(item.label)+'</span><span class="ls-def-context-arrow">›</span></button>'+
   '<div id="'+identifier+'" class="ls-def-context-nested" role="menu" aria-label="'+safe(item.label)+'">'+
    item.children.map(label=>'<button type="button" role="menuitem"'+(item.label==="Selections"&&label==="Adjacent"
      ?' data-definitions-command="adjacent"':' disabled aria-disabled="true"')+'>'+
      '<span class="ls-def-context-glyph">◈</span><span>'+safe(label)+'</span></button>').join("")+
   '</div></div>';
 }).join("")+'</div>').join("");
 document.body.append(menu);
 const parents=()=>[...menu.querySelectorAll("[data-def-context-parent]")];
 const clearOpen=()=>{
  for(const group of parents()){
   group.classList.remove("is-open");
   group.querySelector(":scope > button")?.setAttribute("aria-expanded","false");
  }
 };
 const hide=()=>{if(menu.hidden)return;menu.hidden=true;clearOpen();};
 function show(x,y){
  clearOpen();menu.hidden=false;
  menu.style.left=Math.max(6,Math.min(x,innerWidth-menu.offsetWidth-8))+"px";
  menu.style.top=Math.max(6,Math.min(y,innerHeight-menu.offsetHeight-8))+"px";
  menu.classList.toggle("ls-def-context-flip",menu.getBoundingClientRect().right+275>innerWidth);
  const first=menu.querySelector(":scope .ls-def-context-group button:not(:disabled)");
  first?.focus({preventScroll:true});
 }
 const expand=group=>{
  if(!group)return;
  for(const other of parents()){
   const active=other===group;
   other.classList.toggle("is-open",active);
   other.querySelector(":scope > button")?.setAttribute("aria-expanded",String(active));
  }
 };
 tree.addEventListener("contextmenu",event=>{
  const summary=event.target.closest("#ls-component-definitions > summary");
  if(!summary)return;
  event.preventDefault();event.stopPropagation();
  show(event.clientX,event.clientY);
 });
 menu.addEventListener("pointerover",event=>{
  const group=event.target.closest("[data-def-context-parent]");
  if(group)expand(group);
 });
 menu.addEventListener("click",event=>{
  const button=event.target.closest("button");
  if(!button)return;
  const command=button.dataset.definitionsCommand;
  if(command==="adjacent"||command==="pml"){
   hide();
   document.dispatchEvent(new CustomEvent(command==="pml"?"ls-open-pml-editor":"ls-open-adjacent-editor",{detail:{create:true}}));
   return;
  }
  if(button.disabled)return;
  const group=button.closest("[data-def-context-parent]");
  if(group)expand(group);
 });
 document.addEventListener("pointerdown",event=>{
  if(!menu.hidden&&!menu.contains(event.target))hide();
 },true);
 document.addEventListener("keydown",event=>{
  if(menu.hidden)return;
  if(event.key==="Escape"){event.preventDefault();hide();return;}
  const focused=document.activeElement;
  if(event.key==="ArrowRight"&&focused?.closest("[data-def-context-parent]")){
   const group=focused.closest("[data-def-context-parent]");
   expand(group);group.querySelector(".ls-def-context-nested button:not(:disabled)")?.focus({preventScroll:true});
  }
  if(event.key==="ArrowLeft"&&focused?.closest(".ls-def-context-nested")){
   const parent=focused.closest("[data-def-context-parent]");
   parent?.querySelector(":scope > button")?.focus({preventScroll:true});
   clearOpen();
  }
 });
 document.addEventListener("ls-wizard-finished",hide);
 window.LSWaveDefinitionsContext=Object.freeze({isOpen:()=>!menu.hidden,close:hide});
})();
