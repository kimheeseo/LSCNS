/**
 * Resizable COMSOL-style desktop panels.
 * Sizes stay in this page's browser session and never modify the numerical
 * model. Accessible draggable/keyboard separator controls.
 */
(function initComsolPaneResizers(){
 "use strict";
 const layout=document.querySelector("body.ls-comsol-shell .layout");
 if(!layout||layout.querySelector("#ls-split-model"))return;
 layout.style.position="relative";
 const config={
  model:{from:".ls-model-tree",min:175,max:400,varName:"--ls-model-width",defaultSize:228},
  settings:{from:".sidebar",min:265,max:760,varName:"--ls-settings-width",defaultSize:330},
  graphics:{from:"main.workspace",min:150,max:400,varName:"--ls-study-width",defaultSize:194}
 };
 const handles=new Map();
 const clamp=(x,lo,hi)=>Math.min(hi,Math.max(lo,x));
 const getScreenLayout=()=>({
  compact:window.innerWidth<835,
  active:document.body.classList.contains("ls-material-active")||
         document.body.classList.contains("ls-adjacent-active"),
  noStudy:getComputedStyle(layout.querySelector(".ls-study-pane")||layout).display==="none"
 });
 const getElement=name=>layout.querySelector(config[name].from);
 const getCurrentWidth=name=>{
  if(name==="graphics"){
   const study=layout.querySelector(".ls-study-pane");
   return study?.getBoundingClientRect().width||config[name].defaultSize;
  }
  return getElement(name)?.getBoundingClientRect().width||config[name].defaultSize;
 };
 function adjust(name,nextWidth){
  const cfg=config[name],state=getScreenLayout(),containerWidth=layout.getBoundingClientRect().width;
  const minCanvas=state.active?230:270;
  const otherWidth=name==="model"?getCurrentWidth("settings"):getCurrentWidth("model");
  const studyWidth=state.active||state.noStudy?0:(name==="graphics"?0:getCurrentWidth("graphics"));
  const available=name==="graphics"
   ?containerWidth-getCurrentWidth("settings")-getCurrentWidth("model")-minCanvas
   :containerWidth-otherWidth-studyWidth-minCanvas;
  const max=Math.max(cfg.min,Math.min(cfg.max,available));
  const size=clamp(Math.round(nextWidth),cfg.min,max);
  document.body.style.setProperty(cfg.varName,size+"px");
  reposition();
  document.dispatchEvent(new CustomEvent("ls-panels-resized",{detail:{panel:name,width:size}}));
 }
 function reposition(){
  const rect=layout.getBoundingClientRect(),mode=getScreenLayout();
  for(const [name,handle]of handles){
   const target=getElement(name),r=target?.getBoundingClientRect();
   const usable=!!r&&r.width>0&&!mode.compact&&(name!=="graphics"||(!mode.active&&!mode.noStudy));
   handle.hidden=!usable;
   if(!usable)continue;
   handle.style.left=(r.right-rect.left)+"px";
   handle.style.top="0";
   handle.style.height=Math.max(150,rect.height)+"px";
   handle.setAttribute("aria-valuenow",String(Math.round(getCurrentWidth(name))));
  }
 }
 for(const name of Object.keys(config)){
  const cfg=config[name],handle=document.createElement("div");
  handle.id="ls-split-"+name;
  handle.className="ls-panel-splitter";handle.hidden=true;
  handle.tabIndex=0;handle.setAttribute("role","separator");
  handle.setAttribute("aria-orientation","vertical");
  handle.setAttribute("aria-label",{
   model:"Resize Model Builder width",
   settings:"Resize Settings width",
   graphics:"Resize Graphics and Study widths"
  }[name]);
  handle.setAttribute("aria-valuemin",String(cfg.min));
  handle.setAttribute("aria-valuemax",String(cfg.max));
  handle.title="Drag to change pane width; arrow keys adjust by 12 px";
  const grip=document.createElement("span");grip.textContent="⋮";grip.setAttribute("aria-hidden","true");
  handle.append(grip);layout.append(handle);handles.set(name,handle);
  let startX=0,startSize=0,dragging=false;
  handle.addEventListener("pointerdown",e=>{
   if(e.button!==0)return;
   e.preventDefault();dragging=true;startX=e.clientX;startSize=getCurrentWidth(name);
   handle.setPointerCapture(e.pointerId);handle.classList.add("is-dragging");
  });
  handle.addEventListener("pointermove",e=>{
   if(!dragging)return;
   const delta=e.clientX-startX;
   adjust(name,name==="graphics"?startSize-delta:startSize+delta);
  });
  const stop=()=>{dragging=false;handle.classList.remove("is-dragging");};
  handle.addEventListener("pointerup",stop);
  handle.addEventListener("pointercancel",stop);
  handle.addEventListener("lostpointercapture",stop);
  handle.addEventListener("keydown",e=>{
   if(!["ArrowLeft","ArrowRight","Home"].includes(e.key))return;
   e.preventDefault();
   const direction=e.key==="Home"?0:e.key==="ArrowRight"?1:-1;
   adjust(name,e.key==="Home"?cfg.defaultSize:getCurrentWidth(name)+12*direction*(name==="graphics"?-1:1));
  });
 }
 const observer=new ResizeObserver(reposition);
 observer.observe(layout);
 new MutationObserver(reposition).observe(document.body,{attributes:true,attributeFilter:["class"]});
 document.addEventListener("ls-open-cad",reposition);
 document.addEventListener("ls-open-material-editor",reposition);
 document.addEventListener("ls-adjacent-editor-activated",reposition);
 document.addEventListener("ls-panels-resized",()=>{});
 window.addEventListener("resize",reposition);
 requestAnimationFrame(reposition);
 window.LSWavePaneResizers=Object.freeze({get:()=>Object.fromEntries(Object.entries(config).map(([name,cfg])=>[name,document.body.style.getPropertyValue(cfg.varName)||"auto"])),set:(name,width)=>{if(config[name])adjust(name,width);}});
})();
