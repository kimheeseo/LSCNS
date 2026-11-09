(function comsolConfigurationPanel(){
 "use strict";
 const get=id=>document.getElementById(id);
 const sidebar=document.querySelector(".layout > .sidebar");
 if(!sidebar||get("ls-config-pane"))return;
 const LENGTHS={m:1,mm:1e-3,um:1e-6,nm:1e-9,pm:1e-12};
 const K="LS optical settings: screenshot-inspired parameter and circle editors";
 const baseCircle=()=>[
  {label:"Circle 1",type:"Solid",radius:"rb",sector:"360",base:"Center",x:"0[um]",y:"0[um]",rotation:"0",layer:"0[um]"},
  {label:"Circle 2",type:"Solid",radius:"d/2",sector:"360",base:"Center",x:"0[um]",y:"0[um]",rotation:"0",layer:"0[um]"}
 ];
 let circles=baseCircle(),selected=1,params=[],view=null;
 const fromInput=(id,other)=>{const n=Number(get(id)?.value);return Number.isFinite(n)&&n>0?n:other;};
 const precision=x=>Number(x.toPrecision(9)).toString();
 // Reference values transcribed from the user's COMSOL Parameters 1 screenshot.
 // These model/geometry defaults are intentionally independent of solver presets.
 function defaults(){
  return [
   {name:"d",expression:"5[um]",description:"Hole diameter"},
   {name:"Delta",expression:"6.75[um]",description:"Hole separation"},
   {name:"rb",expression:"10[um]",description:"Radius for calculation domain"},
   {name:"nbg",expression:"1.45",description:"Background refractive index"},
   {name:"nair",expression:"1",description:"Air refractive index"},
   {name:"wl",expression:"1.45[um]",description:"Wavelength"},
   {name:"neffMax",expression:"1.4454",description:"Maximum expected effective index"},
   {name:"neffMin",expression:"1.429",description:"Minimum expected effective index"},
   {name:"wlr",expression:"wl/sqrt(nbg^2-neffMax^2)",description:"Wavelength in the radial direction"}
  ];
 }
 // Preserve established solver presets on Cancel / legacy model switch.
 // A newly completed ewfd wizard (or Restore defaults) uses the nine-row
 // screenshot-based default Parameters 1 table above.
 function legacyDefaults(){
  const wl=fromInput("wavelength_nm",1550),rc=fromInput("core_radius_um",8),rb=fromInput("outer_radius_um",40);
  const dn=Number(get("delta_n_core")?.value)||0.005;
  return [
   {name:"d",expression:`${precision(2*rc)}[um]`,description:"Core diameter / Circle 2"},
   {name:"Delta",expression:"6.75[um]",description:"Hole separation in reference screenshot (not linked to radial solver)"},
   {name:"dn_core",expression:precision(dn),description:"Core-cladding index difference for radial FEM"},
   {name:"rb",expression:`${precision(rb)}[um]`,description:"Computation boundary / Circle 1"},
   {name:"nbg",expression:get("clad_material")?.value==="fixed"?String(get("clad_index_override").value):"1.45",description:"Background index (reference; optional fixed nclad)"},
   {name:"nair",expression:"1",description:"Air refractive index (reference)"},
   {name:"wl",expression:`${precision(wl)}[nm]`,description:"Mode-analysis wavelength"},
   {name:"neffMax",expression:"nbg-0.0046",description:"Illustrative index upper bound (not a solver constraint)"},
   {name:"neffMin",expression:"nbg-0.021",description:"Illustrative index lower bound (not a solver constraint)"},
   {name:"wlr",expression:"wl/sqrt(nbg^2-neffMax^2)",description:"Derived radial wavelength (reference only)"}
  ];
 }
 params=legacyDefaults();
 // Public read-only parameter service for the independent geometry editor.
 window.LSWaveParameterService={
  evaluate:()=>evaluateAll(),
  rows:()=>params.map(p=>({...p})),
  calculate:(value)=>{const state=evaluateAll();if(state.errors.size)throw Error([...state.errors.values()][0]);return expression(value,name=>{if(!state.values.has(name))throw Error(`Unknown parameter ${name}`);return state.values.get(name);});}
 };
 function tokenize(input){
  const tokens=[],re=/\s*(?:(\d+(?:\.\d*)?|\.\d+)(?:[eE]([+-]?\d+))?|([a-zA-Z_]\w*)|(\[[a-zA-Zµμ]+\])|([()+\-*/^]))/gy;
  let i=0;
  while(i<input.length){
   if(/^\s*$/.test(input.slice(i)))break;
   re.lastIndex=i;const match=re.exec(input);
   if(!match)throw Error(`Unrecognized expression near "${input.slice(i,i+16)}"`);
   if(match[1]!==undefined)tokens.push({kind:"number",value:Number(match[1]+(match[2]!==undefined?"e"+match[2]:""))});
   else if(match[3])tokens.push({kind:"name",value:match[3]});
   else if(match[4])tokens.push({kind:"unit",value:match[4].slice(1,-1).replace(/[µμ]/g,"u")});
   else tokens.push({kind:"operator",value:match[5]});
   i=re.lastIndex;
  }
  return tokens;
 }
 function expression(input,resolve){
  const t=tokenize(String(input)),state={i:0};
  const peek=()=>t[state.i]?.value;
  const accept=val=>{if(peek()===val){state.i++;return true;}return false;};
  function atom(){
   const tok=t[state.i++];
   if(!tok)throw Error("Missing operand");
   if(tok.kind==="number"){
    let n=tok.value;
    if(t[state.i]?.kind==="unit"){
     const unit=t[state.i++].value;
     if(!(unit in LENGTHS))throw Error(`Unsupported unit [${unit}]`);
     n*=LENGTHS[unit];
    }
    return n;
   }
   if(tok.value==="("){const v=add();if(!accept(")"))throw Error("Missing closing parenthesis");return v;}
   if(tok.kind==="name"){
    if(accept("(")){
     const x=add();if(!accept(")"))throw Error("Missing closing parenthesis");
     if(tok.value==="sqrt"){if(x<0)throw Error("sqrt of negative value");return Math.sqrt(x);}
     if(tok.value==="abs")return Math.abs(x);
     throw Error(`Unsupported function ${tok.value}`);
    }
    return tok.value==="pi"?Math.PI:resolve(tok.value);
   }
   throw Error(`Unexpected token ${tok.value}`);
  }
  function power(){let v=atom();if(accept("^"))v=Math.pow(v,unary());return v;}
  function unary(){if(accept("+"))return unary();if(accept("-"))return -unary();return power();}
  function multiply(){let v=unary();while(peek()==="*"||peek()==="/"){const op=t[state.i++].value;const rhs=unary();if(op==="/"&&rhs===0)throw Error("Division by zero");v=op==="*"?v*rhs:v/rhs;}return v;}
  function add(){let v=multiply();while(peek()==="+"||peek()==="-"){const op=t[state.i++].value;v+= (op==="+"?1:-1)*multiply();}return v;}
  const result=add();
  if(state.i!==t.length)throw Error(`Unexpected token ${peek()}`);
  if(!Number.isFinite(result)||Math.abs(result)>1e100)throw Error("Non-finite result");
  return result;
 }
 function evaluateAll(){
  const lookup=new Map(),values=new Map(),errors=new Map(),busy=new Set();
  params.forEach((p,i)=>{
   if(!/^[a-zA-Z_]\w*$/.test(p.name))errors.set(i,"Invalid parameter name");
   else if(lookup.has(p.name)){errors.set(i,"Duplicate name");errors.set(lookup.get(p.name),"Duplicate name");}
   else lookup.set(p.name,i);
  });
  function lookupName(name){
   if(!lookup.has(name))throw Error(`Unknown parameter: ${name}`);
   const i=lookup.get(name);
   if(errors.has(i))throw Error(`${name}: ${errors.get(i)}`);
   if(values.has(name))return values.get(name);
   if(busy.has(i))throw Error(`Circular dependency: ${name}`);
   busy.add(i);
   try{const result=expression(params[i].expression,lookupName);values.set(name,result);return result;}
   finally{busy.delete(i);}
  }
  params.forEach((p,i)=>{if(!errors.has(i))try{lookupName(p.name);}catch(e){errors.set(i,e.message);}});
  return {values,errors};
 }
 function showValue(name,value){
  if(!Number.isFinite(value))return "—";
  // COMSOL-style display in SI units, e.g. 5E-6 m, 6.75E-6 m, 1.2564E-5 m.
  if(["d","Delta","rb","wl","wlr"].includes(name)){
   return Number(value.toPrecision(5)).toExponential().replace("e","E")+" m";
  }
  return precision(value);
 }
 function newEl(tag,attrs={},txt){
  const e=document.createElement(tag);
  for(const [k,v] of Object.entries(attrs))e.setAttribute(k,String(v));
  if(txt!==undefined)e.textContent=txt;
  return e;
 }
 const pane=newEl("section",{id:"ls-config-pane",class:"ls-config-pane",hidden:""});
 pane.innerHTML=`
  <div class="ls-config-heading"><div><strong>Settings</strong><small id="ls-config-subtitle">Parameters</small></div><button id="ls-config-back" type="button" title="Original fiber design fields">← Fiber settings</button></div>
  <section id="ls-parameters-view" class="ls-config-body">
   <div class="ls-setting-line"><label for="ls-param-label">Label</label><input id="ls-param-label" value="Parameters 1" /></div>
   <div class="ls-config-section-title">▾ Parameters</div>
   <div class="ls-params-scroll"><table class="ls-params-table" aria-label="Global Definitions Parameters"><thead><tr><th>Name</th><th>Expression</th><th>Value</th><th>Description</th><th></th></tr></thead><tbody id="ls-parameter-rows"></tbody></table></div>
   <div class="ls-config-actions"><button id="ls-params-add" type="button">＋ Add parameter</button><button id="ls-params-reset" type="button">Restore defaults</button><button id="ls-params-export" type="button">Export JSON</button><button id="ls-params-import" type="button">📂 Load from File</button><input type="file" id="ls-params-file" accept=".txt,.csv,.tsv,.json,text/plain,text/csv,application/json" hidden/></div>
   <p class="ls-config-help">Load from File: COMSOL-style tab/space-delimited TXT, CSV/TSV (Name, Expression, Description), or an exported parameter JSON file. Unsupported expressions are retained with a validation warning. Expressions support +, −, ×, ÷, ^, sqrt(), named parameters and [m], [mm], [um], [nm]. Values use SI units internally. All nine reference rows can be edited, removed, or supplemented. The screenshot reference settings are not necessarily valid FEM inputs. Radial FEM uses its existing delta_n_core input unless an optional dn_core parameter is added. Delta is a hole pitch, not index contrast.</p>
  </section>
  <section id="ls-geometry-view" class="ls-config-body" hidden>
   <div class="ls-config-actions ls-circle-choices"><button data-ls-circle="0" type="button">◯ Circle 1 (c1)</button><button data-ls-circle="1" type="button">◉ Circle 2 (c2)</button></div>
   <div class="ls-setting-line"><label for="ls-geo-label">Label</label><input id="ls-geo-label" /></div>
   <div class="ls-config-section-title">▾ Object Type</div>
   <div class="ls-geometry-field"><label for="ls-geo-type">Type</label><select id="ls-geo-type"><option>Solid</option><option>Curve</option></select></div>
   <div class="ls-config-section-title">▾ Size and Shape</div>
   <div class="ls-geometry-field"><label for="ls-geo-radius">Radius</label><input id="ls-geo-radius" spellcheck="false" /><span>m</span></div>
   <div class="ls-geometry-field"><label for="ls-geo-sector">Sector angle</label><input id="ls-geo-sector" spellcheck="false" /><span>deg</span></div>
   <div class="ls-config-section-title">▾ Position</div>
   <div class="ls-geometry-field"><label for="ls-geo-base">Base</label><select id="ls-geo-base"><option>Center</option></select></div>
   <div class="ls-geometry-field"><label for="ls-geo-x">x</label><input id="ls-geo-x" spellcheck="false" /><span>m</span></div>
   <div class="ls-geometry-field"><label for="ls-geo-y">y</label><input id="ls-geo-y" spellcheck="false" /><span>m</span></div>
   <div class="ls-config-section-title">▾ Rotation Angle</div>
   <div class="ls-geometry-field"><label for="ls-geo-rotation">Rotation</label><input id="ls-geo-rotation" spellcheck="false" /><span>deg</span></div>
   <div class="ls-config-section-title">▾ Layers</div>
   <div class="ls-geometry-field"><label for="ls-geo-layer">Layer 1 thickness</label><input id="ls-geo-layer" spellcheck="false" /><span>m</span></div>
   <div class="ls-geometry-preview"><svg id="ls-geo-svg" viewBox="0 0 420 310" role="img" aria-label="Circle geometry preview"></svg><div id="ls-geo-info" role="status"></div></div>
   <div class="ls-config-actions"><button id="ls-geo-selected" type="button">Build Selected</button><button id="ls-geo-all" type="button">Build All Objects</button></div>
   <label class="ls-config-check"><input id="ls-use-background" type="checkbox" /> Apply nbg as fixed cladding refractive index</label>
   <button class="ls-config-primary" id="ls-geo-apply" type="button">Apply concentric circles to Radial FEM</button>
   <p class="ls-config-help">The current radial solver supports centered 360° solid circles only (Circle 1: domain; Circle 2: core). Off-center geometry, sectors, extra layers and curves are editable preview settings, not an arbitrary 2D FEM geometry solver. Rotation of a complete circle has no physical effect.</p>
  </section>
  <p id="ls-config-status" class="ls-config-status" role="status" aria-live="polite"></p>`;
 sidebar.append(pane);
 const status=(txt,err=false)=>{const e=get("ls-config-status");e.textContent=txt;e.classList.toggle("error",err);};
 function renderParameters(){
  const body=get("ls-parameter-rows");body.replaceChildren();
  params.forEach((p,i)=>{
   const tr=newEl("tr",{"data-row":i});
   for(const field of ["name","expression"]){
    const td=newEl("td"),input=newEl("input",{"data-field":field,"data-index":i,"aria-label":`${field} ${i+1}`});
    input.value=p[field];td.append(input);tr.append(td);
   }
   const vCell=newEl("td"),value=newEl("span",{"data-param-value":i});vCell.append(value);tr.append(vCell);
   const descTd=newEl("td"),desc=newEl("input",{"data-field":"description","data-index":i,"aria-label":`description ${i+1}`});desc.value=p.description;descTd.append(desc);tr.append(descTd);
   const removeCell=newEl("td"),button=newEl("button",{"data-remove":i,type:"button","aria-label":`Remove parameter ${i+1}`},"×");removeCell.append(button);tr.append(removeCell);
   body.append(tr);
  });
  refreshParameters();
 }
 function refreshParameters(){
  const calc=evaluateAll();
  params.forEach((p,i)=>{
   const value=get("ls-parameter-rows")?.querySelector(`[data-param-value="${i}"]`);
   if(!value)return;
   const err=calc.errors.get(i);value.textContent=err||showValue(p.name,calc.values.get(p.name));
   value.closest("tr")?.classList.toggle("ls-param-invalid",!!err);
   value.title=err||"SI value";
  });
  document.dispatchEvent(new CustomEvent("ls-global-parameters-changed",{detail:{valid:!calc.errors.size}}));
  return calc;
 }
 function renderCircle(){
  const c=circles[selected];
  for(const k of ["label","type","radius","sector","base","x","y","rotation","layer"])get("ls-geo-"+k).value=c[k];
  document.querySelectorAll("[data-ls-circle]").forEach(e=>e.classList.toggle("selected",Number(e.dataset.lsCircle)===selected));
  drawGeometry();
 }
 function getGeometry(){
  const {values,errors}=evaluateAll();
  if(errors.size)throw Error(`Parameter error: ${[...errors.values()][0]}`);
  const resolve=n=>{if(!values.has(n))throw Error(`Unknown parameter ${n}`);return values.get(n);};
  const objects=circles.map(c=>{
   const read=k=>expression(c[k],resolve);
   const o={...c,r:read("radius"),xMeters:read("x"),yMeters:read("y"),angle:read("sector"),degrees:read("rotation"),thickness:read("layer")};
   if(!(o.r>0)||!Number.isFinite(o.r))throw Error(`${c.label}: radius must be positive`);
   if(!(o.angle>0&&o.angle<=360))throw Error(`${c.label}: sector must be between 0 and 360 degrees`);
   if(o.thickness<0)throw Error(`${c.label}: layer thickness must not be negative`);
   return o;
  });
  return {objects,values};
 }
 function wedge(cx,cy,r,theta,rotate){
  if(theta>=359.999)return `<circle cx="${cx}" cy="${cy}" r="${r}"/>`;
  const start=(rotate-90)*Math.PI/180, end=start+theta*Math.PI/180;
  const sx=cx+r*Math.cos(start),sy=cy+r*Math.sin(start),ex=cx+r*Math.cos(end),ey=cy+r*Math.sin(end);
  return `<path d="M ${cx} ${cy} L ${sx} ${sy} A ${r} ${r} 0 ${theta>180?1:0} 1 ${ex} ${ey} Z"/>`;
 }
 function drawGeometry(){
  const svg=get("ls-geo-svg"),info=get("ls-geo-info");
  try{
   const {objects}=getGeometry(),extent=Math.max(...objects.map(o=>o.r+Math.hypot(o.xMeters,o.yMeters)),1e-6);
   const scale=125/extent;
   const graphics=objects.map((o,i)=>{
    const cx=210+o.xMeters*scale,cy=155-o.yMeters*scale,r=o.r*scale;
    const stroke=i===selected?"#1678bd":"#526e8b",fill=o.type==="Solid"?(i===selected?"#8ec8eb":"#deeaf5"):"none";
    return `<g fill="${fill}" fill-opacity=".64" stroke="${stroke}" stroke-width="${i===selected?2.8:1.8}">${wedge(cx,cy,r,o.angle,o.degrees)}</g><text x="${cx}" y="${Math.max(12,cy-r-8)}" text-anchor="middle" fill="${stroke}" font-size="12">${i===0?"c1":"c2"}</text>`;
   }).join("");
   svg.innerHTML=`<rect x="1" y="1" width="418" height="308" rx="4" fill="#fff" stroke="#cedae7"/><path d="M 20 155 H400 M210 18 V292" stroke="#dbe6ee" stroke-width="1" stroke-dasharray="5 5"/>${graphics}<text x="16" y="296" fill="#71879a" font-size="11">2D geometry preview · not FEM field result</text>`;
   info.textContent=`Circle 1: ${(objects[0].r*1e6).toFixed(3)} µm radius · Circle 2: ${(objects[1].r*1e6).toFixed(3)} µm radius`;
  }catch(e){svg.innerHTML="";info.textContent=e.message;}
 }
 function open(kind,circle){
  view=kind;sidebar.classList.add("ls-config-active");document.body.classList.add("ls-configuration-open");pane.hidden=false;
  get("ls-parameters-view").hidden=kind!=="parameters";get("ls-geometry-view").hidden=kind!=="geometry";
  get("ls-config-subtitle").textContent=kind==="parameters"?"Parameters":"Circle";
  status(kind==="parameters"?"Edit parameter expressions, then use Geometry → Apply to run the FEM.":"Editing geometry; FEM remains unchanged until Apply.");
  if(kind==="parameters")renderParameters();
  else{if(circle==="c1")selected=0;else if(circle==="c2")selected=1;renderCircle();}
  pane.scrollIntoView({block:"start"});
 }
 function close(){
  pane.hidden=true;view=null;sidebar.classList.remove("ls-config-active");document.body.classList.remove("ls-configuration-open");
 }
 function apply(){
  try{
   const {objects,values}=getGeometry();
   if(document.body.classList.contains("hcf-selection")||document.body.classList.contains("vector-selection")||get("profile")?.value==="mcf")throw Error("Radial FEM apply is available for standard circular SMF profiles only. Select G.652/G.654 or Step-index first.");
   const [out,core]=objects;
   if(out.type!=="Solid"||core.type!=="Solid"||out.angle!==360||core.angle!==360)throw Error("Only 360° solid circles can be applied to the radial FEM.");
   if([out,core].some(o=>Math.abs(o.xMeters)>1e-12||Math.abs(o.yMeters)>1e-12))throw Error("The radial FEM requires concentric circles centered at x=y=0.");
   if([out,core].some(o=>o.thickness>1e-12))throw Error("Additional geometric layers are preview-only. Set Layer 1 thickness to 0 to apply the two-circle model.");
   if(!(out.r>core.r))throw Error("Circle 1 must contain Circle 2 (rb > d/2).");
   const outer=out.r*1e6,inner=core.r*1e6,wl=values.get("wl")*1e9;
   // The reference table does not have dn_core; retain solver's own material input.
   // Users may optionally add dn_core to override it.
   const dn=values.has("dn_core")?values.get("dn_core"):Number(get("delta_n_core")?.value);
   if(!(outer>=15&&outer<=150&&inner>=0.5&&inner<=60&&outer>inner&&wl>=250&&wl<=3600&&dn>0&&dn<0.1))throw Error("Current radial solver limits: outer radius 15–150 µm, core radius 0.5–60 µm, wavelength 250–3600 nm, 0 < Δn < 0.1.");
   const bg=get("ls-use-background").checked?values.get("nbg"):null;
   if(bg!==null&&!(bg>=1&&bg<2))throw Error("nbg must be 1–2");
   const update=(id,v)=>{const e=get(id);if(!e)throw Error(`Missing FEM input: ${id}`);e.value=precision(v);e.dispatchEvent(new Event("input",{bubbles:true}));};
   update("outer_radius_um",outer);update("core_radius_um",inner);update("wavelength_nm",wl);update("delta_n_core",dn);
   if(get("ls-use-background").checked){
    get("clad_material").value="fixed";get("clad_material").dispatchEvent(new Event("input",{bubbles:true}));update("clad_index_override",bg);
   }
   status("Applied: radial FEM inputs updated. Click Compute / Mode Analysis to solve.");get("ls-geo-info").textContent+=" · Applied to existing radial FEM";
  }catch(e){status("Cannot apply: "+e.message,true);}
 }
 get("ls-config-back").addEventListener("click",close);
 get("ls-param-label").addEventListener("input",()=>{});
 get("ls-parameter-rows").addEventListener("input",event=>{
  const el=event.target.closest("[data-field]");if(!el)return;params[Number(el.dataset.index)][el.dataset.field]=el.value;refreshParameters();if(view==="geometry")drawGeometry();
 });
 get("ls-parameter-rows").addEventListener("click",event=>{
  const el=event.target.closest("[data-remove]");if(!el)return;params.splice(Number(el.dataset.remove),1);renderParameters();
 });
 get("ls-params-add").addEventListener("click",()=>{let i=1;while(params.some(p=>p.name==="p"+i))i++;params.push({name:"p"+i,expression:"1",description:"Custom parameter"});renderParameters();});
 get("ls-params-reset").addEventListener("click",()=>{params=defaults();renderParameters();status("Parameters 1 restored to the nine reference defaults. All values remain editable.");});

 // COMSOL-style parameter imports, without changing the physical solver on upload.
 function csvCells(line,delimiter=","){
  const out=[];let cur="",quoted=false;
  for(let i=0;i<line.length;i++){
   const c=line[i];
   if(c==='"'){if(quoted&&line[i+1]==='"'){cur+='"';i++;}else quoted=!quoted;}
   else if(c===delimiter&&!quoted){out.push(cur.trim());cur="";}
   else cur+=c;
  }
  out.push(cur.trim());return out;
 }
 function importParameters(text,filename){
  const ext=filename.split(".").pop().toLowerCase();
  let rows=[];
  if(ext==="json"){
   let data;try{data=JSON.parse(text);}catch(e){throw Error("JSON 형식이 올바르지 않습니다.");}
   if(Array.isArray(data))rows=data;
   else if(Array.isArray(data.parameters))rows=data.parameters;
   else throw Error("JSON 파일에 parameters 배열이 없습니다.");
   rows=rows.map(p=>({name:String(p.name??"").trim(),expression:String(p.expression??"").trim(),description:String(p.description??"").trim()}));
  }else{
   for(let raw of text.replace(/^\uFEFF/,"").split(/\r?\n/)){
    const line=raw.trim();
    if(!line||/^(?:#|%|\/\/)/.test(line))continue;
    let cols;
    if(ext==="csv"||(!line.includes("\t")&&line.includes(",")&&/^[^,]+,[^,]+,/.test(line)))cols=csvCells(line,",");
    else if(ext==="tsv"||line.includes("\t"))cols=csvCells(line,"\t");
    else{
     const m=line.match(/^([A-Za-z_]\w*)\s+(\S+)(?:\s+(.+))?$/);
     if(!m)throw Error("파라미터 행을 인식할 수 없습니다: "+line.slice(0,90));
     cols=[m[1],m[2],m[3]||""];
    }
    const [name,expr,...rest]=cols.map(x=>x.trim());
    if(/^(?:Name|Parameter)$/i.test(name||"")&&/^(?:Expression|Value)$/i.test(expr||""))continue;
    if(!name||!expr)throw Error("Name 또는 Expression이 비어 있습니다.");
    // A four-column COMSOL export may include a calculated Value before Description.
    const description=rest.length>=2?rest.slice(1).join(" "):(rest[0]||"");
    rows.push({name,expression:expr,description});
   }
  }
  if(rows.length<1)throw Error("불러올 파라미터가 없습니다.");
  if(rows.length>500)throw Error("한 번에 최대 500개 파라미터를 가져올 수 있습니다.");
  const names=new Set();
  for(const p of rows){
   if(!/^[A-Za-z_]\w*$/.test(p.name))throw Error("잘못된 파라미터 이름: "+p.name);
   if(names.has(p.name))throw Error("중복 파라미터: "+p.name);
   if(p.expression.length>1000||p.description.length>3000)throw Error(p.name+": 입력 문자열이 너무 깁니다.");
   names.add(p.name);
  }
  // Keep the current model-specific defaults if the import only supplies geometry parameters.
  // Imported definitions override names, preserving explicit user descriptions.
  const old=params;
  params=old.filter(p=>!names.has(p.name)).concat(rows);
  renderParameters();
  const evaluation=evaluateAll();
  status(rows.length+"개 파라미터를 불러왔습니다."+(evaluation.errors.size?" 평가 경고 "+evaluation.errors.size+"개: 표의 빨간 행을 확인하세요.":" Geometry → Apply를 눌러 계산에 반영하세요."),false);
 }
 get("ls-params-import").addEventListener("click",()=>get("ls-params-file").click());
 get("ls-params-file").addEventListener("change",async event=>{
  const file=event.target.files?.[0];
  if(!file)return;
  try{
   if(file.size>1024*1024*2)throw Error("파일 최대 크기는 2 MB입니다.");
   importParameters(await file.text(),file.name);
  }catch(err){status("Load from File 실패: "+err.message,true);}
  finally{event.target.value="";}
 });

 get("ls-params-export").addEventListener("click",()=>{
  const json=JSON.stringify({format:"LS Wave Optics parameters / geometry v1",label:get("ls-param-label").value,parameters:params,circles},null,2);
  const blob=new Blob([json],{type:"application/json"}),url=URL.createObjectURL(blob),a=document.createElement("a");a.href=url;a.download="LS_Fiber_Parameters_Geometry.json";a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
 });
 for(const key of ["label","type","radius","sector","base","x","y","rotation","layer"]){
  get("ls-geo-"+key).addEventListener("input",()=>{circles[selected][key]=get("ls-geo-"+key).value;drawGeometry();});
 }
 get("ls-geometry-view").addEventListener("click",e=>{
  const choice=e.target.closest("[data-ls-circle]");if(choice){selected=Number(choice.dataset.lsCircle);renderCircle();}
 });
 get("ls-geo-selected").addEventListener("click",()=>{drawGeometry();status(`Built ${circles[selected].label} in 2D preview only; compute uses Apply.`);});
 get("ls-geo-all").addEventListener("click",()=>{drawGeometry();status("Both circles refreshed in 2D preview. Radial FEM changes only after Apply.");});
 get("ls-geo-apply").addEventListener("click",apply);
 get("preset")?.addEventListener("change",()=>{params=legacyDefaults();circles=baseCircle();if(view==="parameters")renderParameters();else if(view==="geometry")renderCircle();});
 get("resetBtn")?.addEventListener("click",()=>{queueMicrotask(()=>{params=legacyDefaults();circles=baseCircle();if(view==="parameters")renderParameters();else if(view==="geometry")renderCircle();});});
 document.addEventListener("ls-open-wave-optics",()=>{if(view)close();},true);
 document.addEventListener("ls-wizard-finished",e=>{
  if(e.detail?.physics!=="ewfd"||e.detail?.study!=="mode")return;
  params=defaults();circles=baseCircle();
  get("ls-param-label").value="Parameters 1";
  renderParameters();
  status("New 2D model: nine COMSOL reference parameters loaded. Edit any row or use Load from File.");
 });
 document.addEventListener("ls-open-settings",e=>open(e.detail?.kind==="geometry"?"geometry":"parameters",e.detail?.circle));
 document.addEventListener("keydown",e=>{if(e.key==="Escape"&&view)close();});
})();
