/**
 * COMSOL-inspired multi-pane workbench skin for Universal Silica Fiber Lab.
 * Presentation/navigation only: do not replace, clone or intercept existing FEM,
 * HCF, MCF, Full-vector, PMD, BPM, JSON or validation solvers.
 * All original input ids and event handlers remain authoritative.
 */
const PRESETS = [
 ['g654','G.654.E · Large Aeff'], ['g652','G.652.D · SMF'],
 ['graded','Graded-index'],['trench','G.657 · Trench'],
 ['multi','Multilayer'],['comsol_step','COMSOL Step-index Benchmark'],
 ['mcf','MCF · Two-core'],['hcf','HCF / Nested Tube'],
 ['vector','Full-vector (Colab)']
];
const ICONS={
 model:'<rect x="4" y="4" width="16" height="16" rx="2"/><path d="M4 10h16M10 4v16"/>',
 file:'<path d="M6 3h8l5 5v13H6z"/><path d="M14 3v6h5M9 14h7M9 18h6"/>',
 geometry:'<circle cx="12" cy="12" r="8"/><path d="M4 12h16M12 4v16"/>',
 material:'<path d="m12 3 9 5-9 5-9-5zM3 12l9 5 9-5M3 16l9 5 9-5"/>',
 physics:'<path d="M2 12c3-10 5 10 8 0s5-10 8 0 4 3 4 0"/>',
 mesh:'<path d="m3 4 18 3-4 14L3 18zM3 4l10 7 8-4M3 18l10-7 4 10M13 11l-5 9"/>',
 play:'<path d="m8 5 11 7-11 7z"/>',
 chart:'<path d="M3 20h18M5 16l5-5 4 3 6-8"/>',
 settings:'<path d="M4 7h16M4 12h16M4 17h16"/><circle cx="9" cy="7" r="2" fill="currentColor" stroke="none"/><circle cx="15" cy="12" r="2" fill="currentColor" stroke="none"/><circle cx="10" cy="17" r="2" fill="currentColor" stroke="none"/>',
 export:'<path d="M12 3v13M8 12l4 4 4-4M4 18v3h16v-3"/>',
 book:'<path d="M4 5h7a3 3 0 0 1 3 3v12a3 3 0 0 0-3-3H4zM14 8a3 3 0 0 1 3-3h4v12h-4a3 3 0 0 0-3 3"/>',
 layers:'<path d="m12 2 9 5-9 5-9-5 9-5zM3 12l9 5 9-5M3 17l9 5 9-5"/>',
 home:'<path d="M3 10 12 3l9 7v11H3zM9 21v-8h6v8"/>',
 refresh:'<path d="M20 7v5h-5M4 17v-5h5M5 9a8 8 0 0 1 14-2l1 5M4 12a8 8 0 0 0 15 3"/>',
 folder:'<path d="M3 6h7l2 2h9v12H3z"/>',
 help:'<circle cx="12" cy="12" r="9"/><path d="M9 9a3 3 0 1 1 4 3c-1 1-1 2-1 3M12 18h.01"/>',
 eye:'<path d="M2 12s4-7 10-7 10 7 10 7-4 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/>',
 check:'<path d="m4 12 5 5L20 6"/>',
 fit:'<path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/>',
 code:'<path d="m8 7-5 5 5 5M16 7l5 5-5 5M14 4l-4 16"/>',
 plus:'<path d="M12 4v16M4 12h16"/>',
 wave:'<path d="M2 12c3 0 3-6 6-6s3 12 6 12 3-6 8-6"/>',
 save:'<path d="M4 3h15l2 2v16H3V3zM7 3v6h10V3M7 21v-9h10v9"/>'
};
function icon(name){return `<svg class="ls-icon" viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.65" stroke-linecap="round" stroke-linejoin="round">${ICONS[name]||ICONS.settings}</svg>`;}
const COMSOL_NOTEBOOK='https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/universal_silica_v01/full_vector/LS_Full_Vector_Fiber_v0_1.ipynb';
const HCF_NOTEBOOK='https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/universal_silica_v01/hcf_complex_fem/HCF_Complex_FEM_Research.ipynb';
const ACTIONS={
 home:[['new','New model','file'],['settings','Parameters','settings'],['geometry','Geometry','geometry'],['materials','Materials','material'],['compute','Compute','play'],['mesh','Mesh study','mesh'],['export','Export JSON','save']],
 definitions:[['settings','Parameters','settings'],['comsol','COMSOL test','check'],['g654','G.654.E','geometry'],['hcf','HCF design','layers']],
 geometry:[['geometry','Geometry settings','geometry'],['cadAddCircle','Add Circle','plus'],['cadAddRotate','Rotate','refresh'],['hcf','HCF / NANF','layers'],['mcf','MCF two-core','model'],['fiber3d','3D propagation','eye']],
 sketch:[['geometry','Cross-section','geometry'],['g654','SMF geometry','model'],['hcf','Nested tubes','layers']],
 materials:[['materials','Refractive index','material'],['comsol','Fixed-index ref.','check'],['g654','Silica material','model']],
 physics:[['compute','Mode analysis','physics'],['vector','Full-vector mode','wave'],['pmd','PMD / DGD','physics'],['hcfFem','HCF complex FEM','mesh']],
 mesh:[['mesh','Mesh convergence','mesh'],['compute','Recompute','refresh'],['comsol','Benchmark mesh','chart']],
 study:[['compute','Mode analysis','play'],['sweep','Wavelength sweep','chart'],['mesh','Mesh convergence','mesh'],['vector','Full-vector / Colab','wave'],['hcfFem','HCF FEM / Colab','layers']],
 results:[['results','Results window','chart'],['geometry','Cross-section','geometry'],['fiber3d','3D propagation','eye'],['export','Export JSON','save']],
 developer:[['source','Source code','code'],['colab','Vector Colab','wave'],['help','Help / README','book']]
};
const LABELS=[['home','Home'],['definitions','Definitions'],['geometry','Geometry'],['sketch','Sketch'],['materials','Materials'],['physics','Physics'],['mesh','Mesh'],['study','Study'],['results','Results'],['developer','Developer']];

function findTarget(sel){return document.querySelector(sel);}
function jump(sel){
 const el=findTarget(sel);
 if(!el)return false;
 el.scrollIntoView({behavior:'smooth',block:'start',inline:'nearest'});
 if(el.matches('button,input,select'))el.focus({preventScroll:true});
 return true;
}
function choosePreset(name){
 const preset=findTarget('#preset');if(!preset)return;
 if(preset.value!==name){preset.value=name;preset.dispatchEvent(new Event('change',{bubbles:true}));}
 else if(name==='comsol_step')findTarget('#calculateBtn')?.click();
 updateSelection();
}
function runAction(action){
 switch(action){
 case 'new':findTarget('#resetBtn')?.click();return;
 case 'settings':document.dispatchEvent(new CustomEvent('ls-open-settings',{detail:{kind:'parameters'}}));return;
  case 'geometryOuter':document.dispatchEvent(new CustomEvent('ls-open-cad',{detail:{id:'c1'}}));return;
  case 'geometryInner':document.dispatchEvent(new CustomEvent('ls-open-cad',{detail:{id:'c2'}}));return;
  case 'cadAddCircle':document.dispatchEvent(new CustomEvent('ls-cad-add',{detail:{kind:'circle'}}));return;
  case 'cadAddRotate':document.dispatchEvent(new CustomEvent('ls-cad-add',{detail:{kind:'rotate'}}));return;
 case 'geometry':document.dispatchEvent(new CustomEvent('ls-open-cad',{detail:{id:'c2'}}));return;
 case 'materials':jump('#clad_material');return;
 case 'compute':findTarget('#calculateBtn')?.click();return;
 case 'mesh':findTarget('#meshBtn')?.click();return;
 case 'sweep':findTarget('#sweepBtn')?.click();return;
 case 'export':findTarget('#jsonBtn')?.click();return;
 case 'results':jump(document.body.classList.contains('vector-selection')?'#full-vector-tool':document.body.classList.contains('hcf-selection')?'#hcf-fem':'.metrics');return;
 case 'hcfFem':choosePreset('hcf');jump('#hcf-fem');return;
 case 'hcf':choosePreset('hcf');jump('#hcf-designer');return;
 case 'vector':choosePreset('vector');jump('#full-vector-tool');return;
 case 'mcf':choosePreset('mcf');jump('#crossSection');return;
 case 'g654':choosePreset('g654');return;
 case 'comsol':choosePreset('comsol_step');jump('#comsol-reference');return;
 case 'pmd':jump('#pmd-tool');return;
 case 'fiber3d':jump('#fiber-3d');return;
 case 'source':window.open('https://github.com/kimheeseo/LSCNS/tree/main/COMSOL/universal_silica_v01/web/','_blank','noopener');return;
 case 'colab':window.open(COMSOL_NOTEBOOK,'_blank','noopener');return;
 case 'help':window.open('https://github.com/kimheeseo/LSCNS/blob/main/COMSOL/universal_silica_v01/web/README.md','_blank','noopener');return;
 default:if(PRESETS.some(p=>p[0]===action))choosePreset(action);
 }
}
function ribbonHtml(){
 return `<div class="ls-ribbon" role="navigation" aria-label="광학 해석 리본 메뉴">
  <div class="ls-ribbon-tabs">${LABELS.map(([key,label])=>`<button type="button" class="ls-ribbon-tab${key==='home'?' active':''}" data-tab="${key}" aria-pressed="${key==='home'}">${label}</button>`).join('')}<span class="ls-ribbon-context">LS WAVE OPTICS <b>·</b> Wave Physics Workbench</span></div>
  <div class="ls-ribbon-body"><div class="ls-ribbon-tools" id="ls-ribbon-tools"></div><div class="ls-ribbon-note"><span class="ls-ribbon-key">MODEL</span><strong id="ls-ribbon-model">G.654.E · Large Aeff</strong><small>2D mode analysis / linked studies</small></div></div>
 </div>`;
}
function toolHtml(tab){return (ACTIONS[tab]||ACTIONS.home).map(([action,label,shape],i)=>`<button class="ls-ribbon-command" data-action="${action}" type="button" title="${label}"><span class="ls-command-icon ${i===0?'ls-icon-primary':''}">${icon(shape)}</span><span class="ls-command-label">${label}</span></button>`).join('<span class="ls-ribbon-separator" aria-hidden="true"></span>');}
function modelTreeHtml(){
 const options=PRESETS.map(([value,label])=>`<button class="ls-tree-leaf" data-select-preset="${value}" type="button"><span class="ls-tree-marker ls-dot-blue">◇</span><span>${label}</span></button>`).join('');
 return `<aside class="ls-model-tree" aria-label="Model Builder">
  <div class="ls-pane-title"><span>Model Builder</span><span class="ls-pane-glyph">⌕ ⋯</span></div>
  <div class="ls-tree-toolbar"><button type="button" data-tree-expand="1" title="트리 모두 펼치기">⊞</button><button type="button" data-tree-expand="0" title="트리 모두 접기">⊟</button><button type="button" data-action="new" title="기본 모델 복원">↺</button><span>MODEL</span></div>
  <div class="ls-tree-scroll">
   <div class="ls-tree-root">${icon('model')} LS Wave Optics (Model)</div>
   <details class="ls-tree-group" open><summary>Global Definitions</summary><div class="ls-tree-children"><button class="ls-tree-leaf" data-action="settings">${icon('settings')} Parameters 1</button><button class="ls-tree-leaf" data-action="materials">${icon('material')} Material Properties</button></div></details>
   <details class="ls-tree-group" open><summary>Component 1 (comp1)</summary><div class="ls-tree-children">
    <details class="ls-tree-group" open><summary data-action="geometry">Geometry 1</summary><div class="ls-tree-children"><div id="ls-geometry-node-list" aria-label="Geometry feature sequence"></div><div class="ls-tree-preset-group"><small>Design examples</small>${options}</div></div></details>
    <button class="ls-tree-leaf" data-action="materials">${icon('material')} Materials</button>
    <button class="ls-tree-leaf" data-action="compute">${icon('physics')} Electromagnetic Waves</button>
    <button class="ls-tree-leaf" data-action="mesh">${icon('mesh')} Mesh 1</button>
    </div></details>
   <details class="ls-tree-group" open><summary>Study 1</summary><div class="ls-tree-children"><button class="ls-tree-leaf" data-action="compute">${icon('play')} Mode Analysis</button><button class="ls-tree-leaf" data-action="sweep">${icon('chart')} Wavelength Sweep</button><button class="ls-tree-leaf" data-action="hcfFem">${icon('wave')} HCF Complex FEM</button></div></details>
   <details class="ls-tree-group" open><summary>Results</summary><div class="ls-tree-children"><button class="ls-tree-leaf" data-action="results">${icon('chart')} Derived Values</button><button class="ls-tree-leaf" data-action="fiber3d">${icon('eye')} 3D Propagation</button><button class="ls-tree-leaf" data-action="pmd">${icon('physics')} PMD / DGD</button></div></details>
  </div><div class="ls-tree-bottom">Solution status <span class="ls-green-indicator"></span> <span id="ls-tree-state">Ready</span></div>
 </aside>`;
}
function studyHtml(){
 return `<aside class="ls-study-pane" aria-label="Add Study 및 해석 종류">
  <div class="ls-pane-title"><span>Add Study</span><span class="ls-pane-glyph">✕</span></div>
  <button class="ls-add-study" type="button" id="ls-add-study">＋ Add Study</button>
  <div class="ls-study-heading">Studies</div>
  <div class="ls-study-list">
   <div class="ls-study-group">▸ General Studies</div>
   <button data-action="compute" class="ls-study-item">${icon('physics')} Mode Analysis</button>
   <button data-action="sweep" class="ls-study-item">${icon('chart')} Wavelength Sweep</button>
   <button data-action="mesh" class="ls-study-item">${icon('mesh')} Mesh Convergence</button>
   <div class="ls-study-group">▸ Advanced Research</div>
   <button data-action="vector" class="ls-study-item">${icon('wave')} Full-vector FDFD</button>
   <button data-action="hcfFem" class="ls-study-item">${icon('layers')} HCF Complex FEM</button>
   <button data-action="pmd" class="ls-study-item">${icon('physics')} PMD / DGD</button>
   <button data-action="fiber3d" class="ls-study-item">${icon('eye')} 3D BPM Propagation</button>
  </div>
  <div id="ls-study-more" hidden><div class="ls-study-heading">Select Design</div>${PRESETS.map(([value,label])=>`<button data-select-preset="${value}" class="ls-study-item">${icon('geometry')} ${label}</button>`).join('')}</div>
  <div class="ls-study-footer"><b>Physics interfaces in study</b><small>Scalar Radial FEM · optional Colab</small></div>
 </aside>`;
}
function graphicsHeaderHtml(){
 return `<div class="ls-graphics-head"><div class="ls-graphics-heading">${icon('geometry')} <span>Graphics</span> <small id="ls-graphics-mode">Geometry / Mode Results</small></div><div class="ls-graphics-toolbar"><button title="단면 보기" data-action="geometry" type="button">${icon('geometry')}</button><button title="계산 결과" data-action="results" type="button">${icon('chart')}</button><button title="3D 전파 도구" data-action="fiber3d" type="button">${icon('eye')}</button><button title="파장 스윕" data-action="sweep" type="button">${icon('wave')}</button></div></div>`;
}
function consoleHtml(){
 return `<section class="ls-output-console" aria-label="계산 메시지"><div class="ls-console-tabs"><button class="active" data-console-tab="messages" aria-pressed="true">Messages</button><button data-console-tab="progress" aria-pressed="false">Progress</button><button data-console-tab="log" aria-pressed="false">Log</button><span class="ls-console-spacer"></span><button id="ls-console-collapse" title="메시지 패널 접기" type="button">▾</button></div><div class="ls-console-text" id="ls-console-text" role="status" aria-live="polite">수치해석 엔진 준비 · 결과를 계산하면 메시지가 표시됩니다.</div></section>`;
}
function updateSelection(){
 const value=findTarget('#preset')?.value||'g654';
 const label=PRESETS.find(p=>p[0]===value)?.[1]||value;
 const banner=findTarget('#ls-ribbon-model');if(banner)banner.textContent=label;
 const mode=findTarget('#ls-graphics-mode');if(mode)mode.textContent=value==='hcf'?'Air-core / Nested Tube':value==='vector'?'Full-vector Mode / Colab':value==='comsol_step'?'Reference Step-index / neff':'Geometry / Mode Results';
 for(const button of document.querySelectorAll('[data-select-preset]')){
  const match=button.dataset.selectPreset===value;
  button.classList.toggle('active',match);
  button.setAttribute('aria-current',match?'true':'false');
 }
}
function init(){
 const layout=findTarget('.layout'),topbar=findTarget('.topbar'),main=findTarget('main.workspace');
 if(!layout||!topbar||!main||findTarget('.ls-ribbon'))return;
 document.body.classList.add('ls-comsol-shell');
 topbar.insertAdjacentHTML('afterend',ribbonHtml());
 layout.insertAdjacentHTML('afterbegin',modelTreeHtml());
 layout.insertAdjacentHTML('beforeend',studyHtml()+consoleHtml());
 main.insertAdjacentHTML('afterbegin',graphicsHeaderHtml());
 const ribbon=findTarget('.ls-ribbon'),tools=findTarget('#ls-ribbon-tools');
 function renderTab(tab){
  tools.innerHTML=toolHtml(tab);
  for(const button of ribbon.querySelectorAll('.ls-ribbon-tab')){
   const selected=button.dataset.tab===tab;
   button.classList.toggle('active',selected);button.setAttribute('aria-pressed',String(selected));
  }
 }
 renderTab('home');
 ribbon.addEventListener('click',event=>{
  const tab=event.target.closest('[data-tab]');if(tab){renderTab(tab.dataset.tab);return;}
  const item=event.target.closest('[data-action]');if(item)runAction(item.dataset.action);
 });
 layout.addEventListener('click',event=>{
  const cadNode=event.target.closest('[data-cad-node]');
  if(cadNode){document.dispatchEvent(new CustomEvent('ls-open-cad',{detail:{id:cadNode.dataset.cadNode}}));return;}
  const item=event.target.closest('[data-action]');
  if(item){runAction(item.dataset.action);return;}
  const select=event.target.closest('[data-select-preset]');if(select){choosePreset(select.dataset.selectPreset);return;}
  const expand=event.target.closest('[data-tree-expand]');if(expand){
   const opened=expand.dataset.treeExpand==='1';
   document.querySelectorAll('.ls-tree-group').forEach(el=>el.open=opened);
  }
 });
 findTarget('#ls-add-study')?.addEventListener('click',()=>{
  const panel=findTarget('#ls-study-more');panel.hidden=!panel.hidden;
  if(!panel.hidden)panel.scrollIntoView({block:'nearest'});
 });
 findTarget('#preset')?.addEventListener('change',()=>{updateSelection();writeConsole('messages');});
 findTarget('#resetBtn')?.addEventListener('click',()=>queueMicrotask(updateSelection));
 let consoleTab='messages',previousNotice='',history=[];
 function writeConsole(tab=consoleTab){
  consoleTab=tab;
  const notice=findTarget('#notice')?.textContent.trim()||'계산 메시지 없음';
  const state=findTarget('#engineState')?.textContent.trim()||'엔진 상태 확인 불가';
  if(tab==='log'&&notice!==previousNotice){
   previousNotice=notice;
   history.push(`${new Date().toLocaleTimeString('ko-KR',{hour12:false})} | ${notice}`);
   history=history.slice(-12);
  }
  const output=findTarget('#ls-console-text');if(!output)return;
  output.textContent=tab==='progress'?state:tab==='log'?history.join('\n')||'기록된 메시지 없음':notice;
  for(const b of document.querySelectorAll('[data-console-tab]')){
   const selected=b.dataset.consoleTab===tab;
   b.classList.toggle('active',selected);b.setAttribute('aria-pressed',String(selected));
  }
  const treeStatus=findTarget('#ls-tree-state');if(treeStatus)treeStatus.textContent=state.replace(/\s+/g,' ').slice(0,36);
 }
 findTarget('.ls-console-tabs')?.addEventListener('click',e=>{
  const tab=e.target.closest('[data-console-tab]');if(tab)writeConsole(tab.dataset.consoleTab);
 });
 const collapse=findTarget('#ls-console-collapse');collapse?.addEventListener('click',()=>{
  const minimized=layout.classList.toggle('ls-console-minimized');
  collapse.textContent=minimized?'▴':'▾';collapse.title=minimized?'메시지 패널 펼치기':'메시지 패널 접기';
 });
 const observer=new MutationObserver(()=>{writeConsole(consoleTab);});
 if(findTarget('#notice'))observer.observe(findTarget('#notice'),{childList:true,characterData:true,subtree:true});
 if(findTarget('#engineState'))observer.observe(findTarget('#engineState'),{childList:true,characterData:true,subtree:true});
 updateSelection();writeConsole();
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init,{once:true});else init();
