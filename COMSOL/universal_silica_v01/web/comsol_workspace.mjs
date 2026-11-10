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
 home:[['component1','Component 1','model'],['settings','Parameters','settings'],['geometry','Geometry 1','geometry'],['cadBuildAll','Build All','check'],['waveMaterials','Add Material','material'],['wavePhysics','ewfd Physics','physics'],['waveMesh','Build Mesh','mesh'],['waveSize','Mesh 1','mesh'],['compute','Compute','play'],['waveStudy','Study 1','settings'],['addStudy','Add Study','plus'],['waveResults','Add Plot Group','chart'],['resetDesktop','Reset Desktop','refresh']],
 definitions:[['settings','Parameters','settings'],['comsol','COMSOL test','check'],['g654','G.654.E','geometry'],['hcf','HCF design','layers']],
 geometry:[['geometry','Geometry settings','geometry'],['cadAddCircle','Circle','plus'],['cadAddRotate','Rotate','refresh'],['cadBuildAll','Build All','check'],['cadUndo','Undo','refresh'],['cadRedo','Redo','refresh'],['hcf','HCF / NANF','layers'],['mcf','MCF two-core','model'],['fiber3d','3D propagation','eye']],
 sketch:[['geometry','Cross-section','geometry'],['g654','SMF geometry','model'],['hcf','Nested tubes','layers']],
 materials:[['waveMaterials','Materials','material'],['waveCore','Core index','model'],['waveCladding','Cladding index','material'],['comsol','Index benchmark','check']],
 physics:[['wavePhysics','ewfd Settings','physics'],['waveEquation','Wave Equation','wave'],['wavePEC','PEC Boundary','layers'],['waveInitial','Initial Values','settings'],['vector','Full-vector / Colab','wave']],
 mesh:[['waveMesh','Mesh 1','mesh'],['waveSize','Size settings','settings'],['mesh','Mesh convergence','refresh'],['comsol','Benchmark mesh','chart']],
 study:[['waveMode','Mode Analysis Settings','physics'],['compute','Compute','play'],['sweep','Wavelength Sweep','chart'],['vector','Full-vector / Colab','wave'],['hcfFem','HCF FEM / Colab','layers']],
 results:[['waveResults','Results Settings','chart'],['waveField','Mode profile','wave'],['waveIndex','Index profile','chart'],['waveDerived','Derived Values','settings'],['export','Export JSON','save']],
 developer:[['source','Source code','code'],['colab','Vector Colab','wave'],['help','Context Help · F1','book']]
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
 case 'component1':{
   const tree=findTarget('.ls-model-tree');
   tree?.scrollIntoView({block:'nearest',behavior:'smooth'});
   const comp=[...document.querySelectorAll('.ls-model-tree .ls-tree-group>summary')].find(x=>x.textContent.includes('Component 1'));
   if(comp){comp.parentElement.open=true;comp.focus({preventScroll:true});}
   return;
 }
 case 'addStudy':findTarget('#ls-add-study')?.click();return;
 case 'resetDesktop':{
   document.dispatchEvent(new CustomEvent('ls-open-cad',{detail:{id:'root'}}));
   return;
 }
 case 'openDesign':findTarget('#importDesign')?.click();return;
 case 'saveDesign':findTarget('#designBtn')?.click();return;
 case 'settings':document.dispatchEvent(new CustomEvent('ls-open-settings',{detail:{kind:'parameters'}}));return;
 case 'definitionsAdjacent':document.dispatchEvent(new CustomEvent('ls-open-adjacent-editor',{detail:{create:true}}));return;
 case 'definitionsPML':document.dispatchEvent(new CustomEvent('ls-open-pml-editor',{detail:{create:true}}));return;
  case 'geometryOuter':document.dispatchEvent(new CustomEvent('ls-open-cad',{detail:{id:'c1'}}));return;
  case 'geometryInner':document.dispatchEvent(new CustomEvent('ls-open-cad',{detail:{id:'c2'}}));return;
  case 'cadAddCircle':document.dispatchEvent(new CustomEvent('ls-cad-add',{detail:{kind:'circle'}}));return;
  case 'cadAddRotate':document.dispatchEvent(new CustomEvent('ls-cad-add',{detail:{kind:'rotate'}}));return;
 case 'geometry':document.dispatchEvent(new CustomEvent('ls-open-cad',{detail:{id:'root'}}));return;
  case 'cadBuildAll':document.dispatchEvent(new CustomEvent('ls-open-cad',{detail:{id:'root'}}));document.dispatchEvent(new CustomEvent('ls-cad-command',{detail:{command:'build-all'}}));return;
  case 'cadUndo':document.dispatchEvent(new CustomEvent('ls-cad-command',{detail:{command:'undo'}}));return;
  case 'cadRedo':document.dispatchEvent(new CustomEvent('ls-cad-command',{detail:{command:'redo'}}));return;
 case 'materials':case 'waveMaterials':case 'waveCore':case 'waveCladding':case 'wavePhysics':case 'waveEquation':case 'wavePEC':case 'waveInitial':case 'waveMesh':case 'waveSize':case 'waveStudy':case 'waveMode':case 'waveResults':case 'waveField':case 'waveIndex':case 'waveDerived':
   document.dispatchEvent(new CustomEvent('ls-open-wave-optics',{detail:{panel:action}}));return;
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
 case 'help':document.dispatchEvent(new CustomEvent('ls-show-context-help'));return;
 default:if(PRESETS.some(p=>p[0]===action))choosePreset(action);
 }
}
function ribbonHtml(){
 return `<div class="ls-ribbon" role="navigation" aria-label="광학 해석 리본 메뉴">
  <div class="ls-ribbon-tabs"><button type="button" class="ls-ribbon-file" disabled aria-disabled="true">File</button>${LABELS.map(([key,label])=>`<button type="button" class="ls-ribbon-tab${key==='home'?' active':''}" data-tab="${key}" aria-pressed="${key==='home'}">${label}</button>`).join('')}<div class="ls-ribbon-quick" role="toolbar" aria-label="Quick Access Toolbar"><button type="button" title="Open Design JSON" aria-label="Open Design JSON" data-action="openDesign">Open</button><button type="button" title="Save Design JSON" aria-label="Save Design JSON" data-action="saveDesign">Save</button><span class="ls-quick-separator" aria-hidden="true"></span><button type="button" title="Undo Geometry (Ctrl+Z)" aria-label="Undo Geometry" data-action="cadUndo" disabled>↶</button><button type="button" title="Redo Geometry (Ctrl+Shift+Z)" aria-label="Redo Geometry" data-action="cadRedo" disabled>↷</button><button type="button" title="Build All Geometry (F8)" aria-label="Build All Geometry" data-action="cadBuildAll">▣</button></div><span class="ls-ribbon-context">LS WAVE OPTICS <b>·</b> Wave Physics Workbench</span></div>
  <div class="ls-ribbon-body"><div class="ls-ribbon-tools" id="ls-ribbon-tools"></div><div class="ls-ribbon-note"><span class="ls-ribbon-key">MODEL</span><strong id="ls-ribbon-model">G.654.E · Large Aeff</strong><small>2D mode analysis / linked studies</small></div></div>
 </div>`;
}

/** Screenshot ribbon commands use existing actions only. A null action is visual-only. */
const HOME_GROUPS=[
 {label:'Application',columns:[[['Application Builder',null,'application']]]},
 {label:'Model',columns:[[['Component 1','component1','component',true]],[['Add Component',null,'diamond',true]]]},
 {label:'Definitions',columns:[[['Parameters','settings','parameters',true]],[['Variables',null,'variables',true],['Functions',null,'function',true],['Parameter Case',null,'parameters',false,true]]]},
 {label:'Geometry',columns:[[['Build All','cadBuildAll','build']],[['Import',null,'import'],['LiveLink',null,'link',true,true]]]},
 {label:'Materials',columns:[[['Add Material','waveMaterials','addMaterial',false,false,true]]]},
 {label:'Physics',columns:[[['Electromagnetic Waves, Frequency Domain','wavePhysics','physics',true]],[['Add Physics',null,'addPhysics',false,false,true]]]},
 {label:'Mesh',columns:[[['Build Mesh','waveMesh','buildMesh']],[['Mesh 1','waveSize','mesh',true]]]},
 {label:'Study',columns:[[['Compute','compute','compute']],[['Study 1','waveStudy','study',true]],[['Add Study','addStudy','addStudy',false,false,true]]]},
 {label:'Results',columns:[[['Select Plot Group',null,'selectPlot',true,true]],[['Add Plot Group','waveResults','plot',true]]]},
 {label:'Layout',columns:[[['Windows',null,'windows',true]],[['Reset Desktop','resetDesktop','reset',true]]]}
];
const DEFINITION_GROUPS=[
 {label:'Variables',columns:[[['Local Variables',null,'variables']]]},
 {label:'Functions',columns:[[['Analytic',null,'analytic'],['Interpolation',null,'interpolation'],['Piecewise',null,'piecewise']],[['More Functions',null,'function',true]]]},
 {label:'Selections',columns:[[['Explicit',null,'explicit'],['Complement',null,'complement'],['Adjacent','definitionsAdjacent','adjacent']],[['Ball/Disk',null,'disk'],['Box',null,'box'],['Cylinder',null,'cylinder',false,true]],[['Union',null,'union'],['Intersection',null,'intersection'],['Difference',null,'difference']],[['Colors',null,'palette',true]]]},
 {label:'Probes',columns:[[['Update Probes',null,'updateProbe']],[['Probes',null,'probe',true]]]},
 {label:'Coupling',columns:[[['Nonlocal Couplings',null,'link',true]],[['Pairs',null,'pairs',true]]]},
 {label:'Coordinate Systems',columns:[[['Coordinate Systems',null,'coordinates',true]],[['Perfectly Matched Layer','definitionsPML','pml'],['Infinite Element Domain',null,'infinite',false,true],['Absorbing Layer',null,'absorbing',false,true]]]},
 {label:'Deformed Mesh',columns:[[['Moving Mesh',null,'movingMesh',true]],[['Deformed Geometry',null,'deformed',true]]]},
 {label:'Optimization',columns:[[['Optimization',null,'optimization',true]]]},
 {label:'View',columns:[[['View',null,'view']]]}
];
const RIBBON_GLYPHS={
 application:'<text x="6" y="22" font-size="27" fill="#49a6a0" stroke="none">A</text>',
 parameters:'<text x="5" y="22" font-size="24" fill="#559ec2" stroke="none">Pᵢ</text>',
 variables:'<text x="2" y="21" font-size="23" fill="#559ec2" stroke="none">a=</text>',
 function:'<text x="1" y="21" font-size="21" fill="#559ec2" stroke="none">f(x)</text>',
 component:'<ellipse cx="16" cy="13" rx="13" ry="10" fill="#579cc4"/><path d="m13 22 7 5-1-7" fill="#579cc4"/><text x="11" y="17" fill="white" stroke="none" font-size="14">Ω</text>',
 diamond:'<path d="m3 13 7-9h13l7 9-14 14zM3 13h27M10 4l6 23 7-23M10 4l-1 9 7 14 8-14-1-9"/>',
 build:'<path d="M6 5h20v23H6zM6 12h20M13 5v23M20 5v23" fill="#e8a482" stroke="#c37e58"/>',
 addMaterial:'<path d="M5 5h8v8H5zM17 5h8v8h-8zM5 17h8v8H5z" fill="#eeb961" stroke="#b18b52"/><path d="M23 16v13M17 23h13"/>',
 addPhysics:'<circle cx="16" cy="15" r="8"/><path d="M16 1v5M16 24v6M2 15h5M25 15h6M5 4l5 5M23 22l5 5M5 26l5-5M23 8l5-5M16 9v12M10 15h12"/>',
 buildMesh:'<path d="M6 5h20v23H6zM6 12h20M13 5v23M20 5v23" fill="#b1b4b7" stroke="#787e83"/>',
 compute:'<path d="M6 12h20M6 18h20" stroke="#2399aa" stroke-width="4"/>',
 study:'<path d="M2 15c5-17 12 17 17 0s12 17 12 0"/>',
 addStudy:'<path d="M2 8c4-14 8 14 12 0s8 14 12 0M20 15v15M13 23h16"/>',
 selectPlot:'<path d="M5 6h18v20H5zM9 3h18v20M13 1h18v20" fill="#dfa4c6" stroke="#bd8caa"/>',
 plot:'<path d="M4 5h21v19H4zM8 2h21v19H8z" fill="#c4b2c4" stroke="#7b7681"/><path d="M11 11h20v17H11z" fill="#7d647b"/>',
 windows:'<path d="M3 3h21v18H3zM9 9h21v18H9zM9 14h21" fill="#f1f8fc"/>',
 reset:'<path d="M3 5h26v23H3zM3 10h26M23 16a7 7 0 1 0 0 7M23 16v5h-5"/>',
 coordinates:'<path d="M5 3v24h24M5 27l15-16M1 8l4-5 4 5M24 23l5 4-5 4"/><text x="12" y="8" font-size="10" stroke="none" fill="#559ec2">z y</text>',
 view:'<path d="M5 3v24h24" stroke="#be5e9b"/><path d="m1 8 4-5 4 5M24 23l5 4-5 4"/><text x="10" y="12" font-size="12" stroke="none" fill="#6ba06b">xy</text>',
 mesh:'<path d="m2 28 14-25 14 25zM7 20h18M11 12h10M16 3v25M7 20l9 8 9-8M11 12l14 16M21 12 7 28" stroke="#888f94"/>',
 palette:'<path d="M27 19c-5 14-25 8-24-5S26-3 28 10c1 5-7 2-6 6z" fill="#ece9e2"/><circle cx="9" cy="11" r="2" fill="#66aac9"/><circle cx="17" cy="7" r="2" fill="#edbd50"/><circle cx="10" cy="20" r="2" fill="#ca74a3"/><circle cx="20" cy="21" r="2" fill="#81ac8c"/>',
 pairs:'<path d="M3 5h10v23H3zM20 5h10v23H20z" fill="#c2d7e2"/><path d="M13 10c10-8-3 20 7 12"/>',
 link:'<path d="m13 9 6-6c7-4 12 2 7 8l-7 7M18 23l-6 6c-7 4-12-2-7-8l7-7M10 23 23 9"/>',
 optimization:'<path d="m8 3-6 12 6 12h16l6-12-6-12z" fill="#e8f2f9"/><text x="11" y="23" font-size="25" stroke="none" fill="#559ec2">?</text>',
 pml:'<path d="M3 3h24v25H3zM7 3v25M11 3v25M15 3v25M19 3v25M23 3v25"/>',
 explicit:'<path d="M5 5h15v15H5zM12 12h15v15H12z" fill="#a5cde3"/>',
 disk:'<circle cx="16" cy="16" r="12" fill="#d4e8f1"/><ellipse cx="16" cy="16" rx="5" ry="12"/>',
 box:'<path d="m3 10 13-7 13 7v15l-13 5-13-5zM3 10l13 6 13-6M16 16v14" fill="#c5dfed"/>',
 analytic:'<path d="M2 27V4M2 27h28M4 24c10 0 7-20 16-20s4 22 10 22"/>',
 interpolation:'<path d="M2 27V4M2 27h28M4 24l7-15 9 9 9-14"/><path d="M4 22v4M9 7v4M18 16v4M27 2v4" stroke-width="4"/>',
 piecewise:'<path d="M2 27V4M2 27h28M3 22h9V10h9V4h8"/>',
 probe:'<path d="m7 3 19 4-5 20H3zM7 3l14 24M3 27 26 7M6 14h18"/><path d="M28 1 9 22" stroke-width="3"/>',
 deformed:'<path d="M3 4h26v23H3zM16 4v23" fill="#ccdee9"/>',
 movingMesh:'<path d="M4 4h23l-5 24H1zM2 16h23M10 4 6 28M19 4l-5 24"/><path d="m20 1 5 4-5 4M6 28l-5-4 5-4"/>',
};
function ribbonIcon(shape){
 const fallback={import:'folder',complement:'layers',adjacent:'layers',cylinder:'geometry',union:'layers',intersection:'layers',difference:'layers',updateProbe:'refresh',infinite:'layers',absorbing:'layers'};
 const drawing=RIBBON_GLYPHS[shape];
 return drawing?`<svg class="ls-icon" viewBox="0 0 32 32" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round">${drawing}</svg>`:icon(fallback[shape]||shape);
}
function screenshotToolHtml(groups,tab){
 return `<div class="ls-screenshot-groups ls-${tab}-groups" role="toolbar" aria-label="${tab==='home'?'Home':'Definitions'} toolbar">`+
 groups.map((group,i)=>`<section class="ls-shot-group" data-shot-group="${i}" aria-label="${group.label}"><div class="ls-shot-columns">`+
 group.columns.map(column=>`<div class="ls-shot-column${column.length>1?' ls-shot-stack':''}">`+column.map(([label,action,shape,dropdown,muted,accent])=>{
  const enabled=!!action;
  return `<button type="button" class="ls-shot-command${muted?' ls-shot-muted':''}${accent?' ls-shot-accent':''}"${enabled?` data-action="${action}"`:' disabled aria-disabled="true"'} title="${label}${enabled?'':' — 모양만 표시'}" aria-label="${label}"><span class="ls-shot-icon">${ribbonIcon(shape)}</span><span class="ls-shot-name">${label}${dropdown?'<span class="ls-shot-arrow" aria-hidden="true">▾</span>':''}</span></button>`;
 }).join('')+'</div>').join('')+`</div><div class="ls-shot-caption">${group.label}</div></section>`).join('')+'</div>';
}
function toolHtml(tab){
 if(tab==='home')return screenshotToolHtml(HOME_GROUPS,tab);
 if(tab==='definitions')return screenshotToolHtml(DEFINITION_GROUPS,tab);
 return (ACTIONS[tab]||ACTIONS.home).map(([action,label,shape],i)=>`<button class="ls-ribbon-command" data-action="${action}" type="button" title="${label}"><span class="ls-command-icon ${i===0?'ls-icon-primary':''}">${icon(shape)}</span><span class="ls-command-label">${label}</span></button>`).join('<span class="ls-ribbon-separator" aria-hidden="true"></span>');
}
function modelTreeHtml(){
 const options=PRESETS.map(([value,label])=>`<button class="ls-tree-leaf" data-select-preset="${value}" type="button"><span class="ls-tree-marker ls-dot-blue">◇</span><span>${label}</span></button>`).join('');
 return `<aside class="ls-model-tree" aria-label="Model Builder">
  <div class="ls-pane-title"><span>Model Builder</span><span class="ls-pane-glyph">⌕ ⋯</span></div>
  <div class="ls-tree-toolbar"><button type="button" data-tree-expand="1" title="트리 모두 펼치기">⊞</button><button type="button" data-tree-expand="0" title="트리 모두 접기">⊟</button><button type="button" data-action="new" title="기본 모델 복원">↺</button><span>MODEL</span></div>
  <div class="ls-tree-scroll">
   <div class="ls-tree-root">${icon('model')} LS Wave Optics (Model)</div>
   <details class="ls-tree-group" open><summary>Global Definitions</summary><div class="ls-tree-children"><button class="ls-tree-leaf" data-action="settings">${icon('settings')} Parameters 1</button><button class="ls-tree-leaf" data-action="waveMaterials">${icon('material')} Material Properties</button></div></details>
   <details class="ls-tree-group" open><summary>Component 1 (comp1)</summary><div class="ls-tree-children">
    <details id="ls-component-definitions" class="ls-tree-group" open><summary>Definitions</summary><div class="ls-tree-children"><details id="ls-definitions-selections-group" class="ls-tree-group" hidden open><summary>Selections</summary><div class="ls-tree-children" id="ls-definitions-selections-list"></div></details><details id="ls-definitions-artificial-group" class="ls-tree-group" hidden open><summary>Artificial Domains</summary><div class="ls-tree-children" id="ls-definitions-artificial-list"></div></details><span class="ls-tree-leaf">Component Definitions (2D)</span></div></details>
    <details class="ls-tree-group" open><summary data-action="geometry"><span id="ls-cad-root-name">Geometry 1</span></summary><div class="ls-tree-children"><div id="ls-geometry-node-list" aria-label="Geometry feature sequence"></div><span class="ls-tree-leaf ls-cad-union-reference">▦ Form Union (fin) · preview only</span><details id="ls-cad-cumulative-group" class="ls-tree-group" hidden open><summary>Cumulative Selections</summary><div class="ls-tree-children" id="ls-cad-cumulative-list"></div></details><div class="ls-tree-preset-group"><small>Design examples</small>${options}</div></div></details>
    <details class="ls-tree-group" open><summary data-action="waveMaterials">Materials</summary><div class="ls-tree-children"><div id="ls-cad-material-list" aria-label="Component material nodes"></div><button class="ls-tree-leaf" data-action="waveCore">${icon('material')} Core (n₁)</button><button class="ls-tree-leaf" data-action="waveCladding">${icon('material')} Cladding (n₂)</button></div></details>
    <details class="ls-tree-group" open><summary data-action="wavePhysics">Electromagnetic Waves, Frequency Domain (ewfd)</summary><div class="ls-tree-children"><button class="ls-tree-leaf" data-action="waveEquation">${icon('wave')} Wave Equation, Electric 1</button><button class="ls-tree-leaf" data-action="wavePEC">${icon('physics')} Perfect Electric Conductor 1</button><button class="ls-tree-leaf" data-action="waveInitial">${icon('settings')} Initial Values 1</button></div></details>
    <details class="ls-tree-group" open><summary data-action="waveMesh">Mesh 1</summary><div class="ls-tree-children"><button class="ls-tree-leaf" data-action="waveSize">${icon('mesh')} Size 1 · Radial mesh</button></div></details>
    </div></details>
   <details class="ls-tree-group" open><summary data-action="waveStudy">Study 1</summary><div class="ls-tree-children"><button class="ls-tree-leaf" data-action="waveMode">${icon('play')} Step 1: Mode Analysis</button><button class="ls-tree-leaf" data-action="sweep">${icon('chart')} Wavelength Sweep</button><button class="ls-tree-leaf" data-action="hcfFem">${icon('wave')} HCF Complex FEM (Colab)</button></div></details>
   <details class="ls-tree-group" open><summary data-action="waveResults">Results</summary><div class="ls-tree-children"><button class="ls-tree-leaf" data-action="waveField">${icon('wave')} Mode Intensity · I(r)</button><button class="ls-tree-leaf" data-action="waveIndex">${icon('chart')} Refractive Index · n(r)</button><button class="ls-tree-leaf" data-action="waveDerived">${icon('settings')} Derived Values · n_eff / MFD / A_eff</button><button class="ls-tree-leaf" data-action="fiber3d">${icon('eye')} 3D Propagation</button><button class="ls-tree-leaf" data-action="pmd">${icon('physics')} PMD / DGD</button></div></details>
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
  ribbon.classList.toggle("ls-ribbon-is-definitions",tab==="definitions");
  const screenshotTab=tab==="home"||tab==="definitions";
  ribbon.classList.toggle("ls-ribbon-is-screenshot",screenshotTab);
  document.body.classList.toggle("ls-screenshot-ribbon",screenshotTab);
  document.body.classList.toggle("ls-definitions-ribbon",tab==="definitions");
  const note=findTarget(".ls-ribbon-note");if(note)note.hidden=screenshotTab;
  for(const button of ribbon.querySelectorAll('.ls-ribbon-tab')){
   const selected=button.dataset.tab===tab;
   button.classList.toggle('active',selected);button.setAttribute('aria-pressed',String(selected));
  }
 }
 renderTab('home');
  document.addEventListener('ls-cad-history',event=>{
   for(const key of ['cadUndo','cadRedo']){
    const enabled=key==='cadUndo'?event.detail?.canUndo:event.detail?.canRedo;
    for(const btn of document.querySelectorAll('.ls-ribbon [data-action="'+key+'"]'))btn.disabled=!enabled;
   }
  });
 ribbon.addEventListener('click',event=>{
  const tab=event.target.closest('[data-tab]');if(tab){renderTab(tab.dataset.tab);return;}
  const item=event.target.closest('[data-action]');if(item&&!item.disabled&&item.getAttribute("aria-disabled")!=="true")runAction(item.dataset.action);
 });
 layout.addEventListener('click',event=>{
  const cadNode=event.target.closest('[data-cad-node]');
  if(cadNode){document.dispatchEvent(new CustomEvent('ls-open-cad',{detail:{id:cadNode.dataset.cadNode}}));return;}
  const item=event.target.closest('[data-action]');
  if(item){
    if(item.tagName==="SUMMARY"){
     const r=item.getBoundingClientRect(),arrow=event.clientX-r.left<22;
     // COMSOL Model Builder: click triangle to expand; click label to open Settings.
     if(arrow)return;
     event.preventDefault();
    }
    runAction(item.dataset.action);return;
   }
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

