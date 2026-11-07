// Studio chrome uses the existing host controls and single render loop.
export function createIndustrialStudio(T,status){
 const $=id=>document.getElementById(id),shell=$('shell'),main=document.querySelector('.main');
 document.body.classList.add('industrial-studio');
 const make=(tag,cls,text)=>{const e=document.createElement(tag);e.className=cls;if(text)e.textContent=text;return e;};
 const top=document.querySelector('.top'),actions=document.querySelector('.top-actions');
 const commands=make('details','studio-commands'),summary=make('summary','','설정 · 도구');commands.append(summary,actions);top.append(commands);
 const engine=make('span','studio-engine','엔진 확인 중');engine.id='studioEngine';top.prepend(engine);
 const quality=make('select','studio-quality');quality.setAttribute('aria-label','렌더링 품질');
 for(const [value,label] of [['low','저사양'],['medium','표준'],['high','고품질']]){const option=new Option(label,value);quality.add(option);}
 quality.value=window.LS3D_CONFIG?.quality||'medium';quality.addEventListener('change',()=>{if(window.LS3D_CONFIG)window.LS3D_CONFIG.quality=quality.value;});
 const qLabel=make('label','studio-quality-label','렌더링 ');qLabel.append(quality);top.insertBefore(qLabel,commands);
 const heading=document.querySelector('.heading'),operations=$('mobileOperations');
 if(operations&&heading)operations.prepend(heading);
 const viewbar=document.querySelector('.vp-top');main.insertBefore(viewbar,$('viewport'));
 const extras=make('details','studio-view-options'),extraSummary=make('summary','','추가 시점');extras.append(extraSummary);
 for(const id of ['btnWorkers','btnWalk','btnDoubleViewport','btnFull']){if($(id))extras.append($(id));}
 document.querySelector('.vp-tools').append(extras);
 // Keep simulation history available without making it dominate the canvas.
 const scenario=$('scenarioStepPanel'),control=document.querySelector('.control');
 const scenarioToggle=make('button','studio-scenario-toggle','시나리오 단계');scenarioToggle.type='button';scenarioToggle.setAttribute('aria-controls','scenarioStepPanel');scenarioToggle.setAttribute('aria-expanded','false');
 shell.classList.add('studio-scenario-closed');scenarioToggle.addEventListener('click',()=>{const closed=shell.classList.toggle('studio-scenario-closed');scenarioToggle.setAttribute('aria-expanded',String(!closed));});control.append(scenarioToggle);
 const modelNote=make('span','studio-model-note','CONCEPT MODEL · 교육 / 초기 검토용');$('viewport').append(modelNote);
 const legend=make('details','studio-legend'),legendTitle=make('summary','','범례 · 조작 안내');legend.append(legendTitle,document.querySelector('.vp-bottom'));$('viewport').append(legend);
 const nav=document.querySelector('.side .section-label');if(nav)nav.textContent='공간 탐색 / FACILITY';
 const logo=document.querySelector('.logo');logo.querySelector('b').textContent='LS · TWIN STUDIO';logo.querySelector('small').textContent='FACILITY VISUALIZATION';
 const tree=document.querySelector('.workspace-tree'),search=make('input','studio-tree-search');search.type='search';search.placeholder='장비 이름 / ID';search.setAttribute('aria-label','탐색 트리 장비 검색');tree.querySelector('h2').after(search);
 const empty=make('p','studio-tree-empty','검색 결과가 없습니다.');empty.hidden=true;tree.append(empty);
 let query='',treeFirst=null,treeCount=-1;
 const filter=()=>{let matches=0;tree.querySelectorAll('[data-asset-id]').forEach(b=>{const a=T.assetMap.get(b.dataset.assetId);const hit=!query||[a?.name,a?.id,a?.type].join(' ').toLocaleLowerCase().includes(query);b.hidden=!hit;if(hit)matches++;});empty.hidden=matches>0;};
 search.addEventListener('input',()=>{query=search.value.trim().toLocaleLowerCase();filter();});
 const inspector=document.querySelector('.right-inner'),intro=make('div','studio-inspector-intro');intro.innerHTML='<small>PROPERTIES</small><strong>선택 항목</strong><span>사양 · 부품 · 연결</span>';inspector.prepend(intro);if($('assetStatus'))intro.append($('assetStatus')); 
 const assetVisual=$('assetVisual');if(assetVisual){const details=make('details','studio-model-reference'),s=make('summary','','모델 안내 보기');assetVisual.before(details);details.append(s,assetVisual);}
 const workTools=document.querySelector('.workspace-tools');workTools.setAttribute('aria-label','3D 화면 조작');
 // Dismiss drawers on Escape while preserving the host's rack/inspection Escape handling.
 document.addEventListener('keydown',e=>{if(e.key==='Escape'){commands.open=false;extras.open=false;}});
 let lastZone='',lastMode='',lastEngine='',initialFit=false,scaleInitialized=false;
 function adaptDisplayMode(){if(document.body.classList.contains('display-mobile')&&!shell.classList.contains('explorer-closed'))document.querySelector('[aria-controls=workspaceExplorer]')?.click();}
 adaptDisplayMode();window.addEventListener('ls3d-display-mode-changed',adaptDisplayMode);
 function update(){
  const scaleToggle=$('phase1-toggle');if(!scaleInitialized&&scaleToggle){scaleInitialized=true;if(scaleToggle.getAttribute('aria-pressed')==='true')scaleToggle.click();}
  const state=status.status==='ready'?'THREE · r'+(status.adapter?.THREE.REVISION||'186'):'LEGACY · WEBGL';
  if(state!==lastEngine){lastEngine=state;engine.textContent=state;engine.dataset.ready=String(status.status==='ready');}
  if(!initialFit&&status.status==='ready'){initialFit=true;if(T.currentZone==='campus'&&!T.rackIsolateActive)T.cameraDesired.distance=Math.max(T.cameraDesired.distance,200);}
  if(document.activeElement!==quality)quality.value=window.LS3D_CONFIG?.quality||'medium';
  const oldQuality=$('qualityToggle');if(oldQuality){const names={low:'저사양',medium:'표준',high:'고품질'};oldQuality.textContent='품질: '+names[quality.value];}
  if(lastZone!==T.currentZone||lastMode!==T.facilityMode){lastZone=T.currentZone;lastMode=T.facilityMode;const zone=T.Z.find(z=>z.id===T.currentZone);viewbar.dataset.zone=T.currentZone;intro.querySelector('span').textContent=zone?.ko||'시설 모델';}
  const first=tree.querySelector('[data-asset-id]'),count=tree.querySelectorAll('[data-asset-id]').length;
  if(first!==treeFirst||count!==treeCount){treeFirst=first;treeCount=count;filter();}
 }
 update();return {update};
}
