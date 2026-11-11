/** Common Model Builder operations, aligned to Introduction to COMSOL
 * Multiphysics 5.5, pp. 10, 18, 25, 75, 176–178. No solver changes. */
(function guidebookInteractions(){
 "use strict";
 const $=id=>document.getElementById(id),tree=document.querySelector('.ls-model-tree');
 if(!tree||$('ls-guide-help'))return;
 const visible=el=>!!el&&!el.hidden&&el.getClientRects().length>0;
 const editing=el=>!!el?.closest('input:not([type="checkbox"]):not([type="radio"]),textarea,select,[contenteditable="true"]');
 const topics={
  parameters:{title:'Parameters',pages:'21–25, 65–67',text:'Global Definitions → Parameters 1에서 이름·식·단위를 입력합니다. 매개변수 이름은 대소문자를 구분하며 Geometry 치수에 같은 이름을 사용할 수 있습니다.',steps:['Load from file로 매개변수 표를 읽습니다.','Circle의 Radius와 Layers에는 rb+wlr, wlr 같은 식을 입력할 수 있습니다.','Geometry를 변경한 뒤 Build하고, 지원되는 동심원 모델은 Apply Geometry to FEM 후 다시 계산합니다.']},
  geometry:{title:'Geometry',pages:'17–18, 68–72, 176–178',text:'Model Builder의 Circle·Array·Rotate·Mirror 노드를 선택하면 해당 Settings가 열립니다. 입력 개체와 변환 결과를 Graphics에서 확인합니다.',steps:['Array: 입력 개체를 선택한 뒤 Build Selected를 누르면 배열 결과가 보입니다. Rectangular는 x/y 개수, Linear는 Size를 사용합니다.','F6: 이전 노드까지 Build · F7: 선택 노드 Build · F8: 전체 Geometry Build.','F2: 이름 변경 · F3: 비활성화 · F4: 활성화. F3/F4는 현재 상태를 반전하지 않습니다.','Ctrl+↑/↓: 선택 Geometry 노드 순서 이동. 빈 Graphics 영역 드래그로 이동하고 휠로 확대·축소합니다. Zoom Extents는 전체 도형이 보이도록 크기와 중심을 맞춥니다.']},
  material:{title:'Materials · Selection',pages:'25, 73–76, 177',text:'Materials → Blank Material에서 Geometric entity level과 Selection을 설정합니다. Manual에서는 영역을 개별 선택하며 같은 영역을 다시 클릭하면 해제됩니다.',steps:['All / Clear 버튼 또는 Graphics에서 Ctrl+A / Ctrl+D로 현재 레벨의 선택을 관리합니다.','선택 목록의 체크박스에 초점을 둔 뒤 Delete를 누르면 해당 항목만 선택 해제됩니다.','Refractive index의 n과 ki에 유효한 식을 입력하면 Status에 ✓가 표시됩니다. 사용자 Material 배정은 현재 Geometry 미리보기용입니다.']},
  adjacent:{title:'Definitions · Adjacent',pages:'25, 75, 177',text:'Input Entities의 Domain·Boundary·Point 레벨을 선택하고 Graphics에서 개별 입력 항목을 선택합니다. 선택한 입력은 목록과 파란 경계에 함께 표시됩니다.',steps:['All / Clear 또는 Graphics의 Ctrl+A / Ctrl+D는 현재 레벨에만 적용됩니다.','선택 목록의 항목에 초점을 두고 Delete를 누르면 해당 입력만 제거합니다.','Output Entities 설정은 현재 미리보기입니다. FEM 토폴로지에서 인접 경계를 산출하는 기능은 아직 연결되지 않았습니다.']},
  pml:{title:'Definitions · Perfectly Matched Layer',pages:'18, 25, 177',text:'Definitions 오른쪽 메뉴 → Perfectly Matched Layer를 선택하면 Artificial Domains 아래에 PML 노드가 생깁니다. 공통 모델 트리와 Selection 조작은 가이드북 기준입니다.',steps:['Manual에서 영역을 선택하고 All / Clear 또는 Ctrl+A / Ctrl+D로 선택을 관리합니다.','Geometry Type, stretching, wavelength source 설정은 노드를 다시 열어도 유지됩니다.','PML 고유 설정은 별도 구현된 미리보기입니다. 이 입문서에 Wave Optics PML 수식이 수록되어 있거나 실제 흡수 해석이 연결되었다는 뜻은 아닙니다.']},
  study:{title:'Mesh · Study · Results',pages:'19–20, 46–58, 85–95',text:'지원되는 동심원 모델은 Geometry → Materials → Mesh → Study → Results 순서로 계산합니다. 임의의 2D 도형 미리보기가 실제 FEM 해석 영역으로 자동 변환되지는 않습니다.',steps:['Model Workflow에서 Geometry 적용 여부와 Study 차단 사유를 확인합니다.','Mode Analysis의 Compute는 현재 연결된 스칼라 Radial FEM을 실행합니다.','입력을 바꾼 뒤에는 재계산하여 현재 Results인지 확인합니다.']},
  navigation:{title:'Model Builder · Keyboard',pages:'10, 18, 176–178',text:'현재까지 수정된 모델 트리와 Settings를 유지하면서 키보드로 노드를 탐색합니다. F1 도움말은 선택한 설정에 맞춰 갱신됩니다.',steps:['↑/↓: 보이는 이전/다음 노드 · ←/→: 가지 접기/펼치기.','Alt+←/→: 이전/다음 선택 노드로 돌아가기.','Graphics나 선택 목록에서 Ctrl+A: 전체 선택 · Ctrl+D: 선택 해제. 텍스트 입력 중에는 일반 편집 동작을 유지합니다.']}
 };
 const help=document.createElement('section');help.id='ls-guide-help';help.hidden=true;
 help.setAttribute('role','region');help.setAttribute('aria-label','Context help');
 help.innerHTML='<header><strong>Help · F1</strong><button type="button" id="ls-guide-help-close" aria-label="Close help">×</button></header><div id="ls-guide-help-content"></div>';
 document.body.append(help);
 let wavePanel='',lastFocus=null,lastTopic='navigation';
 function topic(){
  if(visible($('ls-pml-editor')))return 'pml';
  if(visible($('ls-adjacent-editor')))return 'adjacent';
  if(visible($('ls-materials-editor')))return 'material';
  if(visible($('ls-cad-pane')))return 'geometry';
  if(visible($('ls-config-pane')))return 'parameters';
  if(visible($('ls-wave-settings')))return /Materials|Core|Cladding/.test(wavePanel)?'material':'study';
  return 'navigation';
 }
 function renderHelp(){
  if(help.hidden)return;
  const key=topic(),t=topics[key];lastTopic=key;
  const content=$('ls-guide-help-content');content.replaceChildren();
  const title=document.createElement('h3');title.textContent=t.title;content.append(title);
  const p=document.createElement('p');p.textContent=t.text;content.append(p);
  const list=document.createElement('ul');for(const step of t.steps){const li=document.createElement('li');li.textContent=step;list.append(li);}content.append(list);
  const source=document.createElement('p');source.className='ls-guide-help-source';source.textContent='공통 조작 참고: Introduction to COMSOL Multiphysics 5.5, pp. '+t.pages;content.append(source);
  const note=document.createElement('small');note.textContent='현재 툴의 지원 범위에 맞춘 안내입니다. 모델 트리를 선택하면 안내가 함께 바뀝니다.';content.append(note);
 }
 function openHelp(){if(help.hidden)lastFocus=document.activeElement;help.hidden=false;renderHelp();$('ls-guide-help-close').focus();}
 function closeHelp(){help.hidden=true;if(lastFocus?.isConnected)lastFocus.focus({preventScroll:true});}
 $('ls-guide-help-close').addEventListener('click',closeHelp);
 document.addEventListener('ls-show-context-help',openHelp);
 document.addEventListener('ls-open-wave-optics',e=>{wavePanel=e.detail?.panel||'';});
 // Editors often rebuild their tree buttons. Store stable selectors, not DOM nodes.
 const attributes=['data-cad-node','data-cad-cumulative','data-ls-blank-material','data-adjacent-node','data-pml-node','data-action','data-select-preset'];
 function reference(el){
  for(const name of attributes)if(el?.hasAttribute(name))return el.tagName.toLowerCase()+'['+name+'="'+CSS.escape(el.getAttribute(name))+'"]';
  return null;
 }
 const history=[];let position=-1,replaying=false;
 document.addEventListener('click',e=>{
  if(!tree.contains(e.target))return;
  const el=e.target.closest('button.ls-tree-leaf,summary'),ref=reference(el);
  if(ref&&!replaying&&history[position]!==ref){history.splice(position+1);history.push(ref);position=history.length-1;}
  queueMicrotask(renderHelp);
 },true);
 function activate(el){
  const ref=reference(el);
  if(el.matches('button'))el.click();
  else if(el.dataset.action){const group=el.parentElement,wasOpen=group.open;el.click();group.open=wasOpen;}
  const target=(ref&&tree.querySelector(ref))||el;
  if(target.isConnected)target.focus({preventScroll:true});
  target.scrollIntoView({block:'nearest'});
 }
 function navigateHistory(delta){
  let next=position+delta;
  while(next>=0&&next<history.length){
   const el=tree.querySelector(history[next]);
   if(el){for(let parent=el.parentElement;parent&&parent!==tree;parent=parent.parentElement)if(parent.tagName==='DETAILS')parent.open=true;
    position=next;replaying=true;try{activate(el);}finally{replaying=false;}return;}
   next+=delta;
  }
 }
 function selectionOwner(){
  if(visible($('ls-materials-editor'))&&$('ls-material-select-mode')?.value==='manual')return 'material';
  if(visible($('ls-adjacent-editor')))return 'adjacent';
  if(visible($('ls-pml-editor'))&&$('ls-pml-selection')?.value==='manual')return 'pml';
  return null;
 }
 const rowSelectors={material:'[data-material-domain]',adjacent:'[data-adjacent-list-key]',pml:'[data-pml-domain-key]'};
 const keyAttribute={material:'data-material-domain',adjacent:'data-adjacent-list-key',pml:'data-pml-domain-key'};
 document.addEventListener('keydown',e=>{
  if(e.key==='F1'){e.preventDefault();e.stopImmediatePropagation();openHelp();return;}
  if(e.key==='Escape'&&!help.hidden){e.preventDefault();e.stopImmediatePropagation();closeHelp();return;}
  if(document.querySelector('dialog[open]')||editing(e.target))return;
  const commandKey=e.ctrlKey||e.metaKey;
  const owner=selectionOwner(),surface=e.target.closest?.('#ls-cad-graphics,.sidebar');
  if(owner&&surface){
   const command=commandKey&&!e.shiftKey?({a:'all',d:'clear'})[e.key.toLowerCase()]:null;
   const row=e.key==='Delete'?e.target.closest?.(rowSelectors[owner]):null;
   if(command||row){
    e.preventDefault();e.stopImmediatePropagation();
    const focusId=e.target.id,commandValue=command||'remove';
    document.dispatchEvent(new CustomEvent('ls-selection-command',{detail:{owner,command:commandValue,key:row?.getAttribute(keyAttribute[owner])}}));
    const newFocus=focusId?$(focusId):$('ls-cad-view');newFocus?.focus({preventScroll:true});return;
   }
  }
  if(e.altKey&&!commandKey&&['ArrowLeft','ArrowRight'].includes(e.key)){e.preventDefault();e.stopImmediatePropagation();navigateHistory(e.key==='ArrowLeft'?-1:1);return;}
  const node=e.target.closest?.('.ls-tree-scroll summary,.ls-tree-scroll button.ls-tree-leaf');
  if(!node||!tree.contains(node)||commandKey||e.altKey||e.shiftKey)return;
  const nodes=[...tree.querySelectorAll('.ls-tree-scroll summary,.ls-tree-scroll button.ls-tree-leaf:not(:disabled)')].filter(visible);
  const at=nodes.indexOf(node);if(at<0)return;
  let next=null;
  if(e.key==='ArrowDown')next=nodes[at+1];
  else if(e.key==='ArrowUp')next=nodes[at-1];
  else if(e.key==='ArrowRight'&&node.tagName==='SUMMARY'){
   if(!node.parentElement.open)node.parentElement.open=true;else next=nodes[at+1];
  }else if(e.key==='ArrowLeft'){
   if(node.tagName==='SUMMARY'&&node.parentElement.open)node.parentElement.open=false;
   else next=(node.tagName==='SUMMARY'?node.parentElement.parentElement:node.parentElement)?.closest('details')?.querySelector(':scope > summary');
  }else return;
  e.preventDefault();e.stopImmediatePropagation();if(next)activate(next);
 },true);
 const viewport=$('ls-cad-view');if(viewport){viewport.setAttribute('tabindex','0');viewport.setAttribute('aria-description','Graphics selection: Ctrl+A selects all; Ctrl+D clears the current entity level.');}
 // Context changes may originate from the ribbon or a settings link.
 const sidebar=document.querySelector('.layout > .sidebar');
 if(sidebar)new MutationObserver(()=>{if(!help.hidden&&topic()!==lastTopic)renderHelp();}).observe(sidebar,{attributes:true,subtree:true,attributeFilter:['hidden']});
})();
