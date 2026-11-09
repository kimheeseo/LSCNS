/**
 * Optical Fiber FEM Tool startup wizard.
 * UI/navigation only. The original numerical solvers and their input ids are untouched.
 * Scope: 2D Wave Optics -> Electromagnetic Waves, Frequency Domain (ewfd).
 */
(function startupWizard(){
 "use strict";
 const root=document.getElementById("ls-startup-wizard");
 if(!root)return;
 const dimensions=[
  ["3D","▣","3D"],["2D Axisymmetric","◉","2D Axisymmetric"],
  ["2D","◉","2D"],["1D Axisymmetric","⌢","1D Axisymmetric"],
  ["1D","━","1D"],["0D","•","0D"]
 ];
 const interfaces=[
  ["ewbe","Electromagnetic Waves, Beam Envelopes (ewbe)",false],
  ["ewfd","Electromagnetic Waves, Frequency Domain (ewfd)",true],
  ["teew","Electromagnetic Waves, Time Explicit (teew)",false],
  ["ewt","Electromagnetic Waves, Transient (ewt)",false]
 ];
 let step="dimension",dimension="",selected="ewfd",added=false,filter="";
 let opticsOpen=true,rayOpen=false,waveOpen=true;
 function q(id){return root.querySelector("#"+id);}
 function base(){
  return '<div class="lsw-window">'+
   '<div class="lsw-titlebar"><div class="lsw-quickicons"><span class="lsw-appmark">◈</span><span aria-hidden="true">▤　▱　▣　↶　↷</span></div><strong>Untitled · Optical Fiber FEM Tool</strong><span aria-hidden="true">—　□　×</span></div>'+
   '<nav class="lsw-top-tabs" aria-label="Workbench tabs">'+
    ["File","Home","Definitions","Geometry","Materials","Physics","Mesh","Study","Results","Developer"].map(function(s){return '<span class="lsw-top-tab">'+s+'</span>';}).join("")+
   '</nav><div class="lsw-page"><div class="lsw-body" id="lsw-body"></div><footer class="lsw-bottom">'+
    '<div class="lsw-buttons"><button type="button" class="lsw-btn" data-wizard="help"><span class="lsw-btn-icon">?</span> Help</button>'+
    '<button type="button" class="lsw-btn" data-wizard="cancel"><span class="lsw-btn-icon">⊗</span> Cancel</button>'+
    '<button type="button" class="lsw-btn lsw-btn-primary" data-wizard="done"><span class="lsw-btn-icon">☑</span> Done</button></div>'+
    '<span class="lsw-muted" id="lsw-footer-status"></span></footer></div></div>';
 }
 function dimensionPage(){
  return '<div class="lsw-main"><h1>Select Space Dimension</h1><div class="lsw-dimension-tiles" role="group" aria-label="Space Dimension">'+
   dimensions.map(function(d){return '<button type="button" class="lsw-dimension" data-dimension="'+d[0]+'" aria-pressed="'+(dimension===d[0])+'"><span class="lsw-dimension-icon">'+d[1]+'</span><span class="lsw-dimension-label">'+d[2]+'</span></button>';}).join("")+
   '</div><p class="lsw-muted" id="lsw-dimension-note">'+(dimension&&dimension!=="2D"?"해당 차원은 현재 브라우저 FEM 엔진에 연결되지 않습니다. 2D를 선택해 주세요.":"광섬유 횡단면 모드 해석을 시작하려면 2D를 선택하세요.")+'</p></div>';
 }
 function treeHtml(){
  const query=filter.trim().toLowerCase();
  const matches=function(s){return !query||s.toLowerCase().includes(query);};
  const showRay=matches("Ray Optics")||matches("Geometrical Ray Optics")||matches("Optics");
  const visibleInterfaces=interfaces.filter(function(x){return matches(x[1])||matches("Wave Optics")||query==="optics";});
  const showWave=visibleInterfaces.length>0;
  const any=showRay||showWave;
  return any?'<button type="button" class="lsw-cat" data-toggle="optics" aria-expanded="'+opticsOpen+'"><span class="lsw-cat-icon">'+(opticsOpen?"▾":"▸")+'</span>🔭　Optics</button>'+
    (opticsOpen?'<div class="lsw-branch">'+
      (showRay?'<button type="button" class="lsw-cat" data-toggle="ray" aria-expanded="'+rayOpen+'"><span class="lsw-cat-icon">'+(rayOpen?"▾":"▸")+'</span>◌　Ray Optics</button>'+
       (rayOpen?'<div class="lsw-branch"><p class="lsw-tree-future">Ray Optics 인터페이스는 후속 개발 예정입니다.</p></div>':""):"")+
      (showWave?'<button type="button" class="lsw-cat" data-toggle="wave" aria-expanded="'+waveOpen+'"><span class="lsw-cat-icon">'+(waveOpen?"▾":"▸")+'</span>≋　Wave Optics</button>'+
       (waveOpen?'<div class="lsw-branch" role="listbox" aria-label="Wave Optics interfaces">'+visibleInterfaces.map(function(x){
         return '<button type="button" class="lsw-leaf" data-interface="'+x[0]+'" aria-selected="'+(selected===x[0])+'"'+(x[2]?"":' disabled title="미구현 인터페이스"')+'><span style="color:#2187bd">≋</span><span>'+x[1]+(x[2]?"":' <small>· 예정</small>')+'</span></button>';
       }).join("")+'</div>':""):"")+
     '</div>':""):'<p class="lsw-tree-future">검색 결과가 없습니다.</p>';
 }
 function physicsPage(){
  return '<section class="lsw-physics-panel lsw-left" aria-label="Physics selection"><h1>Select Physics</h1>'+
    '<div class="lsw-search"><input id="lsw-search-input" type="search" aria-label="Search physics" placeholder="Search physics interfaces"/><button type="button" class="lsw-btn" data-wizard="search">Search</button></div>'+
    '<div class="lsw-tree" id="lsw-tree">'+treeHtml()+'</div>'+
    '<div class="lsw-list-actions"><button type="button" class="lsw-btn" data-wizard="add" '+(selected!=="ewfd"||added?"disabled":"")+'>Add</button></div>'+
    '<div class="lsw-label">Added physics interfaces:</div><div class="lsw-added" id="lsw-added">'+(added?"≋　Electromagnetic Waves, Frequency Domain (ewfd)":"")+'</div>'+
    '<div class="lsw-list-actions"><button type="button" class="lsw-btn" data-wizard="remove" '+(added?"":"disabled")+'>Remove</button></div>'+
    '<p class="lsw-inline-status" id="lsw-physics-status">'+(added?"ewfd 인터페이스가 추가되었습니다. Study로 진행하거나 Done을 눌러 기존 툴을 열 수 있습니다.":"Wave Optics 아래 ewfd를 선택한 뒤 Add를 눌러 주세요.")+'</p>'+
    '<div class="lsw-nav"><button type="button" class="lsw-btn" data-wizard="back">❮　Space Dimension</button>'+
    '<button type="button" class="lsw-btn" data-wizard="study" '+(added?"":"disabled")+'>Study　❯</button></div></section>'+
    '<section class="lsw-physics-preview" aria-label="Physics preview"><h2>Optical Fiber · Wave Optics</h2>'+
    '<p>2D 광섬유 횡단면의 광학 모드 분석을 위한 물리 인터페이스를 선택합니다. 현재 계산 엔진에 연결된 인터페이스는 electromagnetic waves, frequency domain (ewfd)입니다.</p>'+
    '<p><b>계산 범위 안내:</b> 브라우저의 기본 FEM 엔진은 스칼라 Radial FEM(LP₀₁ 근사)이며, COMSOL의 2D Full-vector ewfd 솔버와 동일하지 않습니다. 고급 벡터/HCF FEM은 기존 Colab 연동 기능으로 별도 제공됩니다.</p></section>';
 }
 function studyPage(){
  return '<div class="lsw-main"><h1>Select Study</h1><p class="lsw-muted">Selected interface: Electromagnetic Waves, Frequency Domain (ewfd)</p>'+
   '<div class="lsw-study-options"><label class="lsw-study-option"><input type="radio" checked name="lsw-study" value="mode" />'+
   '<span><b>Mode Analysis</b><small>기존 광섬유 LP₀₁ Radial FEM 모드 계산 및 n_eff, MFD, A_eff 결과</small></span></label></div>'+
   '<div class="lsw-summary"><strong>Model Wizard Summary</strong><div>Space Dimension: 2D</div><div>Physics: Wave Optics → Electromagnetic Waves, Frequency Domain (ewfd)</div><div>Study: Mode Analysis</div>'+
   '<div>Done을 누르면 기존 Geometry / Materials / Mesh / Study 워크벤치가 열립니다.</div></div>'+
   '<div class="lsw-nav" style="max-width:550px;margin-top:25px"><button type="button" class="lsw-btn" data-wizard="physics">❮　Select Physics</button></div></div>';
 }
 function render(){
  root.innerHTML=base();
  q("lsw-body").innerHTML=step==="dimension"?dimensionPage():step==="physics"?physicsPage():studyPage();
  if(step==="physics")q("lsw-search-input").value=filter;
  const done=root.querySelector('[data-wizard="done"]');
  if(done)done.disabled=step==="dimension"?dimension!=="2D":!added;
  q("lsw-footer-status").textContent=step==="dimension"?"1 / 3 · Space Dimension":step==="physics"?"2 / 3 · Physics":"3 / 3 · Study";
 }
 function show(){step="dimension";dimension="";selected="ewfd";added=false;filter="";opticsOpen=true;waveOpen=true;rayOpen=false;root.hidden=false;document.documentElement.classList.add("lsw-open");render();}
 function close(panel){
  root.hidden=true;root.replaceChildren();document.documentElement.classList.remove("lsw-open");
  if(!panel)return;
  document.dispatchEvent(new CustomEvent("ls-open-wave-optics",{detail:{panel:panel}}));
 }
 root.addEventListener("click",function(e){
  const dim=e.target.closest("[data-dimension]");
  if(dim){dimension=dim.dataset.dimension;if(dimension==="2D")step="physics";render();return;}
  const toggle=e.target.closest("[data-toggle]");
  if(toggle){
   if(toggle.dataset.toggle==="optics")opticsOpen=!opticsOpen;
   if(toggle.dataset.toggle==="ray")rayOpen=!rayOpen;
   if(toggle.dataset.toggle==="wave")waveOpen=!waveOpen;
   const input=q("lsw-search-input"),position=input?input.selectionStart:0;
   render();if(step==="physics"){q("lsw-search-input").focus();q("lsw-search-input").setSelectionRange(position,position);}return;
  }
  const item=e.target.closest("[data-interface]");
  if(item&&!item.disabled){selected=item.dataset.interface;render();return;}
  const btn=e.target.closest("[data-wizard]");if(!btn||btn.disabled)return;
  switch(btn.dataset.wizard){
   case "help":window.open("https://github.com/kimheeseo/LSCNS/tree/main/COMSOL/universal_silica_v01/web","_blank","noopener");return;
   case "cancel":close();return;
   case "back":step="dimension";render();return;
   case "search":filter=q("lsw-search-input").value;render();return;
   case "add":if(selected==="ewfd"&&dimension==="2D"){added=true;render();}return;
   case "remove":added=false;render();return;
   case "study":if(added){step="study";render();}return;
   case "physics":step="physics";render();return;
   case "done":
    if(step==="dimension"){if(dimension==="2D"){step="physics";render();}return;}
    if(added&&dimension==="2D"){close(step==="study"?"waveMode":"wavePhysics");}
    return;
  }
 });
 root.addEventListener("input",function(e){
  if(e.target.id!=="lsw-search-input")return;
  filter=e.target.value;
  const tree=q("lsw-tree");if(tree)tree.innerHTML=treeHtml();
 });
 root.addEventListener("keydown",function(e){if(e.key==="Escape"){e.preventDefault();close();}});
 // The existing New model ribbon action retains its reset handler and reopens the wizard.
 document.addEventListener("click",function(e){
  if(!root.hidden)return;
  const newModel=e.target.closest('.ls-ribbon [data-action="new"]');
  if(newModel)queueMicrotask(show);
 },true);
 show();
})();
