/* Shared sidebar inspector and non-overlapping tool dock. */
(function(){
'use strict';
const T=window.__LS3D_TEST__;if(!T)return;
const $=id=>document.getElementById(id);
const style=document.createElement('style');
style.textContent=`
.side{overflow-y:auto}.side .logo{flex:none}.side .side-section{flex:1 0 auto}.side .side-bottom{flex:none}
.inspector-tools{margin-top:16px;border-top:1px solid #315165;padding-top:14px}.inspector-tools h3{font-size:11px;color:#9fc6d8;margin:0 0 9px}.tool-dock{display:grid;gap:6px}
.tool-dock #phase2-bar{position:static!important;display:grid!important;grid-template-columns:1fr 1fr;gap:6px;max-width:none!important;margin:0!important;transform:none!important}
.tool-dock .phase58-launch,.tool-dock #phase2-bar button{position:static!important;inset:auto!important;display:block;width:100%;max-width:100%;margin:0!important;transform:none!important;padding:9px 10px;border:1px solid #456c81;border-radius:7px;background:#17394d;color:#d4edf8;font:600 11px system-ui;text-align:left;min-height:36px;white-space:normal}
#mobileInspectorTools{margin:4px 0;border:1px solid #365c73;border-radius:8px;background:#102b3e;padding:9px;color:#cae6f4;font-size:12px}#mobileInspectorTools summary{cursor:pointer}#mobileInspectorTools .tool-dock{margin-top:9px;grid-template-columns:1fr 1fr}
.inspector-selectors{display:grid;gap:10px;padding:13px 0 14px}.inspector-selectors label{display:grid;gap:6px;color:#9bbdd0;font-size:11px}.inspector-selectors select{width:100%;min-width:0;padding:9px;border:1px solid #4a7187;background:#112d42;border-radius:7px;color:#edf8ff;font-size:12px}.inspector-part{margin-bottom:13px;border:1px solid #458d9c;border-left:4px solid #f0c663;background:#123449;border-radius:9px;padding:12px}.inspector-part[hidden]{display:none}.inspector-part small{display:block;color:#9ac6d9;font-size:10px}.inspector-part strong{display:block;color:#f2f9fc;font-size:15px;line-height:1.5;margin:5px 0}.inspector-part p{font-size:12px;line-height:1.7;color:#c1dce8;margin:6px 0}.inspector-part button{border:1px solid #588399;border-radius:6px;background:#21465c;color:#e9f6fc;padding:6px 9px;font-size:11px;cursor:pointer}
.inspector-rack-open{width:100%;padding:10px;border:1px solid #63b8b3;border-radius:8px;background:#15515b;color:#e5fffb;font:700 12px system-ui;cursor:pointer}.inspector-rack-open[hidden]{display:none}
.inspector-rack-ops{border:1px solid #4d8e9c;border-radius:10px;background:#113145;padding:12px;margin:8px 0 14px;color:#e8f8ff}
.inspector-rack-ops[hidden],.inspector-rack-ops button[hidden]{display:none!important}
.inspector-rack-ops h3{font-size:14px;font-weight:800;margin:0 0 7px;color:#a9f6e8}
.inspector-rack-ops p{font-size:11px;line-height:1.65;color:#bed8e5;margin:0 0 11px}
.inspector-rack-ops-grid{display:grid;grid-template-columns:1fr 1fr;gap:7px}
.inspector-rack-ops-grid button{min-height:40px;font:700 11.5px "Pretendard Variable",Pretendard,"Malgun Gothic",system-ui,sans-serif;line-height:1.45;word-break:keep-all;background:#1d4b62;color:#f5fcff;border:1px solid #65a6b8;border-radius:7px;padding:7px 5px;cursor:pointer}
.inspector-rack-ops-grid button:hover{background:#2b687c;border-color:#95e9dd}
#inspectorRackOpStatus{font-size:11px;line-height:1.55;color:#a8d5e6;margin-top:9px}

body.visitor-tour .inspector-tools,body.visitor-tour #mobileInspectorTools{display:none!important}
@media(min-width:791px){#mobileInspectorTools{display:none!important}}
`;
document.head.appendChild(style);
const dockSection=document.createElement('section');dockSection.className='inspector-tools';dockSection.setAttribute('aria-label','검증 및 설계 도구');
dockSection.innerHTML='<h3>검증 · 설계 도구</h3><div class="tool-dock" id="ls3dToolDock"></div>';
$('zoneList').insertAdjacentElement('afterend',dockSection);
const mobile=document.createElement('details');mobile.id='mobileInspectorTools';mobile.innerHTML='<summary>검증 · 설계 도구</summary><div class="tool-dock" id="ls3dMobileToolDock"></div>';
$('viewport').insertAdjacentElement('beforebegin',mobile);
function moveTools(){
 const dock=innerWidth<=790?$('ls3dMobileToolDock'):$('ls3dToolDock');
 for(const id of ['phase2-bar','phase5-open','phase6-open','phase7-open','phase8-open']){const node=$(id);if(node&&node.parentElement!==dock)dock.appendChild(node);}
}
moveTools();window.addEventListener('resize',moveTools);
const selector=document.createElement('section');selector.className='inspector-selectors';
selector.innerHTML='<label for="inspectorAssetSelect">설비 선택<select id="inspectorAssetSelect" aria-label="설비 선택"><option value="">설비를 선택하세요</option></select></label><button id="inspectorRackOpen" class="inspector-rack-open" type="button" hidden>선택 랙 단독 보기 · 문 열기/닫기</button><label for="inspectorPartSelect">제품 · 부품 · 구조물 선택<select id="inspectorPartSelect" aria-label="부품 · 구조물 선택"><option value="">3D 화면에서 부품을 클릭하세요</option></select></label>';
document.querySelector('.right-head').insertAdjacentElement('afterend',selector);
const card=document.createElement('section');card.id='inspectorPartDetails';card.className='inspector-part';card.hidden=true;card.setAttribute('aria-live','polite');
card.innerHTML='<small id="inspectorPartContext"></small><strong id="inspectorPartName"></strong><p id="inspectorPartDescription"></p><button type="button" id="inspectorPartClear">부품 선택 해제</button>';
selector.insertAdjacentElement('afterend',card);
// Persistent drawer controls in the inspector: they cannot be lost behind
// the 3D canvas or a collapsed floating rack toolbar.
const rackOps=document.createElement('section');
rackOps.id='inspectorRackOperations';
rackOps.className='inspector-rack-ops';
rackOps.hidden=true;
rackOps.setAttribute('aria-label','랙 단독 보기 부품 열기 및 넣기');
rackOps.innerHTML='<h3>랙 내부 부품 열기 · 넣기</h3>'+
 '<p>선택 랙의 문을 연 다음, 서버·쉘프·GPU·광 카세트·드라이브를 앞으로 인출하거나 원위치로 넣을 수 있습니다.</p>'+
 '<div class="inspector-rack-ops-grid">'+
 '<button type="button" data-rack-quick="front-camera">전면 카메라</button>'+
 '<button type="button" data-rack-quick="rear-camera">후면 카메라</button>'+
 '<button type="button" data-rack-quick="front-door">전면문 열기 / 닫기</button>'+
 '<button type="button" data-rack-quick="rear-door">후면문 열기 / 닫기</button>'+
 '<button type="button" data-rack-quick="previous-server">이전 부품</button>'+
 '<button type="button" data-rack-quick="next-server">다음 부품</button>'+
 '<button type="button" data-rack-quick="server-slide">선택 서버 / 쉘프 인출 · 복귀</button>'+
 '<button type="button" data-rack-quick="cassette-slide">광 카세트 인출 · 복귀</button>'+
 '<button type="button" data-rack-quick="next-gpu" data-rack-quick-kind="gpu">다음 GPU 슬롯</button>'+
 '<button type="button" data-rack-quick="gpu-toggle" data-rack-quick-kind="gpu">GPU 모듈 인출 · 장착</button>'+
 '<button type="button" data-rack-quick="next-drive" data-rack-quick-kind="storage">다음 드라이브</button>'+
 '<button type="button" data-rack-quick="drive-toggle" data-rack-quick-kind="storage">드라이브 인출 · 장착</button>'+
 '</div><div id="inspectorRackOpStatus" role="status">랙 단독 보기에서 인출·복귀 버튼을 눌러 주세요.</div>';
card.insertAdjacentElement('afterend',rackOps);
rackOps.addEventListener('click',event=>{
 const code=event.target.closest?.('[data-rack-quick]')?.dataset.rackQuick;
 if(!code||rackOps.hidden||!T.rackIsolateActive)return;
 const adapter=window.LS3D_THREE?.adapter;
 const successful=!!adapter?.runRackAction?.(code);
 if(successful){
  adapter.expandRackControls?.();
  $('inspectorRackOpStatus').textContent=
    event.target.textContent+' · 3D 랙 조작을 실행했습니다. 같은 버튼을 다시 누르면 복귀합니다.';
 }else{
  $('inspectorRackOpStatus').textContent=
    '선택한 랙의 3D 엔진 또는 부품을 확인해 주세요. 단독 보기에서 사용 가능합니다.';
 }
});
let part=null,parts=[],assetSignature='',partSignature='',contextAsset=null;
const assetSelect=$('inspectorAssetSelect'),partSelect=$('inspectorPartSelect');
const zoneName=id=>T.Z.find(z=>z.id===id)?.n||id||'데이터센터';
function option(value,label){const o=document.createElement('option');o.value=value;o.textContent=label;return o;}
function renderPart(){
 card.hidden=!part;
 if(part){
  $('inspectorPartName').textContent=part.label||part.name||'선택 부품';
  $('inspectorPartDescription').textContent=part.description||'선택한 설비의 구성 부품입니다. 교육용 개념 모델입니다.';
  const parent=T.assetMap.get(part.assetId);
  $('inspectorPartContext').textContent=(parent?parent.name:zoneName(part.zoneId))+(Number.isFinite(part.uStart)?' · U '+part.uStart+'–'+(part.uStart+(part.uHeight||1)-1):'')+(part.component?' · '+part.component:'');
 }
 partSelect.value=part?.id||'';
}
function setSelection(info){part=info?{...info}:null;contextAsset=info?.assetId||T.selected?.id||null;renderPart();}
function setParts(next){
 parts=next||[];const sig=parts.map(p=>p.id+'|'+p.label).join('\n');if(sig===partSignature)return;partSignature=sig;
 partSelect.replaceChildren(option('',parts.length?'부품 또는 구조물을 선택하세요':'현재 선택 가능한 부품이 없습니다'));
 const groups=new Map();for(const p of parts){const label=p.kind==='equipment'?'선택 설비 설명':p.assetId?'선택 설비 내부 부품':'주변 트레이 · 구조물';if(!groups.has(label)){const group=document.createElement('optgroup');group.label=label;partSelect.appendChild(group);groups.set(label,group);}groups.get(label).appendChild(option(p.id,p.label+(p.assetId?'':' · '+zoneName(p.zoneId))));}
 if(part&&!parts.some(p=>p.id===part.id)&&part.kind!=='mep-route')setSelection(null);renderPart();
}
function sync(){
 const assets=T.visibleAssets().filter(a=>T.currentZone==='campus'||a.zone===T.currentZone);
 const sig=assets.map(a=>a.id).join('|');if(sig!==assetSignature){assetSignature=sig;assetSelect.replaceChildren(option('','설비를 선택하세요'));for(const a of assets)assetSelect.appendChild(option(a.id,a.name));}
 if(part&&contextAsset!==T.selected?.id)setSelection(null);
 assetSelect.value=part&&!part.assetId?'':T.selected?.id||'';
 const rackButton=$('inspectorRackOpen');rackButton.hidden=!T.isMechanicalRack(T.selected);
 rackButton.textContent=T.rackIsolateActive?'← 선택 랙 단독 보기 닫기':'▣ '+(T.selected?.name||'선택 랙')+' 단독 보기 · 문 열기/닫기';
 const active=!!(T.rackIsolateActive&&T.isMechanicalRack(T.selected));
 rackOps.hidden=!active;
 if(active){
  const gpu=/gpu/i.test(T.selected?.type||'');
  const storage=T.selected?.type==='Storage rack';
  for(const button of rackOps.querySelectorAll('[data-rack-quick-kind]')){
   button.hidden=button.dataset.rackQuickKind==='gpu'?!gpu:!storage;
  }
 }
}
assetSelect.addEventListener('change',()=>{const a=T.assetMap.get(assetSelect.value);if(!a)return;window.LS3D_THREE?.adapter?.clearInspectionSelection();setSelection(null);T.pickAsset(a,true);sync();});
$('inspectorRackOpen').addEventListener('click',()=>{if(T.rackIsolateActive)T.closeRackScene();else $('detailBtn')?.click();sync();});
partSelect.addEventListener('change',()=>{const A=window.LS3D_THREE?.adapter;if(partSelect.value)A?.selectInspectionPart(partSelect.value);else{A?.clearInspectionSelection();setSelection(null);}});
$('inspectorPartClear').addEventListener('click',()=>{window.LS3D_THREE?.adapter?.clearInspectionSelection();setSelection(null);});
window.LS3D_INSPECTOR={sync,setParts,setSelection,get selection(){return part;},get parts(){return parts.slice();}};
sync();
})();
