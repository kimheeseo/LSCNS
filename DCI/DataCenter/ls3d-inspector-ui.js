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
selector.innerHTML='<label for="inspectorAssetSelect">설비 선택<select id="inspectorAssetSelect" aria-label="설비 선택"><option value="">설비를 선택하세요</option></select></label><label for="inspectorPartSelect">부품 · 구조물 선택<select id="inspectorPartSelect" aria-label="부품 · 구조물 선택"><option value="">3D 화면에서 부품을 클릭하세요</option></select></label>';
document.querySelector('.right-head').insertAdjacentElement('afterend',selector);
const card=document.createElement('section');card.id='inspectorPartDetails';card.className='inspector-part';card.hidden=true;card.setAttribute('aria-live','polite');
card.innerHTML='<small id="inspectorPartContext"></small><strong id="inspectorPartName"></strong><p id="inspectorPartDescription"></p><button type="button" id="inspectorPartClear">부품 선택 해제</button>';
selector.insertAdjacentElement('afterend',card);
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
 const groups=new Map();for(const p of parts){const label=p.assetId?'선택 설비 내부 부품':'주변 트레이 · 구조물';if(!groups.has(label)){const group=document.createElement('optgroup');group.label=label;partSelect.appendChild(group);groups.set(label,group);}groups.get(label).appendChild(option(p.id,p.label+(p.assetId?'':' · '+zoneName(p.zoneId))));}
 if(part&&!parts.some(p=>p.id===part.id)&&part.kind!=='mep-route')setSelection(null);renderPart();
}
function sync(){
 const assets=T.visibleAssets().filter(a=>T.currentZone==='campus'||a.zone===T.currentZone);
 const sig=assets.map(a=>a.id).join('|');if(sig!==assetSignature){assetSignature=sig;assetSelect.replaceChildren(option('','설비를 선택하세요'));for(const a of assets)assetSelect.appendChild(option(a.id,a.name));}
 if(part&&contextAsset!==T.selected?.id)setSelection(null);
 assetSelect.value=part&&!part.assetId?'':T.selected?.id||'';
}
assetSelect.addEventListener('change',()=>{const a=T.assetMap.get(assetSelect.value);if(!a)return;window.LS3D_THREE?.adapter?.clearInspectionSelection();setSelection(null);T.pickAsset(a,true);sync();});
partSelect.addEventListener('change',()=>{const A=window.LS3D_THREE?.adapter;if(partSelect.value)A?.selectInspectionPart(partSelect.value);else{A?.clearInspectionSelection();setSelection(null);}});
$('inspectorPartClear').addEventListener('click',()=>{window.LS3D_THREE?.adapter?.clearInspectionSelection();setSelection(null);});
window.LS3D_INSPECTOR={sync,setParts,setSelection,get selection(){return part;},get parts(){return parts.slice();}};
sync();
})();
