/* Optical connection launcher. The MEP work checklist has been removed. */
(function(){
'use strict';
const view=document.getElementById('viewport');
if(!window.__LS3D_TEST__||!view)return;
const bar=document.createElement('span');
bar.className='phase2-buttons-top';bar.id='phase2-bar';
bar.innerHTML='<button type="button" id="phase2-opt-open" aria-controls="phase2-opt-panel" aria-expanded="false">◎ 광 연결</button>';
view.appendChild(bar);
})();
