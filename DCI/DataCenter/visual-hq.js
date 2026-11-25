/* Optional high-resolution imagery. The model, canvases and calculations stay authoritative. */
(function(){
  'use strict';
  const Core=window.DCVisualHQCore;if(!Core)return;
  const base=new URL('./assets/visual-hq/',document.baseURI);
  const escape=value=>String(value).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const url=r=>new URL(r.src,base).href;
  let entries=[],loaded=false,intent='asset',explicitPart=null,previewKey='',currentPreview=null;
  let dialog=null,returnFocus=null,viewerToken=0,queue=false,bomKey='',tourHost=null,previewObserver=null;
  const failed=new Set(),slotBindings=new WeakMap();
  function tourContext(){
    const T=window.__LS3D_TEST__;if(!T)return null;
    let zone=T.currentZone,asset=null,part=null,name,description;
    if(intent==='part'&&explicitPart){
      const info=explicitPart;if(typeof info.id!=='string'||!info.id)return null;
      asset=info.assetId?T.assetMap.get(info.assetId):null;
      if(asset&&asset.id!==T.selected?.id)return null;
      zone=asset?.zone||info.zoneId||zone;part=info.id;
      name=info.label||info.name;description=info.description||info.role||'';
    }else if(intent==='asset'){
      asset=T.selected;if(!asset)return null;
      zone=asset.zone;name=asset.name;description=[asset.desc,asset.details].filter(Boolean).join(' ');
    }else{
      const z=T.Z.find(z=>z.id===zone);name=z?.ko||z?.n;description=z?.sub;
    }
    if(!T.visibleZones().some(z=>z.id===zone)||asset&&!T.visibleAssets().some(a=>a.id===asset.id))return null;
    // Profiles and visitor extensions may update the host's shared geometry.
    // Read it afresh rather than retaining an earlier signature for that profile.
    return {scope:'tour',profile:T.facilityMode,composition:Core.composition(T),
      zoneId:zone,assetId:asset?.id||null,partId:part,name,description};
  }
  function ensureDialog(){
    if(dialog)return dialog;
    dialog=document.createElement('dialog');dialog.className='hq-dialog';dialog.id='visualHQDialog';
    dialog.setAttribute('aria-labelledby','visualHQTitle');dialog.setAttribute('aria-describedby','visualHQDescription');
    dialog.innerHTML='<div class="hq-heading"><h2 id="visualHQTitle"></h2><button type="button" data-hq-close aria-label="이미지 닫기">닫기</button></div>'+
      '<div class="hq-image-wrap"><img alt="" hidden><p data-hq-image-status role="status"></p></div>'+
      '<div class="hq-caption"><p id="visualHQDescription"></p><small>기준 시점 이미지 · 실시간 조작 상태는 3D 화면에서 확인하세요.</small><a target="_blank" rel="noopener" hidden>4K 원본 열기</a></div>';
    document.body.appendChild(dialog);
    dialog.querySelector('[data-hq-close]').addEventListener('click',()=>dialog.close());
    dialog.addEventListener('click',event=>{if(event.target===dialog)dialog.close();});
    dialog.addEventListener('cancel',event=>{event.preventDefault();dialog.close();});
    // Capture Escape before the host's rack/visitor shortcuts. Closing an image
    // must not also close a rack or stop a guided visitor session.
    window.addEventListener('keydown',event=>{
      if(dialog.open&&event.key==='Escape'){event.preventDefault();event.stopImmediatePropagation();dialog.close();}
    },true);
    dialog.addEventListener('close',()=>{
      viewerToken++;const image=dialog.querySelector('img');image.removeAttribute('src');image.removeAttribute('srcset');
      if(returnFocus?.isConnected)returnFocus.focus({preventScroll:true});
      returnFocus=null;
    });
    return dialog;
  }
  function openViewer(entry,context,trigger){
    const d=ensureDialog(),token=++viewerToken,img=d.querySelector('img'),state=d.querySelector('[data-hq-image-status]');
    returnFocus=trigger||document.activeElement;
    d.querySelector('h2').textContent=context.name||context.diagram?.nodeKey||'데이터센터';
    d.querySelector('#visualHQDescription').textContent=context.description||'';
    img.hidden=true;img.alt=context.name||context.diagram?.nodeKey||'데이터센터';
    state.hidden=false;state.textContent='이미지를 불러오는 중…';
    img.width=entry.preview.width;img.height=entry.preview.height;img.decoding='async';
    img.sizes='(max-width: 700px) 94vw, 1060px';
    img.onload=()=>{if(token!==viewerToken)return;img.hidden=false;state.hidden=true;};
    img.onerror=()=>{if(token!==viewerToken)return;img.hidden=true;state.textContent='이미지를 불러올 수 없습니다. 기존 상세 화면을 이용해 주세요.';};
    // Full resolution is never in srcset: it downloads only on an explicit click.
    img.srcset=[entry.thumbnail,entry.preview].map(r=>url(r)+' '+r.width+'w').join(', ');
    img.src=url(entry.preview);
    const link=d.querySelector('a');link.href=url(entry.full);link.hidden=false;
    link.textContent=Math.max(entry.width,entry.height)>=3840?'4K 원본 열기':'원본 열기';
    if(!d.open)d.showModal();
    d.querySelector('[data-hq-close]').focus({preventScroll:true});
  }
  function removePreview(){
    if(!tourHost)return;
    previewObserver?.disconnect();previewObserver=null;
    tourHost.querySelector('.hq-trigger')?.remove();
    for(const node of tourHost.children)if(node.dataset.hqHidden==='true'){node.hidden=false;delete node.dataset.hqHidden;}
    currentPreview=null;
  }
  function syncTour(){
    if(!loaded||!tourHost)return;
    const context=tourContext();
    tourHost.dataset.hqComposition=context?.composition||'';
    const entry=context&&Core.resolve(entries,context);
    const key=entry&&!failed.has(url(entry.preview))?JSON.stringify([context,entry.preview.src]):'';
    if(key===previewKey&&(!key||tourHost.querySelector('.hq-trigger')))return;
    previewKey=key;removePreview();if(!key)return;
    const button=document.createElement('button');button.type='button';button.className='hq-trigger';
    button.setAttribute('aria-label',(context.name||'선택 항목')+' 고화질 이미지 보기');
    const img=document.createElement('img');img.className='hq-preview';img.alt=context.name||'';
    img.width=entry.preview.width;img.height=entry.preview.height;img.decoding='async';
    // The button stays hidden until decode succeeds. Native loading=lazy would
    // never load an image in a display:none parent; observe the visible host.
    img.loading='eager';
    img.sizes='(max-width: 790px) 90vw, 280px';
    const label=document.createElement('small');label.textContent='이미지 확대';
    button.append(img,label);button.hidden=true;
    img.onload=()=>{
      if(previewKey!==key||!button.isConnected)return;
      for(const node of tourHost.children)if(node!==button&&!node.hidden){node.hidden=true;node.dataset.hqHidden='true';}
      button.hidden=false;
    };
    img.onerror=()=>{if(previewKey!==key)return;failed.add(url(entry.preview));previewKey='';removePreview();};
    button.addEventListener('click',()=>openViewer(entry,context,button));
    tourHost.appendChild(button);currentPreview={entry,context};
    const load=()=>{
      if(previewKey!==key||!button.isConnected)return;
      img.srcset=[entry.thumbnail,entry.preview].map(r=>url(r)+' '+r.width+'w').join(', ');
      img.src=url(entry.thumbnail);
    };
    if(typeof IntersectionObserver==='function'){
      previewObserver=new IntersectionObserver(changes=>{
        if(previewKey!==key)return;
        if(changes.some(change=>change.isIntersecting)){previewObserver?.disconnect();previewObserver=null;load();}
      },{rootMargin:'120px'});previewObserver.observe(tourHost);
    }else load();
  }
  function svgSlot(svgId,nodeKey,x,y,width,height){
    if(width<1||height<1)return '';
    return '<g data-hq-slot="'+escape(nodeKey)+'" data-hq-svg="'+escape(svgId)+'" data-hq-x="'+x+'" data-hq-y="'+y+'" data-hq-width="'+width+'" data-hq-height="'+height+'"></g>';
  }
  const observer=typeof IntersectionObserver==='function'?new IntersectionObserver(changes=>{
    for(const change of changes)if(change.isIntersecting){
      observer.unobserve(change.target);
      for(const slot of change.target.querySelectorAll('[data-hq-slot]'))loadSlot(slot);
    }
  },{rootMargin:'120px'}):null;
  function loadSlot(slot){
    const binding=slotBindings.get(slot);if(!binding||binding.loaded||!slot.isConnected)return;
    binding.loaded=true;
    const {entry,context}=binding,img=document.createElementNS('http://www.w3.org/2000/svg','image');
    for(const attr of ['x','y','width','height'])img.setAttribute(attr,slot.dataset['hq'+attr[0].toUpperCase()+attr.slice(1)]);
    img.setAttribute('preserveAspectRatio','xMidYMid meet');img.style.visibility='hidden';
    img.addEventListener('load',()=>{img.style.visibility='';});
    img.addEventListener('error',()=>{img.remove();slot.removeAttribute('tabindex');slot.removeAttribute('role');slot.style.cursor='';});
    img.setAttribute('href',url(entry.thumbnail));slot.appendChild(img);
    slot.setAttribute('role','button');slot.setAttribute('tabindex','0');slot.setAttribute('aria-label',context.name+' 이미지 보기');
    slot.style.cursor='zoom-in';
  }
  function syncBOM(){
    if(!loaded)return;
    const result=window.DCDesign;if(!result?.usable)return;
    const composition=Core.bomComposition(result);bomKey=composition;
    for(const slot of document.querySelectorAll('[data-hq-slot]')){
      if(slotBindings.get(slot)?.key===composition)continue;
      slot.replaceChildren();slot.removeAttribute('role');slot.removeAttribute('tabindex');slot.style.cursor='';
      const [label,ru]=slot.dataset.hqSlot.split('|');
      const context={scope:'bom',profile:result.input.systemId,composition,
        diagram:{svgId:slot.dataset.hqSvg,nodeKey:slot.dataset.hqSlot},name:label,
        description:ru?'랙 위치 '+ru+'U':''};
      const entry=Core.resolve(entries,context);slotBindings.set(slot,{key:composition,entry,context,loaded:false});
      if(!entry)continue;
      // Empty SVG groups have zero bounds; observe their existing diagram.
      if(observer)observer.observe(slot.ownerSVGElement);else loadSlot(slot);
    }
  }
  function scheduleBOM(){
    if(queue)return;queue=true;queueMicrotask(()=>{queue=false;syncBOM();});
  }
  window.DCVisualHQ={
    svgSlot,syncTour,
    selectZone(){intent='zone';explicitPart=null;},
    selectAsset(){intent='asset';explicitPart=null;},
    selectPart(info){intent=info?'part':'asset';explicitPart=info?{...info}:null;syncTour();},
    get diagnostics(){return {loaded,readyAssets:entries.length,hasPreview:!!currentPreview,bomComposition:bomKey};},
    captureContext(){return tourContext()||window.DCDesign&&{scope:'bom',profile:window.DCDesign.input.systemId,composition:Core.bomComposition(window.DCDesign)};}
  };
  async function init(){
    tourHost=document.getElementById('assetVisual');
    // No image fetches until the manifest supplies an approved, exact match.
    try{
      const response=await fetch(new URL('manifest.json?v='+Core.REVISION,base),{cache:'no-cache'});
      if(response.ok)entries=Core.normalize(await response.json());
    }catch(_){entries=[];}
    loaded=true;syncTour();syncBOM();
    for(const id of ['rackSvg','coolingSvg','powerArchitectureSvg']){
      const node=document.getElementById(id);if(node)new MutationObserver(scheduleBOM).observe(node,{childList:true});
    }
    document.addEventListener('dc:design',scheduleBOM);window.addEventListener('dc-design-updated',scheduleBOM);
    document.addEventListener('click',event=>{
      const slot=event.target.closest?.('[data-hq-slot]'),binding=slot&&slotBindings.get(slot);
      if(binding?.entry&&slot.getAttribute('role')==='button')openViewer(binding.entry,binding.context,slot);
    });
    document.addEventListener('keydown',event=>{
      const slot=event.target.closest?.('[data-hq-slot]'),binding=slot&&slotBindings.get(slot);
      if(binding?.entry&&slot.getAttribute('role')==='button'&&['Enter',' '].includes(event.key)){
        event.preventDefault();openViewer(binding.entry,binding.context,slot);
      }
    });
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init,{once:true});else init();
})();
