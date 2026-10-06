(() => {
  'use strict';
  const M=window.DCBomCatalog, $=id=>document.getElementById(id);
  const esc=x=>String(x??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const fmt=x=>x==null?'미산정':Number(x).toLocaleString('ko-KR',{maximumFractionDigits:2});
  const isLsVendor=p=>/LS\s*CABLE\s*&\s*SYSTEM|LS전선|LS CNS/i.test(String(p?.vendor||''));
  const productName=p=>{
    if(/COMMSCOPE/i.test(String(p?.vendor||''))){
      const s=p?.specs||{}, part=String(s['Part Number']||p?.id||'').trim();
      const type=String(s['Product Type']||p?.category||'').trim();
      return [part,type].filter(Boolean).join(' · ')||String(p?.name||'');
    }
    return String(p?.name||'');
  };
  const link=(p,label)=>'<div class="bom-candidate"><strong>'+esc(label)+'</strong><span class="bom-candidate-name">'+esc(p.vendor+' · '+productName(p))+'</span><a target="_blank" rel="noopener" href="'+esc(p.source)+'">제품 링크 ↗</a></div>';
  function diversifyCandidates(list){
    const seen=new Set(), uniq=[];
    for(const p of list||[]){
      const key=[p.vendor,p.id,p.source].join('|');
      if(!seen.has(key)){seen.add(key);uniq.push(p);}
    }
    if(uniq.length<2)return uniq;
    const first=uniq[0], top=[first], used=new Set([first]);
    const nearEnough=p=>Number(p?.score??-999)>=Number(first?.score??0)-10;
    const ls=uniq.slice(1).find(p=>isLsVendor(p)&&nearEnough(p)&&p.evidence!=='품목 분류 일치');
    if(!isLsVendor(first)&&ls){top.push(ls);used.add(ls);}
    const different=uniq.slice(1).find(p=>p.vendor!==first.vendor&&nearEnough(p)&&!used.has(p))
      ||uniq.slice(1).find(p=>p.vendor!==first.vendor&&!used.has(p));
    if(top.length<3&&different){top.push(different);used.add(different);}
    const topVendors=()=>new Set(top.map(p=>p.vendor));
    for(const p of uniq){
      if(top.length>=3)break;
      if(!used.has(p)&&!topVendors().has(p.vendor)){top.push(p);used.add(p);}
    }
    for(const p of uniq){
      if(top.length>=3)break;
      if(!used.has(p)&&!(isLsVendor(first)&&isLsVendor(p))){top.push(p);used.add(p);}
    }
    for(const p of uniq){if(top.length>=3)break;if(!used.has(p)){top.push(p);used.add(p);}}
    const remaining=uniq.filter(p=>!used.has(p));
    const extra=[], extraVendors=new Set(top.map(p=>p.vendor));
    for(const p of remaining){if(extra.length>=2)break;if(!extraVendors.has(p.vendor)){extra.push(p);extraVendors.add(p.vendor);}}
    for(const p of remaining){if(extra.length>=2)break;if(!extra.includes(p))extra.push(p);}
    const head=[...top,...extra];
    return [...head,...remaining.filter(p=>!extra.includes(p))];
  }
  let current=null, revision=0, data=null, vendor='', search='', ready=Promise.resolve();
  function rowsFor(r,a) {
    if(a) {
      const compute=(r.networks||[]).filter(x=>x.network==='compute');
      const computeProtocol=compute[0]?.protocol||'';
      const computeSpeed=compute[0]?.speed||'';
      const rackRu=r.assumptions?.find(x=>x.path?.endsWith('.rackLimitRu'))?.value;
      const bomRows=a.bom.map(b=>{
        const selection=a.selections.find(x=>x.network===b.network&&x.tier+' / '+x.distanceClass===b.segment);
        let requirement=b.requirement||(selection?{connector:selection.connector,fiberType:selection.fiberType||selection.fiber,protocol:selection.protocol}:undefined);
        // Patch panels are selected by fiber family first; exact adapter layout is finalized with the module/cassette.
        if(b.item==='패치패널') requirement={fiberType:selection?.fiberType||selection?.fiber||''};
        const fibers=/패치리드|광 케이블/.test(b.item)?selection?.installedFibers:undefined;
        return {...b,requirement,fibers,segment:b.network+' / '+b.segment,installed:b.installedQty,purchase:b.purchaseQty,candidates:[]};
      });
      const moduleRows=a.bom.filter(b=>b.item==='패치패널').map(b=>{
        const selection=a.selections.find(x=>x.network===b.network&&x.tier+' / '+x.distanceClass===b.segment);
        const fiber=selection?.fiberType||selection?.fiber||b.media||'';
        const connector=String(selection?.connector||'');
        const lcPath=/\bLC\b/i.test(connector);
        const group=a.bom.find(x=>x.item==='커넥터 종단 그룹'&&x.network===b.network&&x.segment===b.segment);
        const terminations=Number(group?.installedQty||0);
        const moduleCapacity=6; // Corning EDGE 12F module = 6 LC duplex adapter positions
        const installed=lcPath&&terminations?Math.ceil(terminations/moduleCapacity):0;
        const spareRatio=b.installedQty?Number(b.purchaseQty||b.installedQty)/Number(b.installedQty):1;
        const purchase=installed?Math.ceil(installed*spareRatio):0;
        const uncertainty=Number(b.uncertaintyPct||0);
        return {item:'광 Module / Cassette',segment:b.network+' / '+b.segment,installed,purchase,low:installed?Math.max(0,Math.floor(purchase*(1-uncertainty/100))):0,high:installed?Math.ceil(purchase*(1+uncertainty/100)):0,unit:'개',referenceOnly:true,
          spec:lcPath
            ? '12F EDGE module 계획 · '+fiber+' · 6×LC duplex/module · MTP/MPO rear · 현재 종단 '+terminations+'그룹 기준 · 실제 housing slot/극성/loss budget 확인'
            : '현재 '+(connector||'MTP/MPO')+' 직접 경로에서는 LC conversion module 0개 · LC breakout 대안 적용 시 '+fiber+' EDGE module 후보 검토',
          requirement:{fiberType:fiber},candidates:[]};
      });
      const refRows=r.references.map(x=>{
        let spec=x.note||'', requirement;
        if(x.item==='IT 랙') spec='IT rack enclosure · '+(rackRu?rackRu+'U · ':'')+r.summary.rackKw+' kW/rack 계획값(장비 정격과 구분)';
        if(x.item==='Leaf'){spec='Leaf switch · '+computeSpeed+'G · '+computeProtocol+' · down/up 포트 구성 검증';requirement={protocol:computeProtocol};}
        if(x.item==='Spine'){spec='Spine switch · '+computeSpeed+'G · '+computeProtocol+' · fabric/core 포트 구성 검증';requirement={protocol:computeProtocol};}
        if(x.item==='Super-spine'){spec='Core/Super-spine switch · '+computeSpeed+'G · '+computeProtocol+' · core fabric 포트 구성 검증';requirement={protocol:computeProtocol};}
        return {item:x.item,segment:'설계 참고 / '+x.item,installed:x.qty,purchase:null,unit:x.unit,referenceOnly:true,spec,requirement,candidates:[]};
      });
      const facilityRows=['UPS','Generator','Transformer'].map(item=>({item,segment:'Facility Power / '+item,installed:null,purchase:null,unit:'대',referenceOnly:true,
        spec:fmt(r.summary.facilityKw/1000)+' MW 시설부하 기준 · '+item+' 용량/수량·N+1/2N·전압/주파수·현장 조건 별도 sizing',candidates:[]}));
      const hasTransceiver=bomRows.some(b=>/트랜시버/.test(String(b.item||'')));
      const transceiverRows=hasTransceiver?[]:(()=>{
        const ratio=(()=>{
          const q=bomRows.find(b=>Number(b.installed)>0&&Number(b.purchase)>=Number(b.installed));
          return q?Number(q.purchase)/Number(q.installed):1;
        })();
        return (a.selections||[]).filter(s=>s.network!=='management'&&Number(s.speed)>0&&Number(s.links)>0).map(s=>{
          const installed=2*Number(s.links);
          const purchase=Math.ceil(installed*ratio);
          const uncertainty=20;
          return {item:'플러그형 트랜시버 (광 대안)',segment:(s.network||'network')+' / '+(s.tier||'link')+' / '+(s.distanceClass||''),installed,purchase,low:Math.floor(purchase*(1-uncertainty/100)),high:Math.ceil(purchase*(1+uncertainty/100)),unit:'개',referenceOnly:true,transceiver:true,
            spec:Number(s.speed||0)+'G · '+(s.protocol||'네트워크')+' · 현재 '+(s.media||'직접연결')+' 대신 SMF/MMF 플러그형 광링크 적용 시 양단 트랜시버 수량 · 현재 케이블과 중복 구매 금지',
            requirement:{speed:Number(s.speed||0)||undefined,protocol:s.protocol||''},candidates:[]};
        });
      })();
      return [...bomRows,...transceiverRows,...moduleRows,...refRows,...facilityRows];
    }
    const rows=r.productRequirements?.length?r.productRequirements:r.bom?.map(b=>({...b,profile:r.optical?.[b.segment]?.profile,installed:b.installedQty??b.qty,purchase:b.qty,candidates:[]}))||[];
    return rows.map(b=>({...b,candidates:[]}));
  }
  function aggregate(rows) {
    const map=new Map();for(const r of rows){const key=[r.segment,r.item,r.media,r.spec,JSON.stringify(r.requirement||r.profile),r.lengthM,r.unit,r.referenceOnly].join('|');if(!map.has(key)){map.set(key,{...r,phases:r.phase?[r.phase]:[]});continue;}const x=map.get(key);for(const k of ['installed','purchase','low','high'])if(r[k]!=null)x[k]=(x[k]||0)+r[k];if(r.phase&&!x.phases.includes(r.phase))x.phases.push(r.phase);}return [...map.values()];
  }
  function host() {
    const view=$('view-bom');if(!view)return null;
    let h=$('bom-live-products');if(!h){h=document.createElement('section');h.id='bom-live-products';h.className='panel';view.prepend(h);}
    // Keep a separate host: capacity rendering locates its own .cap-content.
    h.classList.remove('cap-content');
    h.style.setProperty('display','block','important');
    // Both modes use this one product table. Hide the old static shortlist.
    const old=$('design-products');if(old)old.hidden=true;
    return h;
  }
  function paint() {
    const h=host();if(!h||!current)return;
    const {rows,r}=current;
    const shown=rows.filter(x=>!search||[x.item,x.segment,x.spec].join(' ').toLowerCase().includes(search.toLowerCase())).map(x=>({...x,candidates:diversifyCandidates(x.candidates.filter(p=>!vendor||p.vendor===vendor))}));
    const vendors=[...new Set(data.products.map(x=>x.vendor))].sort((a,b)=>a.localeCompare(b));
    h.innerHTML='<h2>설계 제품·케이블 품목 리스트</h2><p id="bom-catalog-status" role="status">업로드 카탈로그 '+data.catalogs.length+'개 · 제품 '+data.products.length+'개 · GitHub '+esc(data.revision.slice(0,7))+' 기준'+(data.snapshot?' · 배포 스냅샷':'')+(data.errors.length?' · 일부 조회 실패 '+data.errors.length+'개':' · 동기화 완료')+'</p><p>현재 설계 조건에 맞는 제품을 기본 3개 표시합니다. 가장 적합한 제품 1개와 적합 후보 2개를 바로 보여주며, 추천 후보 버튼을 누르면 유사 제품을 포함해 최대 5개까지 확인할 수 있습니다. 가격: 산정 불가.</p><div class="cap-actions"><label>후보 업체 <select id="bom-catalog-vendor"><option value="">전체 업체</option>'+vendors.map(v=>'<option'+(vendor===v?' selected':'')+'>'+esc(v)+'</option>').join('')+'</select></label><label>품목 / 규격 검색 <input id="bom-item-search" type="search" value="'+esc(search)+'"></label><button type="button" id="bom-catalog-refresh">제품 목록 동기화</button><button type="button" id="exportProductCsv">제품 요구규격 CSV</button></div><div class="tableWrap bom-live-wrap"><table id="bom-live-table" class="bom-live-table"><thead><tr><th>구간 / 품목</th><th>설치 / 구매 범위</th><th>요구 규격</th><th>추천 제품 · 링크</th></tr></thead><tbody>'+shown.map((x,rowIndex)=>'<tr><td>'+esc(x.segment||'설계 참고')+' · '+esc(x.item)+(x.referenceOnly?' · <span>참고·중복 구매 제외</span>':'')+'</td><td>'+fmt(x.installed)+' / '+(x.low!=null?fmt(x.low)+'–'+fmt(x.high):fmt(x.purchase))+' '+esc(x.unit||'')+(x.phases.length?' · Phase '+x.phases.join(', ')+' 합계':'')+'</td><td>'+esc(x.spec||[x.profile?.media,x.profile?.connector].filter(Boolean).join(' · '))+(x.lengthM!=null?' · '+fmt(x.lengthM)+'m':'')+'</td><td>'+(x.candidates.length?(link(x.candidates[0],'가장 적합한 제품')+(x.candidates.slice(1,3).length?x.candidates.slice(1,3).map(p=>link(p,'적합 후보')).join(''):'')+(x.candidates.length>3?'<button type="button" class="bom-candidate-toggle" data-target="bom-candidates-'+rowIndex+'">추천 후보 더보기</button><div id="bom-candidates-'+rowIndex+'" class="bom-candidate-more" hidden>'+x.candidates.slice(3,5).map(p=>link(p,'유사 제품')).join('')+'</div>':'')):'추천 제품 없음')+'</td></tr>').join('')+'</tbody></table></div>'+(data.errors.length?'<details><summary>조회 실패 파일</summary>'+data.errors.map(x=>'<p>'+esc(x.path+': '+x.error)+'</p>').join('')+'</details>':'');
    $('bom-item-search').oninput=e=>{search=e.target.value;};
    $('bom-item-search').onchange=e=>{search=e.target.value;paint();$('bom-item-search').focus();};
    $('bom-catalog-vendor').onchange=e=>{vendor=e.target.value;paint();};
    $('bom-catalog-refresh').onclick=()=>refresh(true);
    $('exportProductCsv').onclick=()=>downloadCsv();
    h.querySelectorAll('.bom-candidate-toggle').forEach(btn=>btn.onclick=()=>{
      const box=document.getElementById(btn.dataset.target);if(!box)return;
      box.hidden=!box.hidden;btn.textContent=box.hidden?'추천 후보 더보기':'추천 후보 닫기';
    });
    r.catalogRows=rows;r.catalogSnapshot={revision:data.revision,catalogs:data.catalogs.length,products:data.products.length,errors:data.errors};
  }
  async function refresh(force=false) {
    if(!current)return;
    current.rows=[];current.r.catalogRows=[];current.r.catalogCandidates=[];
    if(current.r.productRequirements)for(const row of current.r.productRequirements)row.candidates=[];
    const token=++revision,h=host();if(h)h.innerHTML='<h2>설계 제품·케이블 품목 리스트</h2><p id="bom-catalog-status" role="status">업로드된 전체 카탈로그를 동기화하는 중…</p>';
    ready=(async()=>{try{const loaded=await M.load(force);if(token!==revision)return;data=loaded;const raw=rowsFor(current.r,current.a);for(const x of raw)x.candidates=diversifyCandidates(M.match(x,data.products));current.rows=aggregate(raw);if(current.a){current.r.catalogCandidates=current.rows.map(x=>({segment:x.segment||x.item,spec:x.spec||'',candidates:x.candidates}));}else if(current.r.productRequirements){for(let i=0;i<current.r.productRequirements.length;i++)current.r.productRequirements[i].candidates=raw[i].candidates;}paint();}catch(e){if(token!==revision)return;current.r.catalogRows=[];if(h)h.innerHTML='<h2>설계 제품·케이블 품목 리스트</h2><p role="alert">'+esc(e.message)+' · 제품 연결 미완료. 이전 후보를 재사용하지 않습니다.</p><button id="bom-catalog-retry">다시 동기화</button>';if($('bom-catalog-retry'))$('bom-catalog-retry').onclick=()=>refresh(true);}})();
    return ready;
  }
  const headers=['구간','품목','설치','구매 기준','하한','상한','단위','요구 규격','추천 구분','벤더','제품','공식 URL','카탈로그 경로','카탈로그 Git SHA'];
  function exportRows() {return (current?.rows||[]).flatMap(x=>{const cs=x.candidates.length?diversifyCandidates(x.candidates).slice(0,5):[{}];return cs.map((p,i)=>[x.segment||'',x.item,x.installed??'',x.purchase??'',x.low??'',x.high??'',x.unit,x.spec||x.profile?.media||'',p.name?(i===0?'가장 적합한 제품':i<3?'적합 후보':'유사 제품'):'추천 제품 없음',p.vendor||'',p.name||'',p.source||'',p.path||'',data?.revision||'']);});}
  async function downloadCsv() {await ready;if(!current?.rows)return;const csv=[headers,...exportRows()].map(row=>row.map(v=>'"'+String(v??'').replace(/^([=+@-])/,'\t$1').replace(/"/g,'""')+'"').join(',')).join('\r\n');const url=URL.createObjectURL(new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8'})),a=document.createElement('a');a.href=url;a.download='datacenter-catalog-product-requirements.csv';document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),10000);}
  document.addEventListener('dc:capacity-bom',e=>{current={r:e.detail.r,a:e.detail.a,rows:[]};refresh();});
  document.addEventListener('dc:design',e=>{if(window.DCCapacityDesign)return;current=e.detail?.usable?{r:e.detail,a:null,rows:[]}:null;requestAnimationFrame(()=>{if(current)refresh();else{const h=$('bom-live-products');if(h)h.remove();}});});
  document.addEventListener('dc:catalog-refresh',()=>refresh(true));
  document.addEventListener('dc:capacity-invalid',()=>{revision++;current=null;const h=$('bom-live-products');if(h)h.remove();});
  document.addEventListener('dc:stale',()=>{revision++;current=null;const h=$('bom-live-products');if(h)h.innerHTML='<h2>설계 제품·케이블 품목 리스트</h2><p>요구조건 변경 · 설계안을 다시 계산하세요.</p>';});
  window.DCBomCatalogUI={get ready(){return ready;},refresh,exportRows,headers};
})();
