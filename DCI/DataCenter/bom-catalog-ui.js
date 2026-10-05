(() => {
  'use strict';
  const M=window.DCBomCatalog, $=id=>document.getElementById(id);
  const esc=x=>String(x??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const fmt=x=>x==null?'미산정':Number(x).toLocaleString('ko-KR',{maximumFractionDigits:2});
  const link=p=>'<a target="_blank" rel="noopener" href="'+esc(p.source)+'">'+esc(p.vendor+' · '+p.name)+'</a><small style="display:block">'+esc(p.status+' / '+p.matchBasis)+'</small>';
  let current=null, revision=0, data=null, vendor='', ready=Promise.resolve();
  function rowsFor(r,a) {
    if(a) return [...a.bom.map(b=>({...b,segment:b.network+' / '+b.segment,installed:b.installedQty,purchase:b.purchaseQty,candidates:[]})),...r.references.map(x=>({item:x.item,installed:x.qty,purchase:null,unit:x.unit,referenceOnly:true,spec:x.note,candidates:[]})),...['UPS','Generator','Transformer'].map(item=>({item,installed:null,purchase:null,unit:'대',referenceOnly:true,spec:'개별 설비 용량/수량·이중화·현장 조건 별도 설계',candidates:[]}))];
    const rows=r.productRequirements?.length?r.productRequirements:r.bom?.map(b=>({...b,profile:r.optical?.[b.segment]?.profile,installed:b.installedQty??b.qty,purchase:b.qty,candidates:[]}))||[];
    return rows.map(b=>({...b,candidates:[]}));
  }
  function aggregate(rows) {
    const map=new Map();for(const r of rows){const key=[r.segment,r.item,r.media,r.spec,JSON.stringify(r.requirement||r.profile),r.lengthM,r.unit,r.referenceOnly].join('|');if(!map.has(key)){map.set(key,{...r,phases:r.phase?[r.phase]:[]});continue;}const x=map.get(key);for(const k of ['installed','purchase','low','high'])if(r[k]!=null)x[k]=(x[k]||0)+r[k];if(r.phase&&!x.phases.includes(r.phase))x.phases.push(r.phase);}return [...map.values()];
  }
  function host() {
    const view=$('view-bom');if(!view)return null;
    let h=$('bom-live-products');if(!h){h=document.createElement('section');h.id='bom-live-products';h.className='panel';view.prepend(h);}
    h.classList.toggle('cap-content',!!current?.a);
    // Both modes use this one product table. Hide the old static shortlist.
    const old=$('design-products');if(old)old.hidden=true;
    return h;
  }
  function paint() {
    const h=host();if(!h||!current)return;
    const {rows,r}=current;
    const shown=rows.map(x=>({...x,candidates:x.candidates.filter(p=>!vendor||p.vendor===vendor)}));
    const vendors=[...new Set(data.products.map(x=>x.vendor))].sort((a,b)=>a.localeCompare(b));
    h.innerHTML='<h2>설계 제품·케이블 품목 리스트</h2><p id="bom-catalog-status" role="status">업로드 카탈로그 '+data.catalogs.length+'개 · 제품 '+data.products.length+'개 · GitHub '+esc(data.revision.slice(0,7))+' 기준'+(data.errors.length?' · 일부 조회 실패 '+data.errors.length+'개':' · 동기화 완료')+'</p><p>현재 설계 수량에 등록 제품을 연결합니다. 규격이 확인된 후보와 분류만 일치하는 관련 제품을 구분합니다. 제품 등록은 호환성 인증을 의미하지 않으며, 관련 제품·설계 참고 항목은 구매량에 추가하지 않습니다. 가격: 산정 불가.</p><div class="cap-actions"><label>후보 업체 <select id="bom-catalog-vendor"><option value="">전체 업체</option>'+vendors.map(v=>'<option'+(vendor===v?' selected':'')+'>'+esc(v)+'</option>').join('')+'</select></label><button type="button" id="bom-catalog-refresh">제품 목록 동기화</button><button type="button" id="exportProductCsv">제품 요구규격 CSV</button></div><div class="tableWrap"><table id="bom-live-table"><thead><tr><th>구간 / 품목</th><th>설치 / 구매 범위</th><th>요구 규격</th><th>등록 제품 후보 · 출처 · 검증 상태</th></tr></thead><tbody>'+shown.map(x=>'<tr><td>'+esc(x.segment||'설계 참고')+'<br>'+esc(x.item)+(x.referenceOnly?'<small style="display:block">참고 · 중복 구매 제외</small>':'')+'</td><td>'+fmt(x.installed)+' / '+(x.low!=null?fmt(x.low)+'–'+fmt(x.high):fmt(x.purchase))+' '+esc(x.unit||'')+(x.phases.length?'<small style="display:block">Phase '+x.phases.join(', ')+' 합계</small>':'')+'</td><td>'+esc(x.spec||[x.profile?.media,x.profile?.connector].filter(Boolean).join(' · '))+(x.lengthM!=null?'<br>'+fmt(x.lengthM)+'m':'')+'</td><td>'+(x.candidates.length?x.candidates.slice(0,3).map(link).join('<br>')+(x.candidates.length>3?'<details><summary>추가 후보 '+(x.candidates.length-3)+'개</summary>'+x.candidates.slice(3,20).map(link).join('<br>')+(x.candidates.length>20?'<p>전체 후보는 CSV에서 확인하세요.</p>':'')+'</details>':''):'규격을 확인할 등록 후보 없음 · RFQ')+'</td></tr>').join('')+'</tbody></table></div>'+(data.errors.length?'<details><summary>조회 실패 파일</summary>'+data.errors.map(x=>'<p>'+esc(x.path+': '+x.error)+'</p>').join('')+'</details>':'');
    $('bom-catalog-vendor').onchange=e=>{vendor=e.target.value;paint();};
    $('bom-catalog-refresh').onclick=()=>refresh(true);
    $('exportProductCsv').onclick=()=>downloadCsv();
    r.catalogRows=rows;r.catalogSnapshot={revision:data.revision,catalogs:data.catalogs.length,products:data.products.length,errors:data.errors};
  }
  async function refresh(force=false) {
    if(!current)return;
    const token=++revision,h=host();if(h)h.innerHTML='<h2>설계 제품·케이블 품목 리스트</h2><p id="bom-catalog-status" role="status">업로드된 전체 카탈로그를 동기화하는 중…</p>';
    ready=(async()=>{try{const loaded=await M.load(force);if(token!==revision)return;data=loaded;const raw=rowsFor(current.r,current.a);for(const x of raw)x.candidates=M.match(x,data.products);current.rows=aggregate(raw);if(current.a){current.r.catalogCandidates=current.rows.map(x=>({segment:x.segment||x.item,spec:x.spec||'',candidates:x.candidates}));}else if(current.r.productRequirements){for(let i=0;i<current.r.productRequirements.length;i++)current.r.productRequirements[i].candidates=raw[i].candidates;}paint();}catch(e){if(token!==revision)return;current.r.catalogRows=[];if(h)h.innerHTML='<h2>설계 제품·케이블 품목 리스트</h2><p role="alert">'+esc(e.message)+' · 제품 연결 미완료. 이전 후보를 재사용하지 않습니다.</p><button id="bom-catalog-retry">다시 동기화</button>';if($('bom-catalog-retry'))$('bom-catalog-retry').onclick=()=>refresh(true);}})();
    return ready;
  }
  const headers=['구간','품목','설치','구매 기준','하한','상한','단위','요구 규격','벤더','제품','검증 상태','매칭 근거','공식 URL','카탈로그 경로','카탈로그 Git SHA'];
  function exportRows() {return (current?.rows||[]).flatMap(x=>{const cs=x.candidates.length?x.candidates:[{}];return cs.map(p=>[x.segment||'',x.item,x.installed??'',x.purchase??'',x.low??'',x.high??'',x.unit,x.spec||x.profile?.media||'',p.vendor||'',p.name||'',p.status||(x.referenceOnly?'설계 참고 · 미산정':'RFQ'),p.matchBasis||'',p.source||'',p.path||'',data?.revision||'']);});}
  async function downloadCsv() {await ready;if(!current?.rows)return;const csv=[headers,...exportRows()].map(row=>row.map(v=>'"'+String(v??'').replace(/^([=+@-])/,'\t$1').replace(/"/g,'""')+'"').join(',')).join('\r\n');const url=URL.createObjectURL(new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8'})),a=document.createElement('a');a.href=url;a.download='datacenter-catalog-product-requirements.csv';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);}
  document.addEventListener('dc:capacity-bom',e=>{current={r:e.detail.r,a:e.detail.a,rows:[]};refresh();});
  document.addEventListener('dc:design',e=>{if(window.DCCapacityDesign)return;current=e.detail?.usable?{r:e.detail,a:null,rows:[]}:null;requestAnimationFrame(()=>{if(current)refresh();else{const h=$('bom-live-products');if(h)h.remove();}});});
  document.addEventListener('dc:catalog-refresh',()=>refresh(true));
  document.addEventListener('dc:capacity-invalid',()=>{revision++;current=null;const h=$('bom-live-products');if(h)h.remove();});
  document.addEventListener('dc:stale',()=>{revision++;current=null;const h=$('bom-live-products');if(h)h.innerHTML='<h2>설계 제품·케이블 품목 리스트</h2><p>요구조건 변경 · 설계안을 다시 계산하세요.</p>';});
  window.DCBomCatalogUI={get ready(){return ready;},refresh,exportRows,headers};
})();
