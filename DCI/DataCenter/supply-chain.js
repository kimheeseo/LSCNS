(() => {
'use strict';
const VERSION='7.4.18';
let selectedCategory='all',returnFocus;
const textOf=e=>(e&&(e.innerText||e.textContent)||'').replace(/\s+/g,' ').trim();

const I18N={
  ko:{button:'Supply Chain',title:'BOM Supply Chain · 업체 / 제품 맵',sub:'현재 설계에 실제 사용된 업체·제품·수량을 우선 표시하고, 각 영역별 관련 기업은 별도 참고 목록으로 함께 보여줍니다.',need:'먼저 설계 계산을 실행해 주세요.',center:'현재 BOM',used:'BOM 반영 항목',sources:'제품 / 데이터시트',close:'닫기',refresh:'새로고침',
      cats:{compute:'GPU / Accelerator',cpu:'Server CPU',network:'Network Fabric',optical:'Optical Connectivity',power:'Power / UPS / PDU',cooling:'Cooling / HVAC',rack:'Rack / Physical',storage:'Storage',facility:'Facility / Security'}},
  en:{button:'Supply Chain',title:'BOM Supply Chain · Vendor / Product Map',sub:'Shows vendors/products actually used by the current BOM first, plus a separate related-companies reference list for each category.',need:'Run the design calculation first.',center:'Current BOM',used:'BOM items',sources:'Product / datasheet',close:'Close',refresh:'Refresh',
      cats:{compute:'GPU / Accelerator',cpu:'Server CPU',network:'Network Fabric',optical:'Optical Connectivity',power:'Power / UPS / PDU',cooling:'Cooling / HVAC',rack:'Rack / Physical',storage:'Storage',facility:'Facility / Security'}},
  ja:{button:'サプライチェーン',title:'BOM サプライチェーン',sub:'現在のGeneric BOM / 製品マッチング結果に実際に含まれる製品をサプライチェーン視点で再構成します。',need:'先に設計計算を実行してください。',center:'現在のBOM',used:'BOM項目',sources:'製品 / データシート',close:'閉じる',refresh:'更新',
      cats:{compute:'GPU / Accelerator',cpu:'Server CPU',network:'Network Fabric',optical:'Optical Connectivity',power:'Power / UPS / PDU',cooling:'Cooling / HVAC',rack:'Rack / Physical',storage:'Storage',facility:'Facility / Security'}},
  zh:{button:'供应链',title:'BOM 供应链',sub:'将当前 Generic BOM / 产品匹配结果中实际包含的产品按供应链视角重新整理。',need:'请先执行设计计算。',center:'当前 BOM',used:'BOM 项目',sources:'产品 / 数据表',close:'关闭',refresh:'刷新',
      cats:{compute:'GPU / Accelerator',cpu:'Server CPU',network:'Network Fabric',optical:'Optical Connectivity',power:'Power / UPS / PDU',cooling:'Cooling / HVAC',rack:'Rack / Physical',storage:'Storage',facility:'Facility / Security'}},
  de:{button:'Lieferkette',title:'BOM-Lieferkette',sub:'Ordnet die tatsächlich im aktuellen Generic BOM / Produkt-Matching enthaltenen Produkte als Lieferkettenansicht neu.',need:'Bitte zuerst die Designberechnung ausführen.',center:'Aktuelles BOM',used:'BOM-Positionen',sources:'Produkt / Datenblatt',close:'Schließen',refresh:'Aktualisieren',
      cats:{compute:'GPU / Accelerator',cpu:'Server CPU',network:'Network Fabric',optical:'Optical Connectivity',power:'Power / UPS / PDU',cooling:'Cooling / HVAC',rack:'Rack / Physical',storage:'Storage',facility:'Facility / Security'}}
};

const VENDORS=[
  'NVIDIA','AMD','Intel','Biren Technology','Ampere Computing','Juniper','Cisco','Arista','Broadcom','Credo','Lenovo','Supermicro','Dell','HPE','Hewlett Packard Enterprise',
  'Corning','SENKO','US Conec','ZTT','Sumitomo Electric','Sumitomo','LS Cable & System','LS Cable','YOFC','Hengtong','Lightera','Fujikura','CommScope','Molex','Amphenol','Panduit','Belden',
  'Schneider Electric','Schneider','APC','Vertiv','Eaton','Legrand','ABB','Rittal','Delta','Huawei','Siemens','Generac','Caterpillar','CAT',
  'Pure Storage','NetApp','IBM','Micron','Samsung','Solidigm','Kioxia','Western Digital',
  'Mitsubishi Electric','MPS','Hitachi','Rolls-Royce','Atlas Copco','EnerSys','Cummins','Munters','STULZ','Carrier','Trane','Modine',
  'Fortinet','Palo Alto Networks','Palo Alto','Bosch','Securitas','Oracle','Fujitsu','Emerson','Asetek'
];
const CATEGORY_RULES=[
  ['cpu',/server\s*CPU|\bCPU\b|Xeon|EPYC|Grace CPU|AmpereOne|processor/i],
  ['compute',/\bGPU\b|DGX|H100|H200|B200|B300|GB200|GB300|NVL72|compute|server|accelerator|supermicro|lenovo|dell|hpe|hewlett/i],
  ['network',/switch|leaf|spine|core|fabric|NIC|DPU|ConnectX|BlueField|Spectrum|QFX|Nexus|Arista|Juniper|Cisco|Broadcom|Tomahawk|InfiniBand|Ethernet/i],
  ['optical',/optic|transceiver|fiber|fibre|cable|trunk|patch|MPO|MTP|LC\b|OSFP|QSFP|AOC|DAC|AEC|DR4|FR4|SR8|Corning|Sumitomo|LS Cable|YOFC|Hengtong|Lightera|Fujikura|Credo|CommScope|Molex|Amphenol|Panduit|Belden/i],
  ['power',/UPS|PDU|power|generator|transformer|busway|breaker|switchgear|Schneider|APC|Vertiv|Eaton|Legrand|ABB|Delta|Generac|Caterpillar|\bCAT\b/i],
  ['cooling',/cooling|CDU|RDHx|chiller|HVAC|liquid|rear.?door|coolant|CRAC|CRAH|Vertiv|Schneider|Rittal|Munters|STULZ|Carrier|Trane|Modine|Emerson|Asetek/i],
  ['rack',/\brack\b|cabinet|enclosure|rail|cable manager|rack PDU|Rittal|Legrand/i],
  ['storage',/storage|NVMe|SSD|HDD|RAID|Pure Storage|NetApp|Solidigm|Kioxia|Micron|Western Digital/i],
  ['facility',/security|fire|camera|access control|BMS|building automation|monitoring|sensor|Siemens|Fortinet|Palo Alto|Palo Alto Networks|Bosch|Securitas/i]
];
const RELATED_VENDORS={
  compute:['NVIDIA','AMD','Intel','Biren Technology','Supermicro','Dell Technologies','HPE','Lenovo','Fujitsu','Oracle','IBM'],
  cpu:['Intel','AMD','NVIDIA','Ampere Computing'],
  network:['NVIDIA Networking','Cisco','Arista Networks','Juniper Networks','Broadcom','Marvell','HPE Aruba Networking'],
  optical:['Corning','SENKO','US Conec','ZTT','Sumitomo Electric','Fujikura','Furukawa Electric','Lightera','LS Cable & System','Hengtong','YOFC','CommScope','Molex','Amphenol','Panduit','Belden'],
  power:['Schneider Electric','Vertiv','Eaton','ABB','Siemens','Legrand','Mitsubishi Electric','Cummins','Caterpillar'],
  cooling:['Vertiv','Schneider Electric','Carrier','Trane','Munters','STULZ','Modine','Rittal','Asetek'],
  rack:['Rittal','Legrand','Vertiv','Eaton','Schneider Electric','HPE','Dell Technologies','Supermicro','Fujitsu'],
  storage:['Pure Storage','NetApp','Dell Technologies','HPE','IBM','Micron','Samsung','Kioxia','Solidigm','Western Digital'],
  facility:['Siemens','Schneider Electric','Fortinet','Palo Alto Networks','Bosch','Cisco','Securitas']
};

const VERIFIED_DC_PRODUCTS={
  power:[
    {vendor:'Schneider Electric',product:'Galaxy VXL',role:'3-phase UPS · AI / large data center',spec:'500–1250 kW (400 V)',url:'https://www.se.com/kr/ko/product-range/209756733-galaxy-vxl/'},
    {vendor:'Eaton',product:'9395X UPS',role:'Hyperscale / colocation UPS',spec:'1.0–1.7 MVA · 97.5% online efficiency',url:'https://www.eaton.com/gb/en-gb/catalog/backup-power-ups-surge-it-power-distribution/eaton-9395x-ups.html'},
    {vendor:'ABB',product:'MegaFlex DPA',role:'High-density data-center UPS',spec:'250–1500 kW',url:'https://new.abb.com/ups/ups-and-power-conditioning/megaflex'},
    {vendor:'Vertiv',product:'Liebert APM2',role:'Modular mission-critical UPS',spec:'30–600 kVA · 400 V',url:'https://go.vertiv.com/LiebertAPM2'},
    {vendor:'Mitsubishi Electric',product:'9900D',role:'Hyperscale / colocation UPS',spec:'1200–2000 kVA · 480 V',url:'https://mitsubishicritical.com/uninterruptible-power-supplies/9900d/'},
    {vendor:'Siemens',product:'SIVACON S8',role:'LV power-distribution switchboard',spec:'IEC 61439-2 · data center / critical infrastructure',url:'https://www.siemens.com/en-us/products/sivacon/s8/'},
    {vendor:'Caterpillar',product:'C175-20',role:'Mission-critical / data-center generator',spec:'3150–4000 ekW · 60 Hz',url:'https://www.cat.com/en_US/products/new/power-systems/electric-power/diesel-generator-sets/1000028913.html'},
    {vendor:'Cummins',product:'QSK95 generator platform',role:'Data Center Continuous generator platform',spec:'C3500D5 example: 2500 kW DCC',url:'https://www.cummins.com/en-na/generators/products/qsk95'}
  ],
  cooling:[
    {vendor:'Vertiv',product:'Liebert XDU450',role:'Coolant Distribution Unit',spec:'453 kW nominal · up to 975 kW max',url:'https://www.vertiv.com/en-us/products-catalog/thermal-management/high-density-solutions/liebert-xdu450-coolant-distribution-unit/'},
    {vendor:'Carrier',product:'AquaForce 30XF',role:'Mission-critical data-center air-cooled screw chiller',spec:'Integrated free-cooling platform',url:'https://www.carrier.com/us/en/commercial/chillers/'},
    {vendor:'Trane',product:'CenTraVac CDHH / CVHH',role:'Water-cooled data-center chiller',spec:'CDHH up to 21 MW · CVHH up to 9 MW',url:'https://www.trane.com/commercial/north-america/us/en/products-systems/chillers/data-center-chillers/centravac-data-center-chiller.html'}
  ],
  rack:[
    {vendor:'Vertiv',product:'VR Rack family',role:'Data-center rack / enclosure',spec:'VR3100 / VR3300 / VR3150 / VR3350 families',url:'https://www.vertiv.com/en-us/products-catalog/facilities-enclosures-and-racks/racks-and-containment/vertiv-rack/'}
  ],
  facility:[
    {vendor:'Fortinet',product:'FortiGate 3800G',role:'AI data-center firewall',spec:'200 Gbps threat protection · 400 GbE connectivity',url:'https://www.fortinet.com/solutions/data-center-firewall'}
  ]
};
window.__dcBomVerifiedSupplyCatalog=VERIFIED_DC_PRODUCTS;

function verifiedProductHtml(k){
  const arr=VERIFIED_DC_PRODUCTS[k]||[];
  if(!arr.length)return'';
  return '<div class="sc-verified"><div class="sc-related-title">Verified DC products <span>공식 데이터센터 용도/사양 확인</span></div>'+
    arr.map(x=>'<div class="sc-vproduct"><div><b>'+esc(x.vendor)+'</b> · '+esc(x.product)+'</div><div class="sc-vrole">'+esc(x.role)+'</div><div class="sc-vspec">'+esc(x.spec)+'</div><a href="'+esc(x.url)+'" target="_blank" rel="noopener">Official product / datasheet</a></div>').join('')+
    '</div>';
}
function relatedHtml(k,used){
  const usedSet=new Set((used||[]).map(x=>(x.vendor||'').toLowerCase()).filter(Boolean));
  const arr=(RELATED_VENDORS[k]||[]).filter(v=>!usedSet.has(v.toLowerCase()));
  const products=verifiedProductHtml(k);
  const companies=arr.length?'<div class="sc-related"><div class="sc-related-title">Related companies <span>참고 업체 · 현재 BOM 미선택</span></div><div class="sc-related-chips">'+arr.map(v=>'<span>'+esc(v)+'</span>').join('')+'</div></div>':'';
  return products+companies;
}

function token(s){s=String(s||'').trim().toLowerCase();if(s.includes('한국')||s.includes('korean')||s==='ko'||s==='kr')return'ko';if(s.includes('日本')||s.includes('japanese')||s==='ja'||s==='jp')return'ja';if(s.includes('中文')||s.includes('chinese')||s==='zh'||s==='cn')return'zh';if(s.includes('deutsch')||s.includes('german')||s==='de')return'de';if(s.includes('english')||s==='en')return'en';return null}
function lang(){if(window.__dcBomUiLang&&I18N[window.__dcBomUiLang])return window.__dcBomUiLang;const s=document.documentElement.getAttribute('data-dc-bom-ui-lang');if(s&&I18N[s])return s;return token(document.documentElement.lang)||'ko'}
const tr=()=>I18N[lang()]||I18N.ko;

function findCalc(){return Array.from(document.querySelectorAll('button,input[type="button"],input[type="submit"]')).find(e=>{const s=(e.value||textOf(e)).toLowerCase();return(s.includes('설계')&&s.includes('계산'))||s.includes('calculate')||s.includes('design calculation')||s.includes('計算')||s.includes('设计计算')||s.includes('berechnen')})||null}
function findTable(rx){
  for(const h of document.querySelectorAll('h1,h2,h3,h4,h5,summary,.phaseBand,.section-title,.card-title')){
    if(!rx.test(textOf(h)))continue;
    const scope=h.closest('.card,.panel,section,article,details')||h.parentElement;
    if(scope&&scope.querySelector('table'))return scope.querySelector('table');
    let n=h.nextElementSibling;
    for(let i=0;n&&i<8;i++,n=n.nextElementSibling){if(n.matches&&n.matches('table'))return n;if(n.querySelector&&n.querySelector('table'))return n.querySelector('table')}
  }
  return null;
}
function vendorFrom(text){const low=text.toLowerCase();return VENDORS.find(v=>low.includes(v.toLowerCase()))||''}
function categorize(text){for(const [key,rx] of CATEGORY_RULES)if(rx.test(text))return key;return'facility'}
function rowItems(table,source){
  if(!table)return[];
  const rows=Array.from(table.querySelectorAll('tr'));if(rows.length<2)return[];
  const headers=Array.from(rows[0].children).map(x=>textOf(x).toLowerCase());
  const idx=(rx)=>headers.findIndex(x=>rx.test(x));
  const vendorIdx=idx(/vendor|manufacturer|제조사|기업|maker|hersteller/);
  const productIdx=idx(/product|model|제품|모델|recommended|recommendation|추천/);
  const qtyIdx=idx(/qty|quantity|수량|menge/);
  return rows.slice(1).map((tr,ri)=>{
    const cells=Array.from(tr.children).filter(x=>x.tagName==='TD'||x.tagName==='TH');
    if(!cells.length)return null;
    const texts=cells.map(textOf),row=texts.join(' · ');
    if(!row||row==='-')return null;
    const links=Array.from(tr.querySelectorAll('a[href]')).map(a=>({label:textOf(a)||'Link',url:a.href})).filter(x=>/^https?:/i.test(x.url));
    const vendor=(vendorIdx>=0&&texts[vendorIdx])||vendorFrom(row);
    let product=(productIdx>=0&&texts[productIdx])||'';
    if(!product){product=texts.find((x,i)=>i!==vendorIdx&&i!==qtyIdx&&x&&x!=='-'&&!/^\d+(\.\d+)?$/.test(x))||row}
    const qty=qtyIdx>=0?texts[qtyIdx]:'';
    return{id:source+'-'+ri,source,vendor,product,qty,row,links,category:categorize(row)};
  }).filter(Boolean);
}
function collect(){
  const r=window.DCDesign;
  if(r){const primary=(r.products||[]).map((x,i)=>({id:'product-'+i,source:'receipt',vendor:x.vendor,product:x.product,qty:String(x.qty??''),row:[x.category,x.product,x.model].join(' · '),links:x.source&&/^https?:/i.test(x.source)?[{label:'Product / Datasheet',url:x.source}]:[],category:categorize([x.category,x.product,x.model].join(' · '))}));
    const secondary=(r.bom||[]).map((x,i)=>({id:'bom-'+i,source:'generic',vendor:'',product:x.item,qty:String(x.qty??''),row:[x.category,x.item].join(' · '),links:[],category:categorize([x.category,x.item].join(' · '))}));
    const seen=new Set(primary.map(x=>x.product.toLowerCase()));
    return [...primary,...secondary.filter(x=>!seen.has(x.product.toLowerCase()))];
  }
  const receipt=findTable(/제품.*매칭.*영수증|product.*match.*receipt|product.*receipt|製品.*マッチ|产品.*匹配|produkt.*matching/i);
  const generic=findTable(/generic\s*bom|일반\s*bom|범용\s*bom|汎用.*bom|通用.*bom/i);
  const primary=rowItems(receipt,'receipt'),secondary=rowItems(generic,'generic');
  const seen=new Set(),items=[];
  for(const x of [...primary,...secondary]){
    const k=(x.vendor+'|'+x.product+'|'+x.category).toLowerCase();
    if(seen.has(k))continue;seen.add(k);items.push(x);
  }
  return items;
}
function candidateProductsForRow(row){
  const r=String(row||'');
  if(/generator|genset|backup power|diesel|발전기/i.test(r))return VERIFIED_DC_PRODUCTS.power.filter(x=>/generator/i.test(x.role));
  if(/switchgear|switchboard|LV power|MV power|배전|MCC/i.test(r))return VERIFIED_DC_PRODUCTS.power.filter(x=>/switchboard/i.test(x.role));
  if(/UPS|uninterruptible|무정전/i.test(r))return VERIFIED_DC_PRODUCTS.power.filter(x=>/UPS/i.test(x.role)).slice(0,5);
  if(/CDU|coolant distribution|liquid cooling|cooling|HVAC|chiller|냉각/i.test(r))return VERIFIED_DC_PRODUCTS.cooling;
  if(/rack|cabinet|enclosure|랙|캐비닛/i.test(r))return VERIFIED_DC_PRODUCTS.rack;
  if(/firewall|security|보안/i.test(r))return VERIFIED_DC_PRODUCTS.facility;
  return[];
}
function enrichProductReceipt(){
  const table=findTable(/제품.*매칭.*영수증|product.*match.*receipt|product.*receipt|製品.*マッチ|产品.*匹配|produkt.*matching/i);
  if(!table)return false;
  const rows=Array.from(table.querySelectorAll('tr'));if(rows.length<2)return false;
  const heads=Array.from(rows[0].children).map(x=>textOf(x).toLowerCase());
  let ai=heads.findIndex(x=>/alternative|대안|대체|候補|替代|alternate/.test(x));
  if(ai<0)ai=heads.length-1;
  for(const tr of rows.slice(1)){
    if(tr.getAttribute('data-dc-supply-enriched')==='1')continue;
    const cells=Array.from(tr.children).filter(x=>x.tagName==='TD'||x.tagName==='TH');if(!cells[ai])continue;
    const cand=candidateProductsForRow(textOf(tr));if(!cand.length){tr.setAttribute('data-dc-supply-enriched','1');continue}
    const d=document.createElement('div');d.className='sc-receipt-candidates';
    d.innerHTML='<b>Verified DC candidates</b><br>'+cand.slice(0,5).map(x=>'<a href="'+esc(x.url)+'" target="_blank" rel="noopener">'+esc(x.vendor)+' · '+esc(x.product)+'</a>').join(' · ');
    cells[ai].appendChild(d);tr.setAttribute('data-dc-supply-enriched','1');
  }
  return true;
}

function esc(s){return String(s||'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]))}
function itemHtml(x){
  const vendor=x.vendor||'Vendor TBD';
  const qty=x.qty&&x.qty!=='-'?'<span class="sc-qty">× '+esc(x.qty)+'</span>':'';
  const src=x.source==='receipt'?'Product Match':'Generic BOM';
  const links=x.links.slice(0,2).map(a=>'<a href="'+esc(a.url)+'" target="_blank" rel="noopener">'+esc(a.label||'Product / Datasheet')+'</a>').join(' ');
  return'<div class="sc-item"><div class="sc-vendor">'+esc(vendor)+'<span class="sc-source">'+src+'</span></div><div class="sc-product">'+esc(x.product)+qty+'</div>'+(links?'<div class="sc-links">'+links+'</div>':'')+'</div>';
}
function modal(){
  let m=document.getElementById('supply-chain-modal');if(m)return m;
  m=document.createElement('div');m.id='supply-chain-modal';m.hidden=true;m.setAttribute('role','dialog');m.setAttribute('aria-modal','true');m.setAttribute('aria-label','Supply Chain');
  m.innerHTML='<div class="sc-shell"><div class="sc-top"><div><h2 class="sc-title"></h2><p class="sc-sub"></p></div><div class="sc-top-actions"><button type="button" class="sc-refresh"></button><button type="button" class="sc-close">×</button></div></div><div class="sc-body"></div><div class="sc-foot">v'+VERSION+' · Generated from the current on-screen BOM / product-match result</div></div>';
  document.body.appendChild(m);
  m.querySelector('.sc-close').onclick=close;
  m.querySelector('.sc-close').setAttribute('aria-label',tr().close);
  m.addEventListener('click',e=>{if(e.target===m)close()});
  m.querySelector('.sc-refresh').onclick=render;
  return m;
}
function render(){
  const m=modal(),t=tr(),items=collect();
  m.querySelector('.sc-title').textContent=t.title;
  m.querySelector('.sc-sub').textContent=t.sub;
  m.querySelector('.sc-refresh').textContent=t.refresh;
  
  const groups={};for(const k of Object.keys(t.cats))groups[k]=[];
  items.forEach(x=>(groups[x.category]||(groups[x.category]=[])).push(x));
  const vendorCount=new Set(items.map(x=>x.vendor).filter(Boolean)).size;
  const center='<section class="sc-center"><div class="sc-center-icon">DC</div><h3>'+t.center+'</h3><div class="sc-center-stat"><b>'+items.length+'</b> '+t.used+'</div><div class="sc-center-stat"><b>'+vendorCount+'</b> Vendors</div></section>';
  const order=['compute','cpu','network','optical','power','cooling','rack','storage','facility'];
  const cards=order.filter(k=>selectedCategory==='all'||k===selectedCategory).map(k=>{
    const arr=groups[k]||[];
    const body=(arr.length?arr.map(itemHtml).join(''):'<div class="sc-none">현재 BOM 선택 업체 없음</div>')+relatedHtml(k,arr);
    return'<section class="sc-card sc-'+k+'"><div class="sc-cat-head"><span class="sc-dot"></span><h3>'+esc(t.cats[k])+'</h3><span class="sc-count">'+arr.length+' used</span></div>'+body+'</section>';
  }).join('');
  m.querySelector('.sc-body').innerHTML='<nav class="sc-filters" aria-label="부품별 업체"><button type="button" data-category="all" aria-pressed="'+(selectedCategory==='all')+'">전체</button>'+order.map(k=>'<button type="button" data-category="'+k+'" aria-pressed="'+(selectedCategory===k)+'">'+esc(t.cats[k])+'</button>').join('')+'</nav>'+(window.DCDesignStale?'<p class="sc-stale">입력값 변경됨 · BOM 수량은 이전 계산 기준입니다.</p>':'')+'<div class="sc-map'+(selectedCategory!=='all'?' sc-single':'')+'">'+(selectedCategory==='all'?center:'')+cards+'</div>';
  m.querySelectorAll('[data-category]').forEach(b=>b.onclick=()=>{selectedCategory=b.dataset.category;render();m.querySelector('[data-category="'+selectedCategory+'"]')?.focus();});
}
function close(){const m=document.getElementById('supply-chain-modal');if(m)m.hidden=true;returnFocus?.focus();}
function open(){returnFocus=document.activeElement;render();modal().hidden=false;modal().querySelector('.sc-close').focus();}
function mount(){
  if(document.getElementById('supply-chain-btn'))return true;
  const calc=document.getElementById('run')||findCalc();if(!calc||!calc.parentNode)return false;
  const host=document.createElement('span');host.id='supply-chain-inline-host';
  const b=document.createElement('button');b.type='button';b.id='supply-chain-btn';b.textContent=tr().button;b.title='BOM 설계에 실제 사용된 업체/제품 보기';b.onclick=open;host.appendChild(b);
  calc.insertAdjacentElement('afterend',host);
  return true;
}
function localize(){const b=document.getElementById('supply-chain-btn');if(b)b.textContent=tr().button;const m=document.getElementById('supply-chain-modal');if(m&&!m.hidden)render()}
function start(){
  mount();
  document.addEventListener('keydown',e=>{const m=document.getElementById('supply-chain-modal');if(!m||m.hidden)return;if(e.key==='Escape')close();if(e.key==='Tab'){const a=[...m.querySelectorAll('button,a[href]')],first=a[0],last=a[a.length-1];if(e.shiftKey&&document.activeElement===first){e.preventDefault();last.focus();}else if(!e.shiftKey&&document.activeElement===last){e.preventDefault();first.focus();}}});
  document.addEventListener('dc:design',()=>{localize();enrichProductReceipt();});
  let attempts=0;const boot=setInterval(()=>{attempts++;const mounted=mount();const enriched=enrichProductReceipt();if((mounted&&enriched)||attempts>24)clearInterval(boot)},500);
  document.addEventListener('click',e=>{const q=e.target&&e.target.closest&&e.target.closest('[data-lang],[data-language],button,a,[role="button"]');if(q)setTimeout(()=>{localize();enrichProductReceipt()},50)},true);
  document.addEventListener('change',()=>setTimeout(()=>{localize();enrichProductReceipt()},30),true);
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start,{once:true});else start();
})();
