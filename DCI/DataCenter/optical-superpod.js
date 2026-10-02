
(() => {
  'use strict';
  const TOOL_VERSION='7.3.2';
  const CATALOG={
    ls:{name:'LS Cable & System',home:'https://www.lscns.com/',trunkSizes:[48,96,144,288,432,864,3456],pretermSizes:[48,96,144,288],highCountSizes:[144,288,432,864,3456],family:'Micro Array 48–288F / Lock\'n Roll™ Ribbon up to 3,456F',connectors:'LC / SC / SN / CS / MPO',note:'Public product materials show Micro Array and Lock\'n Roll high-density ribbon families for data-center and high-count applications.',url:'https://www.lssimple.com/en/product/category3.asp?cate1=204&cate2=567&cate3=570'},
    hengtong:{name:'Hengtong',home:'https://www.hengtongglobal.com/',trunkSizes:[],pretermSizes:[],highCountSizes:[],family:'High-density MPO trunk / pre-terminated data-center cable',connectors:'LC / SC / MPO/MTP',note:'Public pages confirm high-density MPO/trunk families for 100G/400G/800G+, while exact current high-count SKU/fiber-count tables require RFQ.',url:'https://www.hengtongglobal.com/data-center-connectivity-solutions'},
    yofc:{name:'YOFC',home:'https://en.yofc.com/',trunkSizes:[12,24,48,72,96,144],pretermSizes:[12,24,48,72,96,144],highCountSizes:[],family:'12–144F MPO/MTP pre-terminated trunk',connectors:'MTP/MPO / LC',note:'YOFC public trunk specifications list 12/24/48/72/96/144F with G.657.A2 and BI OM3/OM4/OM5 options.',url:'https://en.yofc.com/view/3030.html'},
    lightera:{name:'Lightera',home:'https://lightera.com/',trunkSizes:[12,16,24,48,72,96,144,288,432,1728,3456],pretermSizes:[12,16,24,48,72,96,144,288,432],highCountSizes:[144,288,432,1728,3456],family:'DuctSaver® Rollable Ribbon 24–432F / AccuTube®+ RR 1,728–3,456F / R-Pack RR',connectors:'Multifiber connector / MPO-ready rollable ribbon',note:'Lightera positions rollable-ribbon cable for AI/HPC and data centers; DuctSaver is published in 24–432F and AccuTube+ RR in 1,728/3,456F.',url:'https://lightera.com/rollable-ribbon/'},
    sumitomo:{name:'Sumitomo Electric',home:'https://sumitomoelectric.com/',trunkSizes:[288,1728,3456,6912],pretermSizes:[],highCountSizes:[288,1728,3456,6912],family:'FREEFORM RIBBON™ UHFC / 3,456F pre-connectorized MPO / 6,912F high-density cable',connectors:'24F MPO (3,456F pre-connectorized) / mass-fusion ribbon',note:'Official data-center publications describe 1,728F/3,456F UHFC, a 3,456F pre-connectorized cable with 24F MPO, and 6,912F high-density Freeform Ribbon cable.',url:'https://sumitomoelectric.com/products/optical-cables'},
    corning:{name:'Corning',home:'https://www.corning.com/data-center/',trunkSizes:[8,16,24,32,48,72,96,144,192,288,432,576,864,1728,3456],pretermSizes:[8,16,24,32,48,72,96,144,192,288],highCountSizes:[288,432,576,864,1728,3456],family:'EDGE8® MTP® trunks 8–288F / RocketRibbon® up to 3,456F',connectors:'MTP/MPO / LC duplex / VSFF MDC / SN / CS',note:'EDGE8 public selection guides list OS2 preterminated trunks through 288F; RocketRibbon data-center interconnect cable is published up to 3,456 fibers.',url:'https://www.corning.com/data-center/worldwide/en/home/solutions/edge8.html'},
    ztt:{name:'ZTT',home:'https://www.zttgroup.com/',trunkSizes:[17280],pretermSizes:[],highCountSizes:[17280],family:'17,280F ultra-high-density flexible-ribbon optical cable',connectors:'Multi-level branching / multi-mode termination; exact connectorization by project',note:'ZTT announced a 17,280-fiber, 45 mm ultra-high-density cable in 2026 for ultra-large data centers, AI computing clusters and backbone networks.',url:'https://www.zttgroup.com/news/show-733.html'}
  };
  const SWITCH_CATALOG=[
    {vendor:'Cisco',model:'N9364E-SG2-Q',ports:'64 × 800G QSFP-DD',breakout:'2×400G / 8×100G',ru:'2RU',typical:'995 W',max:'2,270 W',role:'AI leaf / spine · high-density 800G',url:'https://www.cisco.com/c/en/us/products/collateral/switches/nexus-9000-series-switches/nexus-9364e-sg2-switch-ds.html'},
    {vendor:'Cisco',model:'N9364E-SG2-O',ports:'64 × 800G OSFP',breakout:'2×400G / 8×100G',ru:'2RU',typical:'995 W',max:'2,270 W',role:'AI leaf / spine · high-density 800G',url:'https://www.cisco.com/c/en/us/products/collateral/switches/nexus-9000-series-switches/nexus-9364e-sg2-switch-ds.html'},
    {vendor:'Cisco',model:'N9364E-SP2R-Q',ports:'64 × 800G QSFP-DD',breakout:'2×400G / 4×200G / 8×100G + lower-speed modes',ru:'3RU',typical:'2,500 W',max:'5,100 W',role:'Deep-buffer spine / DCI · 16GB HBM',url:'https://www.cisco.com/c/en/us/products/collateral/switches/nexus-9000-series-switches/nexus-9364e-sp2r-switches-ds.html'},
    {vendor:'Cisco',model:'N9364E-SP2R-O',ports:'64 × 800G OSFP',breakout:'2×400G / 4×200G / 8×100G + lower-speed modes',ru:'3RU',typical:'2,500 W',max:'5,100 W',role:'Deep-buffer spine / DCI · 16GB HBM',url:'https://www.cisco.com/c/en/us/products/collateral/switches/nexus-9000-series-switches/nexus-9364e-sp2r-switches-ds.html'}
  ];
  const MPO_SOURCES={
    corning8:'https://ecatalog.corning.com/optical-communications/US/en/Fiber-Optic-Cable-Assemblies/Indoor-Cable-Assemblies/Multifiber-Indoor-Cable-Assemblies/EDGE8%C2%AE-MTP%C2%AE-Trunk/p/edge8-mtp-trunk-cable?variant=fiber-count-8',
    corning12:'https://ecatalog.corning.com/optical-communications/US/en/Fiber-Optic-Cable-Assemblies/Indoor-Cable-Assemblies/Multifiber-Indoor-Cable-Assemblies/EDGE%E2%84%A2-MTP%C2%AE-Trunk/p/G757512TPNDDU120F',
    senko16:'https://www.senko.com/wp-content/uploads/2021/09/MPO-16-Connector.pdf',
    usconec16:'https://www.usconec.com/media/1i4pg2b5/mtp-16_connector_handout.pdf',
    corningMdc:'https://ecatalog.corning.com/optical-communications/US/en/Fiber-Optic-Cable-Assemblies/Indoor-Cable-Assemblies/Two-Fiber-Indoor-Cable-Assemblies/Fiber-Optic-Jumper%2C-2F%2C-MDC-to-MDC/p/MUMU02QD120035M',
    senkoSn:'https://www.senko.com/sn-series/',
    senkoCs:'https://www.senko.com/product/cs-standard-connector/',
    corningVsff:'https://www.corning.com/data-center/worldwide/en/home/applications/multi-tenant-data-center/solve-network-challenges-with-very-small-form-factor-connectors.html'
  };
  const NVIDIA_SOURCE='https://docs.nvidia.com/dgx-superpod/reference-architecture-scalable-infrastructure-b200/latest/dgx-superpod-architecture.html';
  const NVIDIA_COMPONENTS='https://docs.nvidia.com/dgx-superpod/reference-architecture-scalable-infrastructure-b200/latest/components.html';
  const state={lastDesign:null,lastGpu:null,userLinks:false};
  const num=v=>{const x=Number(String(v==null?'':v).replace(/[^0-9.+-]/g,''));return Number.isFinite(x)?x:null};
  const txt=el=>(el&&(el.innerText||el.textContent)||'').replace(/\s+/g,' ').trim();

  function inputByLabel(words){
    for(const label of Array.from(document.querySelectorAll('label'))){
      const t=txt(label).toLowerCase();
      if(!words.some(w=>t.includes(w.toLowerCase())))continue;
      if(label.htmlFor){const el=document.getElementById(label.htmlFor);if(el)return el}
      const p=label.parentElement;if(p){const el=p.querySelector('input,select');if(el)return el}
    }
    return null;
  }
  function getGpu(){
    for(const id of ['gpuTarget','requiredGpu','gpuCount','targetGpu','gpu','needGpu']){const el=document.getElementById(id);if(el&&num(el.value)>0)return Math.round(num(el.value))}
    const by=inputByLabel(['필요 gpu','목표 gpu','gpu 수','gpu']);if(by&&num(by.value)>0)return Math.round(num(by.value));
    for(const id of ['mGpu','gpuResult','resultGpu']){const el=document.getElementById(id);if(el&&num(txt(el))>0)return Math.round(num(txt(el)))}
    for(const el of Array.from(document.querySelectorAll('.metric,.kpi,.result-card,.stat,.tile'))){const t=txt(el);if(/\bGPU\b/i.test(t)){const m=t.match(/([0-9][0-9,]*)/);if(m)return Number(m[1].replace(/,/g,''))}}
    return 0;
  }
  function getSystem(){
    for(const id of ['systemId','system','computeSystem','gpuSystem']){const el=document.getElementById(id);if(el&&el.tagName==='SELECT')return txt(el.options[el.selectedIndex])||el.value||'';if(el&&el.value)return String(el.value)}
    const by=inputByLabel(['시스템','system']);if(by){if(by.tagName==='SELECT')return txt(by.options[by.selectedIndex])||by.value||'';return String(by.value||'')}
    return '';
  }
  function getDistance(){
    for(const id of ['distServer','serverLeafDistance','serverToLeafDistance']){const el=document.getElementById(id);if(el&&num(el.value)!=null)return num(el.value)}
    const by=inputByLabel(['server->leaf','server ↔ leaf','server-leaf']);return by&&num(by.value)!=null?num(by.value):null;
  }
  function deepNumber(obj,patterns){
    const q=[obj],seen=new Set();
    while(q.length){const cur=q.shift();if(!cur||typeof cur!=='object'||seen.has(cur))continue;seen.add(cur);
      for(const [k,v] of Object.entries(cur)){const nk=String(k).toLowerCase().replace(/[^a-z0-9]/g,'');if(patterns.some(re=>re.test(nk))&&Number.isFinite(Number(v))&&Number(v)>0)return Number(v);if(v&&typeof v==='object')q.push(v)}
    } return null;
  }
  function inferredLinks(gpu){
    const d=state.lastDesign?deepNumber(state.lastDesign,[/nodeleaf.*(link|cable|connection|count)/,/serverleaf.*(link|cable|connection|count)/,/compute.*(link|cable).*count/,/node.*leaf/]):null;
    if(d)return{value:Math.round(d),source:'solver output'};
    for(const id of ['nodeLeafCableCount','serverLeafLinks','mServerLeafLinks','computeCableCount']){const el=document.getElementById(id);if(el&&num(txt(el)||el.value)>0)return{value:Math.round(num(txt(el)||el.value)),source:'solver UI'}}
    return{value:Math.max(1,gpu||1),source:'GPU proxy'};
  }
  function superpod(gpu){
    if(!gpu)return{label:'GPU 입력 대기',su:'-',refGpu:'-',leaf:'-',spine:'-',core:'-',note:'GPU 수를 입력하면 NVIDIA DGX SuperPOD reference-size class를 표시합니다.'};
    if(gpu<=248)return{label:'≈ 1 SU detailed class',su:1,refGpu:248,leaf:8,spine:4,core:0,note:'Detailed compute-fabric row: 31 DGX nodes / 248 GPUs; one system position is used for UFM connectivity.'};
    if(gpu<=504)return{label:'≈ 2 SU detailed class',su:2,refGpu:504,leaf:16,spine:8,core:0,note:'Detailed compute-fabric row: 63 nodes / 504 GPUs.'};
    if(gpu<=760)return{label:'≈ 3 SU detailed class',su:3,refGpu:760,leaf:24,spine:16,core:0,note:'Detailed compute-fabric row: 95 nodes / 760 GPUs.'};
    if(gpu<=1016)return{label:'≈ 4 SU detailed SuperPOD class',su:4,refGpu:1016,leaf:32,spine:16,core:0,note:'Detailed compute-fabric row: 127 nodes / 1,016 GPUs.'};
    if(gpu<=1024)return{label:'4 SU nominal SuperPOD class',su:4,refGpu:1024,leaf:32,spine:16,core:0,note:'Architecture scaling table uses 128 nodes / 1,024 GPUs for nominal 4-SU sizing.'};
    if(gpu<=2048)return{label:'8 SU SuperPOD class',su:8,refGpu:2048,leaf:64,spine:32,core:0,note:'Published larger-architecture row: 256 nodes / 2,048 GPUs.'};
    if(gpu<=4096)return{label:'16 SU SuperPOD class',su:16,refGpu:4096,leaf:128,spine:128,core:64,note:'At this published scale NVIDIA introduces a Core tier.'};
    if(gpu<=8192)return{label:'32 SU SuperPOD class',su:32,refGpu:8192,leaf:256,spine:256,core:128,note:'Published larger-architecture row: 1,024 nodes / 8,192 GPUs.'};
    if(gpu<=16384)return{label:'64 SU SuperPOD class',su:64,refGpu:16384,leaf:512,spine:512,core:256,note:'Published larger-architecture row: 2,048 nodes / 16,384 GPUs.'};
    return{label:'>64 SU / custom large-scale class',su:'>64',refGpu:'>16,384',leaf:'custom',spine:'custom',core:'custom',note:'Beyond the published table; treat as custom architecture sizing.'};
  }
  function mpoSpec(base,app){
    let chosen=base;
    if(chosen==='auto'){
      if(app==='400g-sr8')chosen='mpo16';
      else if(app==='nvidia-ndr400'||app==='400g-dr4')chosen='mpo12';
      else if(app==='400g-fr4'||app==='shortreach')chosen='duplex';
      else chosen='exact';
    }
    const map={
      mpo8:{
        label:'MPO-8 / 8F MTP — VERIFIED PRODUCT',
        connector:'8F MTP® connector · 8 installed fibers',
        installed:8,
        apps:'Base-8 structured cabling; 8-fiber parallel links only when the selected optic is explicitly compatible',
        breakout:'4 × LC duplex or 8F parallel mapping',
        product:'Corning EDGE8® MTP® Trunk · 8F · example SKU GE5E508QPNDDU100F',
        source:MPO_SOURCES.corning8,
        note:'Corning publicly lists EDGE8 trunks with 8-fiber pinned MTP connectors. Use only when the transceiver/interface is compatible with an 8F MTP presentation.'
      },
      mpo12:{
        label:'MPO-12 / 12F MTP — VERIFIED PRODUCT',
        connector:'12F MTP®/MPO connector · 12 installed fibers',
        installed:12,
        apps:'NDR/DR4 and other interfaces whose exact transceiver datasheet specifies MPO-12; general Base-12 structured cabling',
        breakout:'MPO-12 → LC duplex cassette/harness or 8-active-fiber parallel mapping as specified by the optic',
        product:'Corning EDGE™ MTP® Trunk · 12F · SKU G757512TPNDDU120F',
        source:MPO_SOURCES.corning12,
        note:'Corning publicly lists this 12F MTP-to-MTP trunk. For an 8-active-fiber optic, the physical connector can still be MPO-12 while four fibers are unused.'
      },
      mpo16:{
        label:'MPO-16 / 16F — VERIFIED PRODUCT',
        connector:'MPO-16 / MTP®-16 · 16 installed fibers · offset-key format',
        installed:16,
        apps:'400GBASE-SR8 and other 16-fiber / 8-lane parallel interfaces explicitly specified by the optic',
        breakout:'MPO-16 → 2 × 8F parallel channels only with a verified harness/adapter design',
        product:'SENKO MPO-16 Connector / US Conec MTP®-16 Connector family',
        source:MPO_SOURCES.senko16,
        note:'SENKO and US Conec both publish 16-fiber MPO/MTP connector products. Exact cable assembly and polarity must match the selected transceiver.'
      },
      duplex:{
        label:'Duplex optical interface',
        connector:'LC / CS / SN duplex — exact supported interface from optic datasheet',
        installed:2,
        apps:'FR4/LR4-class wavelength-multiplexed duplex optics',
        breakout:'2-fiber duplex channel',
        product:'No MPO product selected',
        source:'',
        note:'The transceiver face is duplex, so an MPO connector should not be recommended unless a verified cassette/backbone architecture is intentionally used.'
      },
      vsffMdc:{
        label:'VSFF · MDC duplex — VERIFIED PRODUCT',
        connector:'MDC duplex · 2 × 1.25 mm ferrules · 2F',
        installed:2,
        apps:'High-density duplex patching / breakout; use only when the selected transceiver, adapter or cassette explicitly supports MDC',
        breakout:'2F duplex; up to 3× LC-duplex density in supported hardware',
        product:'Corning MDC-to-MDC 2F jumper · representative product MUMU02QD120035M',
        source:MPO_SOURCES.corningMdc,
        note:'Corning publishes MDC cable assemblies and positions MDC as a VSFF duplex interface. Exact polish, fiber type and transceiver compatibility must match the selected design.'
      },
      vsffSn:{
        label:'VSFF · SN duplex — VERIFIED PRODUCT FAMILY',
        connector:'SENKO SN® duplex · 2 × 1.25 mm ferrules · 2F',
        installed:2,
        apps:'High-density Base-2 patching and 200G/400G/800G transceiver breakout where the exact interface supports SN',
        breakout:'2F duplex; SN can also be ganged for higher-density breakout',
        product:'SENKO SN® Series / SN connector family',
        source:MPO_SOURCES.senkoSn,
        note:'SENKO publishes SN as a VSFF duplex connector for high-density data-center and transceiver breakout applications.'
      },
      vsffCs:{
        label:'VSFF · CS duplex — VERIFIED PRODUCT',
        connector:'SENKO CS® duplex · 2F',
        installed:2,
        apps:'High-density duplex patching and supported 200G/400G/800G interfaces',
        breakout:'2F duplex; approximately 2× LC-duplex panel density in supported hardware',
        product:'SENKO CS® Standard Connector',
        source:MPO_SOURCES.senkoCs,
        note:'SENKO publishes CS as a VSFF connector standardized for high-density data-center connectivity. CS, SN and MDC are not mechanically interchangeable.'
      },
      exact:{
        label:'Exact transceiver required',
        connector:'No connector recommendation until the exact optic/interface is selected',
        installed:0,
        apps:'Generic 800G parallel or unspecified optics',
        breakout:'Pending exact optic, polarity and connector definition',
        product:'No product auto-selected',
        source:'',
        note:'Line rate alone is insufficient to choose MPO/MTP, LC or a VSFF interface. The tool withholds a connector recommendation until an exact product/interface is known.'
      }
    };
    return map[chosen]||map.exact;
  }
  function mpoUtil(mpo,active){
    if(!active||!mpo.installed)return 'N/A';
    if(mpo.label==='Duplex')return active===2?'100% at optic interface':'aggregation dependent';
    return Math.min(100,active/mpo.installed*100).toFixed(1)+'% direct-channel utilization';
  }
  function switchRows(){
    return SWITCH_CATALOG.map(s=>'<tr><td>'+s.vendor+'</td><td><b>'+s.model+'</b></td><td>'+s.ports+'</td><td>'+s.breakout+'</td><td>'+s.ru+'</td><td>'+s.typical+' / '+s.max+'</td><td>'+s.role+'</td><td><a href="'+s.url+'" target="_blank" rel="noopener">Cisco datasheet</a></td></tr>').join('');
  }
  function defaultApp(system,distance){const s=String(system||'').toUpperCase();if(/(B200|H200|H100)/.test(s))return'nvidia-ndr400';if(distance!=null&&distance<=3)return'shortreach';return'400g-dr4'}
  function spec(app){
    const m={
      'nvidia-ndr400':{name:'NVIDIA NDR 400G MMF reference',port:'DGX: QSFP112 / switch: twin-port OSFP',optic:'400G multimode parallel NDR',front:'MPO-12 APC harness / reference cable family',fiber:'MMF',fpl:8,chain:'DGX QSFP112 → 400G MMF parallel optic → MPO-12 APC harness/trunk → twin-port OSFP leaf/spine',note:'NVIDIA B200 components list multimode networking, MMF MPO12 APC split cables, QSFP112 and OSFP optics.'},
      '400g-dr4':{name:'400G DR4 structured',port:'QSFP-DD / OSFP',optic:'400G DR4',front:'MPO-12/APC (8 active fibers)',fiber:'OS2 SMF / bend-insensitive SMF',fpl:8,chain:'QSFP-DD/OSFP → DR4 → MPO-12/APC → OS2 structured trunk → MPO-12/APC → DR4',note:'Validate exact polarity, APC/UPC and reach against the selected transceiver datasheet.'},
      '400g-fr4':{name:'400G FR4 duplex',port:'QSFP-DD / OSFP',optic:'400G FR4',front:'LC duplex or supported VSFF adapter',fiber:'OS2 SMF',fpl:2,chain:'QSFP-DD/OSFP → FR4 → LC duplex → OS2 trunk/cassette → LC duplex → FR4',note:'Duplex optics reduce fiber count; transceiver cost/power/reach tradeoffs differ.'},
      '400g-sr8':{name:'400G SR8 parallel',port:'QSFP-DD / OSFP',optic:'400GBASE-SR8',front:'MPO-16',fiber:'OM4 MMF',fpl:16,chain:'QSFP-DD/OSFP → SR8 → MPO-16 → OM4 structured trunk → MPO-16 → SR8',note:'Corning identifies 400GBASE-SR8 as an application requiring MPO-16.'},
      '800g-parallel':{name:'800G parallel — optic-specific',port:'OSFP / QSFP-DD800',optic:'800G parallel optic',front:'MPO-16 or dual-MPO/MPO-12 depending optic',fiber:'OM4 or OS2 depending optic',fpl:16,chain:'800G host/switch port → optic-specific MPO interface → high-density trunk → MPO interface → peer optic',note:'Do not infer connector solely from the 800G line rate; verify the exact transceiver application.'},
      'shortreach':{name:'Short-reach electrical/AOC preference',port:'OSFP / QSFP-DD / QSFP112',optic:'DAC / AEC / AOC candidate',front:'Direct attach',fiber:'No structured fiber for DAC/AEC; AOC integrated',fpl:0,chain:'Device port → DAC/AEC/AOC → peer port',note:'For very short runs, direct-attach media may avoid a structured optical trunk if the platform supports it.'}
    }; return m[app]||m['400g-dr4'];
  }
  function bestTrunk(req,key,scope){
    const v=CATALOG[key];if(!v||!v.trunkSizes.length)return{size:null,count:null,family:v?v.family:'RFQ',note:v?v.note:''};
    const opts=(scope==='backbone'&&v.highCountSizes.length)?v.highCountSizes:v.pretermSizes;
    if(!opts||!opts.length)return{size:null,count:null,family:v.family,note:v.note};
    let size=opts.find(x=>x>=req);
    if(!size)size=Math.max(...opts);
    return{size,count:Math.ceil(req/size),family:v.family,note:v.note};
  }
  function autoVendor(req,scope){if(scope==='backbone'&&req>576)return'ls';if(req>288)return'corning';return'ls'}
  function vendorRows(req,scope){return Object.entries(CATALOG).map(([key,v])=>{let rec='RFQ';const b=bestTrunk(req,key,scope);if(b.size)rec=b.count+' × '+b.size+'F';else rec=(scope==='backbone'?'High-count family':'Data-hall / preterminated family')+' · exact F count RFQ';return'<tr><td><b>'+v.name+'</b></td><td>'+v.family+'</td><td>'+rec+'</td><td><a href="'+v.url+'" target="_blank" rel="noopener">Product</a> · <a href="'+v.home+'" target="_blank" rel="noopener">Homepage</a></td><td>'+v.connectors+'</td></tr>'}).join('')}

  function mount(){
    if(document.getElementById('optical-superpod-advisor'))return;
    const panels=Array.from(document.querySelectorAll('section,article,aside,div')).filter(el=>{const t=txt(el);return t.includes('설계 결과')&&t.includes('Leaf')&&t.includes('Spine')}).sort((a,b)=>a.querySelectorAll('*').length-b.querySelectorAll('*').length);
    const anchor=panels[0]||document.querySelector('.wrap,.container,main')||document.body;
    const box=document.createElement('section');box.id='optical-superpod-advisor';
    box.innerHTML=
      '<div class="osa-head"><div><h3>GPU Scale / Optical Connectivity Advisor</h3><div class="osa-sub">GPU 규모를 NVIDIA 공개 SuperPOD reference-size와 비교하고 optical application·connector·fiber count·vendor trunk 후보를 구체화합니다.</div></div><span class="osa-badge">v'+TOOL_VERSION+' · SOURCE-BASED</span></div>'+
      '<div class="osa-grid">'+
      '<div class="osa-card"><div class="osa-title">GPU 규모 · SuperPOD reference class</div><div class="osa-kpis">'+
      '<div class="osa-kpi"><div class="k">GPU</div><div class="v" id="osaGpu">-</div></div><div class="osa-kpi"><div class="k">Reference class</div><div class="v" id="osaClass">-</div></div><div class="osa-kpi"><div class="k">SU</div><div class="v" id="osaSu">-</div></div><div class="osa-kpi"><div class="k">Ref. GPU ceiling</div><div class="v" id="osaRefGpu">-</div></div></div>'+
      '<table><thead><tr><th>Leaf</th><th>Spine</th><th>Core</th><th>Source</th></tr></thead><tbody><tr><td id="osaLeaf">-</td><td id="osaSpine">-</td><td id="osaCore">-</td><td><a href="'+NVIDIA_SOURCE+'" target="_blank" rel="noopener">NVIDIA DGX SuperPOD RA</a></td></tr></tbody></table><div class="osa-note" id="osaSuperNote"></div></div>'+
      '<div class="osa-card"><div class="osa-title">Optical connectivity interface</div><div class="osa-form">'+
      '<div><label>Application</label><select id="osaApp"><option value="auto">Auto</option><option value="nvidia-ndr400">NVIDIA NDR 400G MMF</option><option value="400g-dr4">400G DR4 · 8F SMF</option><option value="400g-fr4">400G FR4 · 2F SMF</option><option value="400g-sr8">400G SR8 · 16F MMF</option><option value="800g-parallel">800G parallel · optic-specific</option><option value="shortreach">Short reach DAC/AEC/AOC</option></select></div>'+
      '<div><label>Preferred cable vendor</label><select id="osaVendor"><option value="auto">Auto</option><option value="ls">LS Cable & System</option><option value="hengtong">Hengtong</option><option value="yofc">YOFC</option><option value="lightera">Lightera</option><option value="sumitomo">Sumitomo Electric</option><option value="corning">Corning</option><option value="ztt">ZTT</option></select></div>'+
      '<div><label>Optical links / stage</label><input id="osaLinks" type="number" min="1" step="1"></div><div><label>Trunk scope</label><select id="osaScope"><option value="datahall">Data hall / preterminated</option><option value="backbone">Backbone / high-count</option></select></div><div><label>Connector product/interface</label><select id="osaMpoBase"><option value="auto">Auto — exact interface only</option><optgroup label="MPO / MTP"><option value="mpo8">MPO-8 / 8F MTP · Corning EDGE8 verified</option><option value="mpo12">MPO-12 / 12F MTP · Corning EDGE verified</option><option value="mpo16">MPO-16 / 16F · SENKO / US Conec verified</option></optgroup><optgroup label="Duplex / VSFF"><option value="duplex">LC duplex / supported duplex interface</option><option value="vsffMdc">VSFF · MDC · Corning verified</option><option value="vsffSn">VSFF · SN · SENKO verified</option><option value="vsffCs">VSFF · CS · SENKO verified</option></optgroup></select></div></div>'+
      '<div class="osa-chain" id="osaChain">-</div><table><tbody><tr><th>Port / optic</th><td id="osaPort">-</td></tr><tr><th>Connector</th><td id="osaConnector">-</td></tr><tr><th>Fiber</th><td id="osaFiber">-</td></tr><tr><th>Fibers / link</th><td id="osaFpl">-</td></tr><tr><th>Required fiber / equivalent FP</th><td id="osaFiberDemand">-</td></tr><tr><th>Recommended trunk</th><td id="osaTrunk">-</td></tr><tr><th>Connector product/interface</th><td id="osaMpoLabel">-</td></tr><tr><th>Verified product</th><td id="osaMpoProduct">-</td></tr><tr><th>Connector detail / installed fibers</th><td id="osaMpoConnector">-</td></tr><tr><th>Typical applications</th><td id="osaMpoApps">-</td></tr><tr><th>Breakout / migration</th><td id="osaMpoBreakout">-</td></tr><tr><th>Direct channel utilization</th><td id="osaMpoUtil">-</td></tr></tbody></table><div class="osa-note" id="osaOptNote"></div></div>'+
      '<div class="osa-card" style="grid-column:1/-1"><div class="osa-title">Switch equipment candidates · Cisco</div><table><thead><tr><th>Vendor</th><th>Model</th><th>Ports</th><th>Breakout</th><th>RU</th><th>Typical / Max power</th><th>Suggested role</th><th>Official</th></tr></thead><tbody id="osaSwitches"></tbody></table><div class="osa-note">Cisco entries are verified catalog candidates. The current topology solver should use them only after an exact solver profile is mapped; this table does not silently substitute another vendor profile.</div></div><div class="osa-card" style="grid-column:1/-1"><div class="osa-title">Multi-fiber cable candidates</div><table><thead><tr><th>Vendor</th><th>Representative optical cable</th><th>Calculated candidate</th><th>Official links</th><th>Connectivity</th></tr></thead><tbody id="osaVendors"></tbody></table><div class="osa-note">FP = total fibers ÷ 2 as an equivalent duplex-pair count. Parallel MPO links are engineered by active fiber count, so FP is shown only as a capacity/accounting aid. Exact trunk quantity must follow the physical link graph, route grouping, polarity, loss budget and spare policy.</div></div></div>';
    if(anchor.parentNode&&anchor!==document.body)anchor.parentNode.insertBefore(box,anchor.nextSibling);else anchor.appendChild(box);
    document.getElementById('osaLinks').addEventListener('input',()=>{state.userLinks=true;refresh()});
    document.getElementById('osaApp').addEventListener('change',refresh);
    document.getElementById('osaVendor').addEventListener('change',refresh);
    document.getElementById('osaScope').addEventListener('change',refresh);
    document.getElementById('osaMpoBase').addEventListener('change',refresh);
    refresh();
  }
  function refresh(){
    const box=document.getElementById('optical-superpod-advisor');if(!box)return;
    const gpu=getGpu(),system=getSystem(),distance=getDistance(),cls=superpod(gpu),inf=inferredLinks(gpu),linksInput=document.getElementById('osaLinks');
    if(!state.userLinks||state.lastGpu!==gpu){linksInput.value=inf.value;state.userLinks=false}state.lastGpu=gpu;
    const links=Math.max(1,Math.round(num(linksInput.value)||inf.value||1));
    let app=document.getElementById('osaApp').value;if(app==='auto')app=defaultApp(system,distance);const s=spec(app);
    const scope=document.getElementById('osaScope').value;const mpo=mpoSpec(document.getElementById('osaMpoBase').value,app);let vendorKey=document.getElementById('osaVendor').value;const req=s.fpl>0?links*s.fpl:0;if(vendorKey==='auto')vendorKey=autoVendor(req,scope);const vendor=CATALOG[vendorKey],tr=req>0?bestTrunk(req,vendorKey,scope):{size:null,count:null,family:'Direct attach'};
    document.getElementById('osaGpu').textContent=gpu?gpu.toLocaleString():'-';document.getElementById('osaClass').textContent=cls.label;document.getElementById('osaSu').textContent=cls.su;document.getElementById('osaRefGpu').textContent=typeof cls.refGpu==='number'?cls.refGpu.toLocaleString():cls.refGpu;
    document.getElementById('osaLeaf').textContent=cls.leaf;document.getElementById('osaSpine').textContent=cls.spine;document.getElementById('osaCore').textContent=cls.core;
    document.getElementById('osaSuperNote').innerHTML=cls.note+' <a href="'+NVIDIA_SOURCE+'" target="_blank" rel="noopener">Source</a>. Reference-size classification only; not NVIDIA certification of the current design.';
    document.getElementById('osaChain').textContent=s.chain;document.getElementById('osaPort').textContent=s.port+' · '+s.optic;document.getElementById('osaConnector').textContent=s.front;document.getElementById('osaFiber').textContent=s.fiber;document.getElementById('osaFpl').textContent=s.fpl||'Direct attach';
    document.getElementById('osaFiberDemand').textContent=req?req.toLocaleString()+'F / '+Math.ceil(req/2).toLocaleString()+' FP equivalent':'N/A for DAC/AEC/AOC direct attach';
    document.getElementById('osaTrunk').textContent=req?(tr.size?(vendor.name+' · '+tr.count+' × '+tr.size+'F · '+tr.family):(vendor.name+' · exact fiber count RFQ · '+tr.family)):'Direct attach media; structured trunk not required by this advisory.';
    document.getElementById('osaMpoLabel').textContent=mpo.label;
    document.getElementById('osaMpoProduct').innerHTML=mpo.source?('<a href="'+mpo.source+'" target="_blank" rel="noopener">'+mpo.product+'</a>'):mpo.product;
    document.getElementById('osaMpoConnector').textContent=mpo.installed?(mpo.connector+' · '+mpo.installed+'F installed/channel'):mpo.connector;
    document.getElementById('osaMpoApps').textContent=mpo.apps;
    document.getElementById('osaMpoBreakout').textContent=mpo.breakout;
    document.getElementById('osaMpoUtil').textContent=mpoUtil(mpo,s.fpl);
    document.getElementById('osaOptNote').innerHTML=s.note+' <b>Recommendation rule:</b> MPO-8/12/16 and VSFF MDC/SN/CS are shown only when a real manufacturer product or product family is verified; the selected transceiver/adapter must explicitly support that interface. Generic line-rate inference is not treated as a product recommendation. Exact trunk packing still depends on polarity, breakout mapping and cassette/harness design. <a href="'+MPO_SOURCES.corning8+'" target="_blank" rel="noopener">Corning MPO-8</a> · <a href="'+MPO_SOURCES.corning12+'" target="_blank" rel="noopener">Corning MPO-12</a> · <a href="'+MPO_SOURCES.senko16+'" target="_blank" rel="noopener">SENKO MPO-16</a> · <a href="'+MPO_SOURCES.corningMdc+'" target="_blank" rel="noopener">Corning MDC</a> · <a href="'+MPO_SOURCES.senkoSn+'" target="_blank" rel="noopener">SENKO SN</a> · <a href="'+MPO_SOURCES.senkoCs+'" target="_blank" rel="noopener">SENKO CS</a>. Links/stage source: <b>'+(state.userLinks?'user override':inf.source)+'</b>. Trunk scope: <b>'+(scope==='backbone'?'backbone/high-count':'data hall/preterminated')+'</b>. '+(distance!=null?('Server↔Leaf distance input: <b>'+distance+' m</b>. '):'')+'<a href="'+NVIDIA_COMPONENTS+'" target="_blank" rel="noopener">NVIDIA cable/transceiver example</a>.';
    document.getElementById('osaVendors').innerHTML=vendorRows(req||Math.max(144,gpu||144),scope);
    document.getElementById('osaSwitches').innerHTML=switchRows();
  }
  const nativeFetch=window.fetch;
  if(nativeFetch&&!window.__dcBomOpticalFetchWrapped){window.__dcBomOpticalFetchWrapped=true;window.fetch=async function(...args){const res=await nativeFetch.apply(this,args);try{const url=String(args[0]&&args[0].url?args[0].url:args[0]||'');if(url.includes('/api/design'))res.clone().json().then(data=>{state.lastDesign=data;setTimeout(refresh,0)}).catch(()=>{})}catch(_){}return res}}
  function start(){mount();document.addEventListener('change',()=>setTimeout(refresh,0),true);document.addEventListener('click',e=>{if(e.target&&(e.target.tagName==='BUTTON'||e.target.closest('button')))setTimeout(refresh,500)},true);let tries=0;const timer=setInterval(()=>{tries++;if(!document.getElementById('optical-superpod-advisor'))mount();refresh();if(tries>=20)clearInterval(timer)},500)}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start,{once:true});else start();
})();
