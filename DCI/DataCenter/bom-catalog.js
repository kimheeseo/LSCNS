(function (root) {
  'use strict';
  const BASE = 'DCI/DataCenter/product_catalog', REPO = 'kimheeseo/LSCNS';
  const api = 'https://api.github.com/repos/' + REPO;
  const text = x => String(x ?? ''), clean = x => text(x).toUpperCase().replace(/[®™]/g, '');
  const field = (s, keys) => keys.map(k => s[k]).find(x => x != null && text(x) !== '—') || '';
  function kind(x) {
    const t = clean([x.category, x.path, x.item, x.name].join(' '));

    if (/FIBER TEST AND MONITORING|\bOTDR\b|FIBERWATCH|ONMSI|FTH-5000|928-OMS|RTU-4000|RTU-4100/.test(t)) return 'monitoring';

    // DB-map taxonomy overrides: keep the source catalogs intact, but route
    // legacy folder names into the user-facing BOM DB groups.
    if (/FUSION SPLICER|SPLICER SOLUTIONS|\bSPLICERS?\b|90S\+|90R|S179\+|S124M16|S185/.test(t)) return 'splicer';
    if (/DATA CENTER GPUS?|AI ACCELERATORS?|\bGPU\b|INSTINCT MI\d+|GAUDI\s*3|ASCEND\s*9|DRAGONFLY AI|\bBR100\b/.test(t)) return 'gpu';
    if (/(^|[\/\s])CPU([\/\s]|$)|SERVER CPU|CPU AND SUPERCHIPS|\bEPYC\b|\bXEON\b|GRACE CPU|AMPEREONE/.test(t)) return 'cpu';
    if (/RIBBON BREAKOUT\s*&\s*FANOUT KITS?/.test(t)) return 'patch';
    if (/OPTICAL FIBERS?|POWER CABLE/.test(t)) return 'fiber';

    if (/BATTERY RACK|BATTERY CABINET/.test(t)) return 'battery';
    if (/TERMINATION BOX|SUB.?RACK|PANEL RACK MOUNT/.test(t)) return 'panel';
    if (/FUSECONNECT/.test(t)) return 'connector';
    if (/RJ45.*MODULAR PLUG|MODULAR PLUG.*RJ45/.test(t)) return 'connector';
    if (/(EDGE.*MODULE|FIBER.*MODULE|FIBRE.*MODULE|광.*모듈|CASSETTE)/.test(t) && !/TRANSCEIVER|OPTICAL MODULE/.test(t)) return 'module';
    if (/MDC\/MMC CABLING|MPO CABLING SYSTEM/.test(t)) return 'patch';
    if (/ACTIVE ELECTRICAL|\bAEC\b/.test(t)) return 'aec';
    if (/ACTIVE OPTICAL|\bAOC\b/.test(t) && !/TRANSCEIVERS AND AOC/.test(t)) return 'aoc';
    if (/\bDAC\b|DIRECT ATTACH/.test(t)) return 'dac';
    if (/TRANSFORMER|변압기/.test(t)) return 'transformer';
    if (/GENERATOR|발전기/.test(t)) return 'generator';
    if (/\bUPS\b|무정전/.test(t)) return 'ups';
    if (/CABLE MANAGEMENT|CABLE MANAGER|\bTRAYS?\b|\bDUCTS?\b|FIBERRUNNER|PATCHRUNNER|트레이|덕트|케이블 관리/.test(t)) return 'management';
    if (/GANG CLIP|MODULAR JACK|RJ45 PLUG|RJ45.*PLUGS.*JACKS/.test(t)) return 'component';
    if (/TRANSCEIVER|OPTICAL MODULE|광모듈|트랜시버/.test(t)) return 'transceiver';
    if (/CLEAN|FERRULE|DSP|\bPIC\b|TIA|LASER|OPTICAL CHIP/.test(t)) return 'component';
    if (/ADAPTER|어댑터/.test(t) && !/CONNECTX|NETWORK ADAPTER/.test(t)) return 'adapter';
    if (/CONNECTOR|커넥터|종단/.test(t) && !/BOARD CONNECTOR/.test(t)) return 'connector';
    if (/PATCH PANEL|PANEL|HOUSING|ENCLOSURE|패치패널|패널/.test(t) && !/RACK|서버 랙/.test(t)) return 'panel';
    if (/\bNIC\b|CONNECTX|SUPERNIC|NETWORK ADAPTER/.test(t)) return 'nic';
    if (/SWITCH|QUANTUM|SPECTRUM|LEAF|SPINE|스위치/.test(t)) return 'switch';
    if (/GPU SERVER|RACK SERVER|서버|COMPUTE TRAY/.test(t)) return 'server';
    if (/\bRACKS?\b|IT 랙/.test(t)) return 'rack';
    if (/COPPER|CAT.?6|RJ45|동 케이블/.test(t)) return 'copper';
    if (/TRUNK|트렁크/.test(t)) return 'trunk';
    if (/JUMPER|PATCH CORD|HARNESS|ASSEMBL|패치|ASSEMBLY/.test(t)) return 'patch';
    if (/FIBER.*CABLE|OPTIC.*CABLE|광 케이블|광케이블/.test(t)) return 'fiber';
    return 'other';
  }
  function normalize(catalog, path) {
    return Object.entries(catalog.products || {}).filter(([, meta]) => !meta.vendorReferenceOnly).map(([id, meta]) => {
      const s = {...catalog.defaultSpecs, ...meta.specs};
      const p = {id, path, vendor: catalog.company || path.split('/')[0], category: catalog.category || '', name: meta.name || id, description: meta.description || '', specs: s, source: meta.businessUrl || meta.officialUrl || meta.url || catalog.officialUrl || '', checked: meta.checked || catalog.checked || '', referenceOnly: !!meta.referenceOnly};
      // A mixed folder must be classified by the actual product, not by its parent label.
      p.kind = /Fiber Test and Monitoring/i.test(p.category) ? 'monitoring' : kind({...p, category: '', path: ''});
      if (p.kind === 'other' && !/Optical Connectivity and Rack Enclosures/i.test(p.category)) p.kind = kind(p);
      p.speed = field(s, ['Data Rate', '총 속도', '속도', 'Speed', 'Bandwidth']) || (rates(p.name).length ? p.name : '');
      p.length = field(s, ['길이', 'Reach', '거리', 'Length', 'Cable Length']);
      p.connector = field(s, ['Connector', '커넥터', 'Optical Interface', 'Interface']);
      if(!p.connector){const connectorName=p.name.match(/\b(MPO(?:-?\d+)?|MTP(?:-?\d+)?|MDC|MMC|LC|SC|SN-MT)\b/i);if(connectorName)p.connector=connectorName[0];}
      p.package = field(s, ['Package', '폼팩터', 'Form Factor', '포트']) || (p.name.match(/\b(?:QSFP-DD|QSFP28|QSFP56|OSFP-XD|OSFP|CFP\d?|SFP\+?)\b/i)||[''])[0];
      p.standard = field(s, ['광 규격', 'Standard', 'Remark', '규격', 'Standards / Notes']);
      if (!/DR\d|FR\d|LR\d|SR\d|ER\d|BASE-|\bLX\b|\bSX\b|\bZR\b/i.test(p.standard)) p.standard=(p.name.match(/\b(?:\d+GBASE-)?(?:DR\d|FR\d|LR\d|SR\d|ER\d|ZR)\b/i)||[''])[0];
      p.fiber = field(s, ['Fiber Type', 'Fiber Category', '광섬유', 'Fiber']);
      p.fibers = field(s, ['Fiber Count', '심수']);
      p.protocol = field(s, ['Protocol', '프로토콜']);
      if(!p.protocol){const protocol=p.name.match(/InfiniBand|Ethernet|PCIe(?:\s+Gen\d)?|(?:Mini-)?SAS|NVLink/i);if(protocol)p.protocol=/SAS/i.test(protocol[0])?'SAS':protocol[0];}
      p.polarity = field(s, ['Polarity', '극성']); p.gender = field(s, ['Gender']);
      p.jacket = field(s, ['Jacket', '난연 등급', 'Flame Rating']);
      p.status = field(s, ['Status', 'Lifecycle', 'Product Status']);
      p.discontinued = /DISCONTINUED|OBSOLETE|END OF LIFE|\bEOL\b/i.test(text(p.status));
      p.optical=/FIBER|FIBRE|OPTIC|MPO|MTP|MDC|MMC|SENKO|CORNING|US.?CONEC|\bLC\b|\bSC\b/i.test([p.name,p.category,p.path,p.fiber,p.connector].join(' '));
      return p;
    }).filter(p => /^https?:\/\//i.test(p.source));
  }
  function distance(value) {const m = text(value).match(/(\d+(?:\.\d+)?)\s*(km|m)(?!m)/i); return m ? Number(m[1]) * (m[2].toLowerCase() === 'km' ? 1000 : 1) : null;}
  function rates(value) {return [...text(value).matchAll(/(\d+(?:\.\d+)?)\s*(T|G)(?:b|\b|\/)/ig)].map(m => Number(m[1]) * (m[2].toUpperCase() === 'T' ? 1000 : 1));}
  function connector(value) {return clean(value).replace(/MTP/g, 'MPO').replace(/MPO12/g, 'MPO-12').replace(/MPO16/g, 'MPO-16').replace(/\s+/g, '');}
  function match(row, products) {
    const q = row.requirement || row.profile || {}, media = clean(row.media || q.media), itemKind = kind(row);
    let type = itemKind;
    if (row.transceiver || (row.unit === 'module' && /TRANSCEIVER|광모듈|PLUGGABLE/.test(clean(row.item)))) type = 'transceiver';
    else if (row.cable || /CABLE|케이블/.test(clean(row.item))) {
      if (/\bAEC\b/.test(media)) type = 'aec'; else if (/\bDAC\b/.test(media)) type = 'dac'; else if (/\bAOC\b/.test(media)) type = 'aoc';
      else if (/COPPER|RJ45|BASE-T/.test(media + ' ' + clean(q.connector))) type = 'copper';
      else if (type === 'other') type = 'patch';
    }
    const required = {speed: Number(q.speed || rates(row.spec || q.media)[0]) || null, length: Number(row.lengthM ?? q.lengthM) || null, connector: q.connector || '', package: q.package || '', standard: q.standard || '', fiber: q.fiberType || q.fiber || '', fibers: row.fibers || 0, protocol: q.protocol || '', polarity: row.polarity || '', gender: row.gender || '', jacket: row.jacket || ''};
    if (/integrated|未|미확정/i.test(required.package)) required.package='';
    const comparable = (a, b) => clean(a).replace(/BASE-/g, '').replace(/[\s_]/g, '') === clean(b).replace(/BASE-/g, '').replace(/[\s_]/g, '');
    return products.flatMap(p => {
      if (p.discontinued) return [];
      if (p.kind === 'monitoring') return []; // Installation/operations references never replace cable or equipment BOM lines.
      const compatibleKinds = type === 'fiber' ? ['fiber','trunk','patch'] : type === 'patch' ? ['patch'] : [type];
      if (!compatibleKinds.includes(p.kind) || type === 'other' || type === 'component') return [];
      const opticalItem=/SMF|MMF/.test(media)||/MPO|MTP|LC|MDC|MMC/.test(clean(required.connector))||!!required.fiber||type==='module';
      if(opticalItem&&['fiber','patch','trunk','connector','adapter','panel','module'].includes(type)&&!p.optical)return [];
      const confirmed = [], missing = []; let conflict = false;
      const check = (name, need, have, predicate) => {if (!need || /^(REVIEW|PROJECT|N\/A|未)/.test(clean(need))) return; if (!have || /검증 필요|미공개|미정|RFQ|확인 필요/.test(text(have))) missing.push(name); else if (predicate(need, have)) confirmed.push(name); else conflict = true;};
      if (['transceiver','dac','aec','aoc'].includes(type)) {
        check('속도', required.speed, rates(p.speed).length ? p.speed : '', (a,b) => rates(b).includes(a));
        check('길이/거리', required.length, distance(p.length) == null ? '' : p.length, (a,b) => ['transceiver'].includes(type) ? distance(b) >= a : distance(b) >= a && (/최대|UP TO|제품군|FAMILY/i.test(b) || Math.abs(distance(b) - a) < 0.01));
        check('폼팩터', required.package, p.package, comparable);
        check('프로토콜', required.protocol, p.protocol, (a,b) => clean(b).includes(clean(a)));
      }
      if (type === 'transceiver') check('광 규격', required.standard, p.standard, (a,b) => comparable(a,b) || comparable(clean(a).replace(/^\d+G(?:BASE-)?/,''),clean(b).replace(/^\d+G(?:BASE-)?/,'')));
      if (['transceiver','fiber','patch','trunk','connector','adapter','panel','module','copper'].includes(type)) {
        check('커넥터', required.connector, p.connector, (a,b) => {
          const left=connector(a),right=connector(b);
          if(left.includes('RJ45')&&right.includes('RJ45')) return true;
          if(left.split('/')[0]===right.split('/')[0]&&(!left.includes('/')||!right.includes('/'))){missing.push('커넥터 연마/핀');return true;}
          return left===right;
        });
        check('광섬유', required.fiber, p.fiber, (a,b) => {const sm = /OS2|SMF|SINGLE.?MODE/.test(clean(a)); return sm ? /OS2|SMF|SINGLE.?MODE|G\.65[27]/.test(clean(b)) : clean(b).includes(clean(a));});
        check('심수', required.fibers, p.fibers, (a,b) => text(b).trim() === text(a));
        check('극성', required.polarity, p.polarity, comparable);check('핀/gender',required.gender,p.gender,comparable);check('외피/난연',required.jacket,p.jacket,(a,b)=>clean(b).includes(clean(a)));
      }
      if (['fiber','patch','trunk','copper'].includes(type)) check('길이',required.length,distance(p.length)==null?'':p.length,(a,b)=>/최대|UP TO|제품군|FAMILY/i.test(b)?distance(b)>=a:Math.abs(distance(b)-a)<0.01);
      if (['switch','nic'].includes(type)) check('프로토콜', required.protocol, p.protocol, (a,b)=>clean(b).includes(clean(a)));
      if (conflict) return [];
      if (!required.speed && ['transceiver','dac','aec','aoc'].includes(type)) missing.push('요구 속도');
      if (type === 'transceiver' && !required.protocol) missing.push('요구 프로토콜');
      if (type === 'transceiver' && !required.standard) missing.push('요구 광 규격');
      if (type === 'transceiver' && !required.package) missing.push('요구 폼팩터');
      if (type === 'transceiver' && !required.connector) missing.push('요구 커넥터');
      if (['ups','generator','transformer','rack','server','switch','nic','management','panel'].includes(type)) missing.push('용량/구성·현장 적합성');
      if (['patch','fiber','trunk','connector','adapter','module','copper'].includes(type)) missing.push(type === 'module' ? 'housing slot·극성·loss budget 최종 확인' : '길이·극성·핀·현장 구성 최종 확인');
      if (p.referenceOnly) missing.push('제품군/SKU 확인');
      const status = missing.length ? '관련 제품 · 검증 필요' : '규격 후보 · 장비 호환 검증 필요';
      return [{...p, evidence: confirmed.length ? '일치: ' + confirmed.join('·') : '품목 분류 일치', matchBasis: (confirmed.length ? '일치: ' + confirmed.join('·') + '; ' : '') + (missing.length ? '미확정: ' + [...new Set(missing)].join('·') : 'host/FEC/광손실·극성 확인 필요'), status, score: confirmed.length * 10 - missing.length}];
    }).sort((a,b) => b.score - a.score || a.vendor.localeCompare(b.vendor) || a.name.localeCompare(b.name));
  }
  let cached, pending, epoch = 0;
  async function load(force = false) {
    if (force) {epoch++;cached = null;pending = null;}
    if (cached && Date.now() - cached.loadedAt < 300000) return cached;
    if (pending) return pending;
    const generation = epoch;
    const json = async url => {const r = await fetch(url,{cache:'no-store'});if(!r.ok)throw Error('카탈로그 HTTP '+r.status);return r.json();};
    pending = (async () => {
      let revision,paths,snapshot=false,browserBase=null;
      if(typeof location!=='undefined'){
        const manifest=await json(new URL('product_catalog/catalog-manifest.json',location.href));
        if(!Array.isArray(manifest.paths))throw Error('배포 카탈로그 목록 오류');
        revision=manifest.revision||'main';paths=manifest.paths.map(path=>({path}));snapshot=true;
        browserBase=new URL('product_catalog/',location.href);
      }else{
        const ref=await json(api+'/git/ref/heads/main');revision=ref.object.sha;
        const folder=(await json(api+'/contents/DCI/DataCenter?ref='+revision)).find(x=>x.name==='product_catalog');
        if(!folder)throw Error('product_catalog 폴더 없음');
        const tree=await json(api+'/git/trees/'+folder.sha+'?recursive=1');
        if(tree.truncated)throw Error('카탈로그 파일 목록이 불완전합니다.');
        paths=tree.tree.filter(x=>x.type==='blob'&&/(^|\/)catalog\.json$/i.test(x.path));
      }
      const result={revision,snapshot,loadedAt:Date.now(),catalogs:[],products:[],errors:[]};let cursor=0;
      await Promise.all(Array.from({length:6},async()=>{while(cursor<paths.length){const path=paths[cursor++].path;try{const catalogUrl=browserBase?new URL(path.split('/').map(encodeURIComponent).join('/'),browserBase):'https://raw.githubusercontent.com/'+REPO+'/'+revision+'/'+BASE+'/'+path.split('/').map(encodeURIComponent).join('/');const catalog=await json(catalogUrl);result.catalogs.push({path,catalog});result.products.push(...normalize(catalog,path));}catch(e){result.errors.push({path,error:e.message});}}}));
      const unique = new Map();for(const p of result.products){const k=p.vendor+'\0'+p.source+'\0'+p.id;if(!unique.has(k))unique.set(k,p);}result.products=[...unique.values()];
      if (generation === epoch) {cached=result;pending=null;}
      return result;
    })().catch(e=>{if(generation===epoch)pending=null;throw e;});
    return pending;
  }
  const model={kind,normalize,match,load,distance,rates};root.DCBomCatalog=model;
  if(typeof module!=='undefined'&&module.exports)module.exports=model;
})(typeof window!=='undefined'?window:globalThis);
