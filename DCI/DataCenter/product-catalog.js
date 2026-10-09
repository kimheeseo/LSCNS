(() => {
  'use strict';

  const state = { company: '', category: '', products: [], manifest: null, selectedFamilies: new Set(), viewMode: 'cards', dbMapData: null };
  const tourParams = new URLSearchParams(location.search);
  const tourFocus = (tourParams.get('tourFocus') || '').trim();
  let tourSuggestionsApplied = false;
  const $ = id => document.getElementById(id);
  const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const humanBytes = value => {
    const n = Number(value) || 0;
    if (n < 1024) return n + ' B';
    if (n < 1024 * 1024) return (n / 1024).toFixed(0) + ' KB';
    return (n / 1024 / 1024).toFixed(1) + ' MB';
  };

  let catalogIndex = null;

  const staticUrl = path => new URL(
    'product_catalog/' + path.split('/').map(encodeURIComponent).join('/'),
    location.href
  ).toString();

  async function loadCatalogIndex(force) {
    if (catalogIndex && !force) return catalogIndex;
    const url = new URL('product_catalog/catalog-manifest.json', location.href);
    if (force) url.searchParams.set('refresh', Date.now());
    const response = await fetch(url, {cache:'no-store'});
    if (!response.ok) throw new Error('배포 카탈로그 목록을 불러오지 못했습니다. HTTP ' + response.status);
    const manifest = await response.json();
    if (!manifest || !Array.isArray(manifest.paths)) throw new Error('배포 카탈로그 목록 형식이 올바르지 않습니다.');
    catalogIndex = manifest;
    return catalogIndex;
  }

  function catalogPathParts(path) {
    return String(path || '').split('/').filter(Boolean);
  }

  function pathsFor(company, category) {
    const paths = (catalogIndex && catalogIndex.paths) || [];
    return paths.filter(path => {
      const parts = catalogPathParts(path);
      if (company && parts[0] !== company) return false;
      if (category && parts[1] !== category) return false;
      return /catalog\.json$/i.test(path);
    });
  }

  async function fetchCatalog(path, force) {
    const url = new URL(staticUrl(path));
    if (force) url.searchParams.set('refresh', Date.now());
    const response = await fetch(url, {cache:'no-store'});
    if (!response.ok) throw new Error('제품 카탈로그를 불러오지 못했습니다. HTTP ' + response.status);
    return response.json();
  }

  async function walkProducts(company, category, force) {
    await loadCatalogIndex(force);
    const paths = pathsFor(company, category);
    let cursor = 0;
    const entries = [];
    const errors = [];
    await Promise.all(Array.from({length:6}, async () => {
      while (cursor < paths.length) {
        const path = paths[cursor++];
        try {
          const manifest = await fetchCatalog(path, force);
          const parts = catalogPathParts(path);
          const relativeParts = parts.slice(2, -1);
          const group = manifest.displayName || relativeParts.join(' › ') || category;
          Object.entries(manifest.products || {}).forEach(([key, meta]) => {
            if (!meta) return;
            const href = meta.businessUrl || meta.officialUrl || meta.url || manifest.officialUrl || '';
            entries.push({
              file: {name:key, size:0, html_url:href, virtual:true},
              manifest,
              group: String(meta.specs?.['Catalog Group'] || group),
              catalogPath:path
            });
          });
        } catch (error) {
          errors.push({path, error:error.message});
        }
      }
    }));
    if (!entries.length && errors.length) throw new Error('제품 카탈로그 정적 파일을 불러오지 못했습니다.');
    return {entries, errors};
  }

  function shell() {
    return '<section class="panel catalogPanel">' +
      '<div class="catalog-head"><div><h2>업체별 부품 리스트</h2><p>GitHub의 <code>product_catalog</code> 폴더를 기준으로 업체 → 부품군 → 하위 제품군 → 제품 자료를 탐색합니다. 하위 폴더가 여러 단계여도 PDF를 자동 탐색하며, 각 폴더의 <code>catalog.json</code>에 등록된 핵심 사양은 제품 카드에 함께 표시됩니다.</p></div>' +
      '<div class="catalog-head-actions"><button type="button" id="catalogKoreaBtn" class="catalog-refresh" aria-expanded="false">한국 업체·대리점 조회</button><button type="button" id="catalogRefresh" class="catalog-refresh">카탈로그 새로고침</button></div></div>' +
      '<div class="catalog-steps"><span class="active">1 업체</span><span>2 부품군</span><span>3 제품·스펙</span></div>' +
      '<div id="catalogNotice" class="catalog-notice">카탈로그를 불러오는 중입니다.</div>' +
      '<section class="catalog-db-map" aria-labelledby="catalogDbMapTitle"><div class="catalog-db-map-head"><div><h3 id="catalogDbMapTitle">제품 DB 맵</h3><p>부품군별 등록 제품 수를 한눈에 보고, 원하는 DB를 눌러 업체·세부 부품으로 이동합니다.</p></div><span id="catalogDbMapStatus">DB 집계 중…</span></div><div id="catalogDbMap" class="catalog-db-map-body"><div class="catalog-db-loading">구조화 제품 DB를 집계하고 있습니다.</div></div><div id="catalogDbDetail" class="catalog-db-detail" hidden></div></section>' +
      '<div id="catalogTourFocus" class="catalog-tour-focus" hidden><b>Gold Tour 연결</b><span id="catalogTourFocusText"></span><div id="catalogTourSuggestions" class="catalog-tour-suggestions"></div></div>' +
      '<section id="catalogKoreaPanel" class="catalog-korea-panel" hidden><div class="catalog-korea-head"><div><h3>한국 업체·대리점 / 국내 문의처</h3><p>업체별 부품 리스트 제조사의 한국 법인·공급 파트너·공식 문의 경로 포함 · 2026-10-06 확인</p></div></div><div id="catalogKoreaBody" class="catalog-korea-body"><p class="catalog-empty">업체 정보를 불러오는 중입니다.</p></div></section>' +
      '<div class="catalog-layout">' +
        '<section class="catalog-column"><div class="catalog-column-head"><h3>1. 업체</h3><span id="catalogCompanyCount">—</span></div><div id="catalogCompanies" class="catalog-list"></div></section>' +
        '<section class="catalog-column"><div class="catalog-column-head"><h3>2. 부품군</h3><span id="catalogCategoryCount">—</span></div><div id="catalogCategories" class="catalog-list"><p class="catalog-empty">업체를 선택하세요.</p></div></section>' +
        '<section class="catalog-products-column"><div class="catalog-products-head"><div><h3>3. 제품·간략 스펙</h3><p id="catalogContext">부품군을 선택하면 등록된 PDF·공식 URL 제품이 표시됩니다.</p></div><div class="catalog-products-tools"><div id="catalogViewMode" class="catalog-view-mode" role="group" aria-label="제품 정리 방식"><button type="button" data-catalog-view="cards" class="active" aria-pressed="true">카드형</button><button type="button" data-catalog-view="table" aria-pressed="false">표형</button></div><input id="catalogSearch" type="search" placeholder="제품명 / 모델 / 사양 검색" disabled></div></div><div id="catalogFamilyFilters" class="catalog-family-filters" hidden></div><div id="catalogProducts" class="catalog-products"><p class="catalog-empty">제품 자료를 선택하세요.</p></div></section>' +
      '</div>' +
      '<div class="catalog-foot">폴더 추가만으로 업체·부품군·하위 제품군·PDF/URL 제품 목록이 자동 반영됩니다. 상세 스펙은 해당 제품군 폴더 또는 상위 부품 폴더의 catalog.json으로 관리합니다.</div>' +
    '</section>';
  }

  async function loadKoreaSuppliers() {
    const host = $('catalogKoreaBody');
    if (!host || host.dataset.loaded === 'true') return;
    try {
      const response = await fetch('./korea-suppliers.json', {cache:'no-store'});
      if (!response.ok) throw new Error('HTTP ' + response.status);
      const data = await response.json();
      const rows = (data.suppliers || []).map(x =>
        '<tr><td><b>' + esc(x.manufacturer || x.company) + '</b></td>' +
        '<td><b>' + esc(x.company) + '</b><small>' + esc(x.note || '') + '</small></td>' +
        '<td><span class="catalog-relation">' + esc(x.relationship || '국내 문의처') + '</span></td>' +
        '<td>' + esc(x.category) + '</td>' +
        '<td><a href="' + esc(x.website) + '" target="_blank" rel="noopener">공식 URL ↗</a></td>' +
        '<td>' + (x.email ? '<a href="mailto:' + esc(x.email) + '">' + esc(x.email) + '</a>' : '—') + '</td>' +
        '<td>' + (x.phone ? '<a href="tel:' + esc(String(x.phone).replace(/[^+0-9]/g,'')) + '">' + esc(x.phone) + '</a>' : '—') +
        (x.contactUrl ? '<br><a class="catalog-contact-link" href="' + esc(x.contactUrl) + '" target="_blank" rel="noopener">문의 페이지 ↗</a>' : '') + '</td></tr>'
      ).join('');
      host.innerHTML = '<div class="catalog-sheet-wrap"><table class="catalog-sheet catalog-korea-table"><thead><tr><th>카탈로그 제조사</th><th>한국 업체 / 문의처</th><th>관계</th><th>주요 분야</th><th>URL</th><th>이메일</th><th>연락처</th></tr></thead><tbody>' + rows + '</tbody></table></div>';
      host.dataset.loaded = 'true';
    } catch (error) {
      host.innerHTML = '<p class="catalog-empty">한국 업체 정보를 불러오지 못했습니다: ' + esc(error.message) + '</p>';
    }
  }

  function notice(text, mode) {
    const el = $('catalogNotice');
    if (!el) return;
    el.textContent = text || '';
    el.className = 'catalog-notice' + (mode ? ' ' + mode : '');
  }

  function setStep(step) {
    document.querySelectorAll('.catalog-steps span').forEach((el, index) => {
      el.classList.toggle('active', index < step);
    });
  }

  function cleanLabel(value) {
    return String(value || '').replace(/^Connecttor$/i, 'Connector').replace(/^USConnec$/i, 'US Conec');
  }

  function button(label, type, selected) {
    const shown = cleanLabel(label);
    return '<button type="button" class="catalog-select' + (selected ? ' selected' : '') + '" data-' + type + '="' + esc(label) + '" aria-pressed="' + String(!!selected) + '">' +
      '<span>' + esc(shown) + '</span><b>›</b></button>';
  }

  const dbKindLabels = {
    fiber:'Optical Fiber Cable', trunk:'Trunk', patch:'Patch / Harness',
    aoc:'AOC', dac:'DAC', aec:'AEC', copper:'Copper / Cat6',
    connector:'Connector', adapter:'Adapter', module:'Module / Cassette', panel:'Panel / Housing',
    transceiver:'Transceiver', component:'Optical / Electronic Component',
    switch:'Switch', nic:'NIC / Network Adapter', server:'Server', cpu:'CPU', storage:'Storage / SSD',
    gpu:'GPU / AI Accelerator',
    ups:'UPS', generator:'Generator', transformer:'Transformer', battery:'Battery', power:'Power Distribution',
    rack:'Rack', management:'Cable Management', cabletray:'Cable Tray', raceway:'Fiber Raceway', cooling:'Cooling', facility:'Data Center Infrastructure',
    splicer:'Splicer', monitoring:'Fiber Test / Monitoring', other:'Other'
  };

  const dbGroups = [
    {id:'cable', label:'Cable / Interconnect DB', kinds:['fiber','trunk','patch','aoc','dac','aec','copper']},
    {id:'connectivity', label:'Connector / Panel / Adapter DB', kinds:['connector','adapter','module','panel']},
    {id:'optics', label:'Optics / Component DB', kinds:['transceiver','component']},
    {id:'network', label:'Network / Server DB', kinds:['switch','nic','server','cpu','storage']},
    {id:'accelerator', label:'Accelerator DB', kinds:['gpu']},
    {id:'power', label:'Power DB', kinds:['ups','generator','transformer','battery','power']},
    {id:'infra', label:'Rack / Infrastructure DB', kinds:['rack','management','cabletray','raceway','cooling','facility']},
    {id:'splicer', label:'Splicer DB', kinds:['splicer']},
    {id:'monitoring', label:'Fiber Test / Monitoring DB', kinds:['monitoring']},
    {id:'other', label:'Other DB', kinds:['other']}
  ];

  function dbPathParts(product) {
    return String(product?.path || '').split('/').filter(Boolean);
  }

  const dbTaxonomyCache = new WeakMap();

  function dbTaxonomy(product) {
    if (product && typeof product === 'object' && dbTaxonomyCache.has(product)) {
      return dbTaxonomyCache.get(product);
    }

    const path = String(product?.path || '');
    const vendor = String(product?.vendor || dbPathParts(product)[0] || '');
    const category = String(product?.category || '');
    const name = String(product?.name || '');
    const description = String(product?.description || '');
    const t = [path, vendor, category, name, description].join(' ').toUpperCase();
    const raw = String(product?.kind || 'other');

    const hit = (kind, reason) => {
      const result = {kind, reason};
      if (product && typeof product === 'object') dbTaxonomyCache.set(product, result);
      return result;
    };

    if (/FIBER TEST AND MONITORING|\bOTDR\b|FIBERWATCH|ONMSI|FTH-5000|928-OMS|RTU-4000|RTU-4100/.test(t)) return hit('monitoring','fiber test/monitoring reference equipment');

    // 8) Splicer — explicit folder/product family takes priority.
    if (/FUSION[ _/-]*SPLICER|SPLICER SOLUTIONS|\\bSPLICERS?\\b|90S\\+|90R(?:4|12|16)?|S179\\+|S124M16|S185(?:EDV)?|THERMAL JACKET REMOVER|FIBER PROTECTION SLEEVE/.test(t)) {
      return hit('splicer','fusion-splicing product/tool');
    }

    // Cable pathway split for data-center infrastructure.
    // Keep this in the DB-map layer so BOM matching remains backward compatible.
    if (/FIBERGUIDE|FIBERRUNNER|FIBER RACEWAY|FIBRE RACEWAY/.test(t)) {
      return hit('raceway','fiber raceway');
    }
    if (/CABLE TRAY|WIRE BASKET|CABLE RUNWAY|KWIKRAIL|CABLOFIL|G-TRAY|G MINI/.test(t)) {
      return hit('cabletray','cable tray');
    }

    // Current catalog top-level taxonomy. Map all known catalog families
    // into the eight primary DB groups before product-name fallbacks.
    const topCategory = String(dbPathParts(product)[1] || category || '').toUpperCase();
    const fullPath = path.toUpperCase();

    if (/^(FUSION SPLICER SOLUTIONS|FUSION SPLICERS|SPLICE TRAYS)$/.test(topCategory)) {
      return hit('splicer','catalog family: splicer');
    }
    if (/^(GPU ACCELERATORS|DATA CENTER GPUS|AI ACCELERATORS|AI FACTORY PLATFORMS|PROFESSIONAL GPUS REFERENCE|GPU SERVERS|CLOUD TPU)$/.test(topCategory)) {
      return hit('gpu','catalog family: accelerator');
    }
    if (/^(ENTERPRISE SSD|DATA CENTER SSD|ENTERPRISE SSD & CONTROLLERS|SSD CONTROLLERS|ENTERPRISE SSD CONTROLLERS)$/.test(topCategory)) {
      return hit('storage','catalog family: storage / SSD');
    }
    if (/^(CPU|CPU PORTFOLIO|CPU AND SUPERCHIPS|DATA CENTER SWITCHES|DATA CENTER SWITCHING|NETWORK|NETWORKING|RACK SERVERS|SWITCHES NICS|INFINIBAND XDR NDR HDR)$/.test(topCategory)) {
      if (/CPU/.test(topCategory)) return hit('cpu','catalog family: CPU');
      if (/SWITCH|NETWORK|INFINIBAND/.test(topCategory)) return hit('switch','catalog family: network');
      return hit('server','catalog family: server');
    }
    if (/^(UPS|GENERATORS|TRANSFORMERS|DATA CENTER POWER|POWER CONNECTORS AND BUSBAR|POWER DISTRIBUTION PANELS)$/.test(topCategory)) {
      if (topCategory === 'UPS') return hit('ups','catalog family: UPS');
      if (topCategory === 'GENERATORS') return hit('generator','catalog family: generator');
      if (topCategory === 'TRANSFORMERS') return hit('transformer','catalog family: transformer');
      return hit('power','catalog family: power');
    }
    if (/^(CABLE MANAGEMENT|CABLE TRAY|FIBER RACEWAY|LIQUID COOLING|RACK INFRASTRUCTURE|RACK MOUNT|RACKS|DATA CENTER INFRASTRUCTURE|FIBERGUIDE)$/.test(topCategory)) {
      if (topCategory === 'FIBERGUIDE' || topCategory === 'FIBER RACEWAY') return hit('raceway','catalog family: fiber raceway');
      if (topCategory === 'CABLE TRAY') return hit('cabletray','catalog family: cable tray');
      if (topCategory === 'CABLE MANAGEMENT') return hit('management','catalog family: cable management');
      if (topCategory === 'LIQUID COOLING') return hit('cooling','catalog family: cooling');
      if (/RACK/.test(topCategory)) return hit('rack','catalog family: rack');
      return hit('facility','catalog family: infrastructure');
    }
    if (/^(ACTIVE ELECTRICAL CABLES|COPPER DAC|HIGH-SPEED CABLE ASSEMBLIES|STORAGE AND PCIE-SAS INTERCONNECTS|COAXIAL CABLES|COPPER MODULE CABLE ASSEMBLIES|FIBER CABLE ASSEMBLIES|FIBER CABLES|TWISTED PAIR CABLE ASSEMBLIES|TWISTED PAIR CABLES|CABLE|HARNESS|JUMPER|TRUNK|DATA CENTER CABLING|FIBER OPTIC CABLES|OPTICAL FIBERS|OPTICAL CABLE|POWER CABLE|BREAKOUT HARNESS|MPO CABLE|MPO MTP WIRING|AOC DAC ACC AEC|HIGH FIBER COUNT MPO TRUNKS|MTP MPO HARNESSES|MTP MPO JUMPERS|COPPER DATA CABLES|HIGH DENSITY FIBER ASSEMBLIES|CABLE ASSEMBLIES|RIBBON BREAKOUT & FANOUT KITS|PATCH CORD|광통신|통합배선)$/.test(topCategory)) {
      if (/COPPER|TWISTED PAIR/.test(topCategory)) return hit('copper','catalog family: copper cable');
      if (/TRUNK/.test(topCategory)) return hit('trunk','catalog family: trunk');
      if (topCategory === 'AOC DAC ACC AEC') return hit('aoc','catalog family: active/direct cable');
      if (/ASSEMBL|HARNESS|JUMPER|PATCH CORD|BREAKOUT|FANOUT|INTERCONNECT/.test(topCategory)) return hit('patch','catalog family: cable assembly');
      return hit('fiber','catalog family: cable');
    }
    if (/^(ALL IT DATACOM PRODUCTS|BACKPLANE AND ORTHOGONAL CONNECTORS|ETHERNET USB AND EXTERNAL I-O|FIBER OPTIC CONNECTIVITY|HIGH-SPEED BOARD CONNECTORS|MEMORY AND CARD EDGE CONNECTORS|RF AND COAXIAL CONNECTIVITY|RUGGED CIRCULAR AND D-SUB|TERMINAL BLOCKS AND GENERAL INTERCONNECT|WIRE-TO-BOARD AND FFC-FPC|COPPER PANELS MODULES CASSETTES|BUILDING ENTRANCE SOLUTIONS|FIBER PANELS MODULES CASSETTES|ODF|PROPEL PANELS|ACCESSORIES|BRACKET|CONNECTOR|HOUSING|MODULE|PANEL|OPTICAL CONNECTIVITY AND RACK ENCLOSURES|OPTICAL CONNECTIVITY|CASSETTES & INTERCONNECT PANELS|ENTRANCE FRAMES|FIBER PANELS & SHELVES|OTHER ENCLOSURES|WALL MOUNT ENCLOSURES|FIELD CONNECTORS|LEGACY CONNECTORS|MPO-MT CONNECTORS|SC-LC CONNECTORS|VSFF CONNECTORS|MDC CONNECTORS|MMC CONNECTORS|MT FERRULES|MTP CONNECTORS|FIBER OPTIC CLEANERS)$/.test(topCategory)) {
      if (/ADAPTER|ADAPTOR/.test(t)) return hit('adapter','catalog family: adapter');
      if (/MODULE|CASSETTE/.test(topCategory)) return hit('module','catalog family: module/cassette');
      if (/CONNECTOR|FERRULE|ALL IT DATACOM|BACKPLANE|BOARD|MEMORY|RF AND COAXIAL|RUGGED|TERMINAL|WIRE-TO-BOARD|FIELD|LEGACY|MDC|MMC|MTP|VSFF|SC-LC/.test(topCategory)) return hit('connector','catalog family: connector');
      return hit('panel','catalog family: panel/housing');
    }
    if (/^(OPTICAL TRANSCEIVERS|DATACOM TRANSCEIVERS|OPTICAL PHYS AND DSPS|OPTICAL DSPS|SILICON PHOTONICS PICS|OPTICAL COMPONENTS|CPO LIGHT SOURCES|OPTICAL CHIPS AND LASERS|OPTICAL TIAS AND DRIVERS|CONNECTIVITY DSPS|TRANSCEIVER|SENSORS MATERIALS AND OTHER)$/.test(topCategory)) {
      if (/TRANSCEIVER/.test(topCategory)) return hit('transceiver','catalog family: transceiver');
      return hit('component','catalog family: optics/component');
    }
    if (topCategory === 'SWK™ SERIES') {
      if (/SWK.*CABLE ASSEMBL/.test(fullPath)) return hit('patch','SWK cable assembly');
      if (/SWK.*CONNECTOR/.test(fullPath)) return hit('connector','SWK connector');
      if (/SWK.*PANEL/.test(fullPath)) return hit('panel','SWK panel');
    }
    if (topCategory === 'OPTICAL TRANSCEIVERS AND AOC') {
      if (/AOC/.test(t) && !/TRANSCEIVER/.test(name.toUpperCase())) return hit('aoc','AOC');
      return hit('transceiver','optical transceiver/AOC family');
    }
    if (topCategory === 'PDF SOURCES') return hit('component','legacy optical source catalog');

    // 5) Accelerator — GPU/TPU/NPU and dedicated GPU-server / AI platform catalogs.
    if (/GPU SERVERS?|DATA CENTER GPUS?|PROFESSIONAL GPUS?|AI ACCELERATORS?|AI FACTORY PLATFORM|\\bGPU\\b|\\bTPU\\b|\\bNPU\\b|INSTINCT MI\\d+|GAUDI\\s*3|ASCEND\\s*9|ATLAS 900|DRAGONFLY AI|IRONWOOD|TRILLIUM|\\bBR100\\b|BLACKWELL|HOPPER|VERA RUBIN/.test(t)) {
      return hit('gpu','accelerator/GPU/NPU/TPU');
    }

    // 4) CPU belongs to Network / Server rather than Accelerator.
    if (/(^|[\\/\\s])CPU([\\/\\s]|$)|SERVER CPU|CPU AND SUPERCHIPS|\\bEPYC\\b|\\bXEON\\b|GRACE CPU|AMPEREONE/.test(t)) {
      return hit('cpu','server CPU');
    }

    // 1) User-requested cable overrides. Power Cable is deliberately Cable/Interconnect.
    if (/POWER CABLE|OPTICAL FIBERS?|FIBER CABLES?|FIBRE CABLES?|OPTICAL CABLE|COAXIAL CABLE|TWISTED PAIR CABLE|COPPER DATA CABLE|RIBBON BREAKOUT\\s*&\\s*FANOUT|BREAKOUT HARNESS|MPO MTP WIRING|HIGH DENSITY FIBER ASSEMBL|UCFIBRE|UCFUTURE/.test(t)) {
      if (/TWISTED PAIR|COPPER|CAT.?[568]/.test(t)) return hit('copper','copper cable');
      if (/RIBBON BREAKOUT|BREAKOUT|FANOUT|HARNESS|ASSEMBL/.test(t)) return hit('patch','breakout/harness');
      return hit('fiber','fiber/power cable');
    }

    // 6) Dedicated power distribution before generic connector/panel words.
    if (/POWER CONNECTORS?|POWER DISTRIBUTION|BUSBAR|BUSDUCT|CABLE BUS|SWITCHGEAR|\\bPDU\\b|UNINTERRUPTIBLE|\\bUPS\\b|GENERATOR|TRANSFORMER|BATTERY|EHV|MEDIUM.?VOLTAGE|LOW.?VOLTAGE/.test(t)) {
      if (/\\bUPS\\b|UNINTERRUPTIBLE/.test(t)) return hit('ups','UPS');
      if (/GENERATOR/.test(t)) return hit('generator','generator');
      if (/TRANSFORMER/.test(t)) return hit('transformer','transformer');
      if (/BATTERY/.test(t)) return hit('battery','battery');
      return hit('power','power distribution');
    }

    // 7) Rack / physical infrastructure / cooling.
    if (/LIQUID COOLING|COOLANT|\\bCDU\\b|RDHX|CHILLER|CRAC|CRAH|COOLING|DATA CENTER INFRASTRUCTURE|CABLE MANAGEMENT|CABLE MANAGER|FIBERGUIDE|FIBERRUNNER|PATCHRUNNER|CONTAINMENT|AISLE|IT CABINET|ALL-IN-ONE CABINET|MICRO MODULAR DATA CENTER|CONTAINERIZED DATA CENTER/.test(t)) {
      if (/COOL|CHILLER|CDU|RDHX|CRAC|CRAH/.test(t)) return hit('cooling','cooling infrastructure');
      if (/CABLE MANAGEMENT|CABLE MANAGER|FIBERGUIDE|FIBERRUNNER|PATCHRUNNER/.test(t)) return hit('management','cable management');
      return hit('facility','data-center infrastructure');
    }

    // Preserve reliable fine-grained BOM kinds where they already exist.
    if (['fiber','trunk','patch','aoc','dac','aec','copper',
         'connector','adapter','module','panel',
         'transceiver','component',
         'switch','nic','server',
         'ups','generator','transformer','battery',
         'rack','management'].includes(raw)) {
      return hit(raw,'existing BOM kind');
    }

    // 2) Connector / panel / adapter families.
    if (/BUILDING ENTRANCE SOLUTIONS|FIBER PANELS?|FIBRE PANELS?|MODULES?[ _/&-]*CASSETTES?|PATCH PANEL|ADAPTER|ADAPTOR|CONNECTOR|FERRULE|ODF|HOUSING|ENCLOSURE|ENTRANCE FRAME|WALL MOUNT|SPLICE TRAY|CASSETTE|TERMINATION BOX|PATCHING FRAME|DISTRIBUTION FRAME/.test(t)) {
      if (/ADAPTER|ADAPTOR/.test(t)) return hit('adapter','adapter');
      if (/MODULE|CASSETTE/.test(t)) return hit('module','module/cassette');
      if (/CONNECTOR|FERRULE/.test(t)) return hit('connector','connector/ferrule');
      return hit('panel','panel/housing/entrance');
    }

    // 3) Optics / component.
    if (/TRANSCEIVER|OPTICAL MODULE|\\bLASER\\b|LASER CHIP|\\bDSP\\b|SILICON PHOTON|\\bPIC\\b|\\bTIA\\b|PHOTODIODE|PHOTONIC|CPO LIGHT SOURCE|OPTICAL PHY|CW-DFB|DFB LASER|CLEANER/.test(t)) {
      if (/TRANSCEIVER|OPTICAL MODULE/.test(t)) return hit('transceiver','transceiver');
      return hit('component','optical/electronic component');
    }

    // 4) Network / Server.
    if (/NETWORKING|DATA CENTER SWITCH|\\bSWITCH\\b|\\bNIC\\b|SUPERNIC|NETWORK ADAPTER|CONNECTX|BLUEFIELD|QUANTUM|SPECTRUM|LEAF|SPINE|RACK SERVERS?|SERVER PLATFORM|COMPUTE TRAY|PROLIANT|POWEREDGE|SUPERMICRO SYS-/.test(t)) {
      if (/\\bNIC\\b|SUPERNIC|NETWORK ADAPTER|CONNECTX|BLUEFIELD/.test(t)) return hit('nic','NIC/network adapter');
      if (/SWITCH|QUANTUM|SPECTRUM|LEAF|SPINE/.test(t)) return hit('switch','network switch');
      return hit('server','server');
    }

    // 7) Generic racks / accessories / hardware.
    if (/\\bRACKS?\\b|RACK INFRASTRUCTURE|SERVER RACK|CABINET|MOUNTING HARDWARE|MOUNTING PLATE|BRACKET|ACCESSORIES|ACCESSORY/.test(t)) {
      return hit('rack','rack/infrastructure accessory');
    }

    // Vendor/category fallback for broad catalogs that otherwise create a large Other bucket.
    if (/COMMSCOPE/.test(t)) {
      if (/CABLE/.test(t)) return hit('fiber','CommScope cable family');
      return hit('panel','CommScope connectivity/entrance family');
    }
    if (/AMPHENOL/.test(t)) {
      if (/CABLE|AEC|AOC|DAC|ASSEMBL/.test(t)) return hit('patch','Amphenol interconnect');
      if (/OPTICAL|TRANSCEIVER|DSP|SILICON/.test(t)) return hit('component','Amphenol optical component');
      if (/POWER/.test(t)) return hit('power','Amphenol power');
      return hit('connector','Amphenol IT/datacom interconnect');
    }
    if (/CORNING|SENKO|US.?CONEC/.test(t)) return hit('connector','optical connectivity vendor');
    if (/SUMITOMO ELECTRIC/.test(t)) {
      if (/CABLE|JUMPER|SWK/.test(t)) return hit('patch','Sumitomo cable/interconnect');
      return hit('panel','Sumitomo optical connectivity');
    }
    if (/FURUKAWA ELECTRIC/.test(t)) {
      if (/NETWORK/.test(t)) return hit('switch','Furukawa network');
      if (/CABLE/.test(t)) return hit('fiber','Furukawa cable');
      return hit('component','Furukawa optical component');
    }
    if (/FUJIKURA/.test(t)) {
      if (/FIBER|FIBRE|CABLE/.test(t)) return hit('fiber','Fujikura fiber/cable');
      return hit('panel','Fujikura connectivity');
    }
    if (/ZTT/.test(t)) {
      if (/DATA CENTER INFRASTRUCTURE/.test(t)) return hit('facility','ZTT infrastructure');
      return hit('patch','ZTT optical interconnect');
    }
    if (/LS CNS|LS CABLE/.test(t)) return hit('fiber','LS cabling');
    if (/GAON CABLE|TAIHAN/.test(t)) return hit('power','Korean data-center power');
    if (/MOTIVAIR|COOLIT|VERTIV|RITTAL/.test(t) && /COOL|RACK|INFRA/.test(t)) return hit('facility','rack/cooling infrastructure');
    if (/NVIDIA|AMD|INTEL|QUALCOMM|GOOGLE/.test(t)) {
      if (/CPU|PROCESSOR/.test(t)) return hit('cpu','compute CPU fallback');
      if (/GPU|ACCELERATOR|TPU|NPU|AI /.test(t)) return hit('gpu','compute accelerator fallback');
      if (/NETWORK|CONNECTIVITY/.test(t)) return hit('nic','compute/network fallback');
    }
    if (/DRAKA|PRYSMIAN|NEXANS|HYC|IH OPTICS|SHIJIA|ZSINE|NADDOD|BELDEN|PANDUIT/.test(t)) {
      return hit('patch','cable/interconnect vendor');
    }
    if (/SCHNEIDER|EATON|LS ELECTRIC|MPOWERSYS|XEONICS|GREEN POWER|CUMMINS|CATERPILLAR|HITACHI ENERGY/.test(t)) {
      return hit('power','power vendor');
    }
    if (/CISCO|ARISTA|JUNIPER|BROADCOM/.test(t)) return hit('switch','network vendor');
    if (/LUMENTUM|COHERENT|MARVELL|MACOM|AOI|CREDO/.test(t)) return hit('component','optics/component vendor');

    return hit('other','unclassified');
  }

  function dbKindOf(product) {
    return dbTaxonomy(product).kind;
  }

  function dbProductsForKinds(products, kinds) {
    const set = new Set(kinds);
    return products.filter(p => set.has(dbKindOf(p)));
  }

  function dbCountMap(items, keyFn) {
    const map = new Map();
    items.forEach(item => {
      const key = keyFn(item);
      if (key) map.set(key, (map.get(key) || 0) + 1);
    });
    return map;
  }

  function dbTopCategoryForCompany(items, company) {
    const cats = dbCountMap(items.filter(p => dbPathParts(p)[0] === company), p => dbPathParts(p)[1] || p.category || 'Other');
    return [...cats.entries()].sort((a,b) => b[1] - a[1] || a[0].localeCompare(b[0]))[0]?.[0] || '';
  }

  function renderDbDetail(kinds, label) {
    const host = $('catalogDbDetail');
    const source = state.dbMapData?.products || [];
    if (!host || !source.length) return;
    const items = dbProductsForKinds(source, kinds);
    const companies = dbCountMap(items, p => dbPathParts(p)[0] || p.vendor || 'Unknown');
    const categories = dbCountMap(items, p => dbPathParts(p)[1] || p.category || 'Other');
    const paths = new Set(items.map(p => p.path).filter(Boolean));
    const companyRows = [...companies.entries()].sort((a,b) => b[1] - a[1] || a[0].localeCompare(b[0])).slice(0,18);
    const categoryRows = [...categories.entries()].sort((a,b) => b[1] - a[1] || a[0].localeCompare(b[0])).slice(0,16);
    host.hidden = false;
    host.innerHTML =
      '<div class="catalog-db-detail-head"><div><b>' + esc(label) + '</b><span>제품 ' + items.length.toLocaleString('ko-KR') + '개 · catalog ' + paths.size.toLocaleString('ko-KR') + '개 · 업체 ' + companies.size.toLocaleString('ko-KR') + '개</span></div><button type="button" id="catalogDbDetailClose" aria-label="DB 상세 닫기">닫기</button></div>' +
      '<div class="catalog-db-detail-grid"><div><h4>업체별 자료</h4><div class="catalog-db-vendor-list">' +
      companyRows.map(([company,count]) => {
        const category = dbTopCategoryForCompany(items, company);
        return '<button type="button" data-db-company="' + esc(company) + '" data-db-category="' + esc(category) + '"><span>' + esc(cleanLabel(company)) + '</span><b>' + count.toLocaleString('ko-KR') + '</b></button>';
      }).join('') + '</div></div><div><h4>세부 부품군</h4><div class="catalog-db-category-list">' +
      categoryRows.map(([category,count]) => '<span><i>' + esc(category) + '</i><b>' + count.toLocaleString('ko-KR') + '</b></span>').join('') +
      '</div></div></div>';
    $('catalogDbDetailClose').onclick = () => { host.hidden = true; };
    host.querySelectorAll('[data-db-company]').forEach(btn => {
      btn.onclick = async () => {
        const company = btn.dataset.dbCompany;
        const category = btn.dataset.dbCategory;
        await selectCompany(company);
        if (category) await selectCategory(category);
        $('catalogCategories')?.scrollIntoView({behavior:'smooth', block:'start'});
      };
    });
  }

  function renderDbMap(data) {
    const host = $('catalogDbMap');
    const status = $('catalogDbMapStatus');
    if (!host) return;
    const products = data?.products || [];
    state.dbMapData = data || null;
    if (!products.length) {
      host.innerHTML = '<p class="catalog-empty">집계 가능한 구조화 제품이 없습니다.</p>';
      if (status) status.textContent = '0개';
      return;
    }
    const companyCount = new Set(products.map(p => dbPathParts(p)[0] || p.vendor).filter(Boolean)).size;
    const catalogCount = new Set(products.map(p => p.path).filter(Boolean)).size;
    const groups = dbGroups.map(group => {
      const items = dbProductsForKinds(products, group.kinds);
      const pathCount = new Set(items.map(p => p.path).filter(Boolean)).size;
      const kindCounts = dbCountMap(items, p => dbKindOf(p));
      return {...group, count:items.length, pathCount, kindCounts};
    });
    const otherCount = dbProductsForKinds(products, ['other']).length;
    const otherRate = products.length ? (otherCount / products.length * 100) : 0;
    window.__dcDbTaxonomySummary = {
      total: products.length,
      other: otherCount,
      otherRate: Number(otherRate.toFixed(2)),
      groups: Object.fromEntries(groups.map(g => [g.id, g.count]))
    };
    host.innerHTML =
      '<div class="catalog-db-hub"><span>전체 제품 DB</span><b>' + products.length.toLocaleString('ko-KR') + '</b><small>' + companyCount + '개 업체 · ' + catalogCount + '개 catalog · Other ' + otherRate.toFixed(1) + '%</small></div>' +
      '<div class="catalog-db-branches">' +
      groups.map(group =>
        '<article class="catalog-db-node"><div class="catalog-db-node-head"><div><span>' + esc(group.label) + '</span><b>' + group.count.toLocaleString('ko-KR') + '개</b></div><button type="button" data-db-group="' + esc(group.id) + '">상세</button></div>' +
        '<div class="catalog-db-kind-list">' +
        group.kinds.filter(kind => (group.kindCounts.get(kind) || 0) > 0).map(kind =>
          '<button type="button" data-db-kind="' + esc(kind) + '"><span>' + esc(dbKindLabels[kind] || kind) + '</span><b>' + (group.kindCounts.get(kind) || 0).toLocaleString('ko-KR') + '</b></button>'
        ).join('') + '</div><small>catalog ' + group.pathCount.toLocaleString('ko-KR') + '개</small></article>'
      ).join('') + '</div>';
    if (status) status.textContent = '제품 ' + products.length.toLocaleString('ko-KR') + '개';
    host.querySelectorAll('[data-db-group]').forEach(btn => {
      const group = dbGroups.find(x => x.id === btn.dataset.dbGroup);
      btn.onclick = () => group && renderDbDetail(group.kinds, group.label);
    });
    host.querySelectorAll('[data-db-kind]').forEach(btn => {
      const kind = btn.dataset.dbKind;
      btn.onclick = () => renderDbDetail([kind], dbKindLabels[kind] || kind);
    });
  }

  async function loadDbMap(force) {
    const host = $('catalogDbMap');
    const status = $('catalogDbMapStatus');
    if (!host) return;
    if (!window.DCBomCatalog?.load) {
      host.innerHTML = '<p class="catalog-empty">제품 DB 집계 모듈을 찾지 못했습니다.</p>';
      if (status) status.textContent = '집계 불가';
      return;
    }
    host.setAttribute('aria-busy','true');
    if (status) status.textContent = 'DB 집계 중…';
    try {
      const data = await window.DCBomCatalog.load(!!force);
      renderDbMap(data);
    } catch (error) {
      host.innerHTML = '<p class="catalog-empty">제품 DB 집계 실패: ' + esc(error.message) + '</p>';
      if (status) status.textContent = '집계 실패';
    } finally {
      host.removeAttribute('aria-busy');
    }
  }

  function tourTerms(){
    if(!tourFocus)return [];
    const base=tourFocus.toLowerCase().split(/[^a-z0-9가-힣.+/-]+/).filter(x=>x.length>=2);
    const aliases={
      odf:['odf','fiber panel','panel','module'],
      fiber:['fiber','optical','trunk','patch'],
      trunk:['trunk','cable assembl'],
      patch:['patch','cable assembl'],
      transceiver:['transceiver','networking','optical'],
      switch:['networking','switch'],
      pdu:['pdu','power'],
      busway:['busway','power'],
      cable:['cable'],
      copper:['copper','twisted pair'],
      cat6:['cat6','twisted pair'],
      rack:['rack','panel'],
      cabletray:['cable tray','wire basket','kwikrail','cablofil'],
      raceway:['fiber raceway','fiberrunner','fiberguide'],
      ups:['ups','power'],
      transformer:['transformer','power'],
      cpu:['cpu','processor','xeon','epyc','grace'],
      gpu:['gpu','accelerator','nvidia','amd instinct','gaudi','ascend'],
      splicer:['splicer','fusion','90s','90r','s179','s124','s185']
    };
    const out=[...base];
    base.forEach(t=>{(aliases[t]||[]).forEach(a=>out.push(a))});
    return [...new Set(out)];
  }

  function buildTourSuggestions(){
    const host=$('catalogTourFocus'),textEl=$('catalogTourFocusText'),list=$('catalogTourSuggestions');
    if(!host||!list||!tourFocus||!catalogIndex?.paths?.length)return [];
    const terms=tourTerms(), scored=new Map();
    for(const path of catalogIndex.paths){
      const parts=catalogPathParts(path), company=parts[0], category=parts[1];
      if(!company||!category)continue;
      const hay=(company+' '+category+' '+parts.slice(2,-1).join(' ')).toLowerCase();
      let score=0;
      for(const t of terms)if(hay.includes(t))score+=t.length>=5?3:1;
      if(!score)continue;
      const key=company+'|'+category, prev=scored.get(key);
      if(!prev||score>prev.score)scored.set(key,{company,category,score});
    }
    const suggestions=[...scored.values()].sort((a,b)=>b.score-a.score||a.company.localeCompare(b.company)).slice(0,8);
    host.hidden=false;textEl.textContent=' · '+tourFocus;
    list.innerHTML=suggestions.length?suggestions.map((s,i)=>'<button type="button" data-tour-company="'+esc(s.company)+'" data-tour-category="'+esc(s.category)+'"'+(i===0?' class="primary"':'')+'>'+esc(cleanLabel(s.company))+' · '+esc(s.category)+'</button>').join(''):'관련 제품군을 자동 식별하지 못했습니다. 업체/부품군을 직접 선택하세요.';
    list.querySelectorAll('[data-tour-company]').forEach(btn=>btn.onclick=async()=>{
      await selectCompany(btn.dataset.tourCompany);
      await selectCategory(btn.dataset.tourCategory);
      $('catalogProducts')?.scrollIntoView({behavior:'smooth',block:'start'});
    });
    return suggestions;
  }

  async function applyTourFocusOnce(){
    if(tourSuggestionsApplied||!tourFocus)return;
    tourSuggestionsApplied=true;
    const suggestions=buildTourSuggestions();
    if(tourParams.get('openCatalog')==='1'&&suggestions[0]){
      await selectCompany(suggestions[0].company);
      await selectCategory(suggestions[0].category);
      requestAnimationFrame(()=>$('catalogProducts')?.scrollIntoView({behavior:'smooth',block:'start'}));
    }
  }

  async function loadCompanies(force) {
    notice('업체 목록을 불러오는 중입니다.');
    setStep(1);
    state.company = '';
    state.category = '';
    state.products = [];
    state.manifest = null;
    resetFamilyFilters();
    $('catalogCategories').innerHTML = '<p class="catalog-empty">업체를 선택하세요.</p>';
    $('catalogProducts').innerHTML = '<p class="catalog-empty">제품 자료를 선택하세요.</p>';
    $('catalogSearch').value = '';
    $('catalogSearch').disabled = true;
    $('catalogContext').textContent = '부품군을 선택하면 등록된 PDF·공식 URL 제품이 표시됩니다.';

    try {
      await loadCatalogIndex(force);
      const companies = [...new Set((catalogIndex.paths || []).map(path => catalogPathParts(path)[0]).filter(Boolean))]
        .sort((a,b) => a.localeCompare(b))
        .map(name => ({name}));
      $('catalogCompanyCount').textContent = companies.length + '개';
      $('catalogCompanies').innerHTML = companies.length ? companies.map(x => button(x.name, 'company', false)).join('') : '<p class="catalog-empty">등록된 업체가 없습니다.</p>';
      $('catalogCategoryCount').textContent = '—';
      notice(companies.length + '개 업체가 등록되어 있습니다. 업체를 선택하세요.', 'ready');
      bindCompanyButtons();
      loadDbMap(false);
      await applyTourFocusOnce();
    } catch (error) {
      notice(error.message, 'error');
      $('catalogCompanies').innerHTML = '<a class="catalog-fallback" href="https://github.com/kimheeseo/LSCNS/tree/main/DCI/DataCenter/product_catalog" target="_blank" rel="noopener">GitHub product_catalog 열기 ↗</a>';
    }
  }

  function bindCompanyButtons() {
    document.querySelectorAll('#catalogCompanies [data-company]').forEach(el => {
      el.onclick = () => selectCompany(el.dataset.company);
    });
  }

  async function selectCompany(company) {
    state.company = company;
    state.category = '';
    state.products = [];
    state.manifest = null;
    resetFamilyFilters();
    setStep(2);
    document.querySelectorAll('#catalogCompanies [data-company]').forEach(el => {
      const selected = el.dataset.company === company;
      el.classList.toggle('selected', selected);
      el.setAttribute('aria-pressed', String(selected));
    });
    $('catalogCategories').innerHTML = '<p class="catalog-empty">부품군을 불러오는 중입니다.</p>';
    $('catalogProducts').innerHTML = '<p class="catalog-empty">부품군을 선택하세요.</p>';
    $('catalogSearch').value = '';
    $('catalogSearch').disabled = true;
    $('catalogContext').textContent = company + '의 부품군을 선택하세요.';
    notice(company + ' 카탈로그를 불러오는 중입니다.');

    try {
      await loadCatalogIndex(false);
      const categories = [...new Set(pathsFor(company, '').map(path => catalogPathParts(path)[1]).filter(Boolean))]
        .sort((a,b) => a.localeCompare(b))
        .map(name => ({name}));
      $('catalogCategoryCount').textContent = categories.length + '개';
      $('catalogCategories').innerHTML = categories.length ? categories.map(x => button(x.name, 'category', false)).join('') : '<p class="catalog-empty">등록된 부품군이 없습니다.</p>';
      notice(company + ' · ' + categories.length + '개 부품군', 'ready');
      document.querySelectorAll('#catalogCategories [data-category]').forEach(el => {
        el.onclick = () => selectCategory(el.dataset.category);
      });
    } catch (error) {
      notice(error.message, 'error');
      $('catalogCategories').innerHTML = '<p class="catalog-empty">부품군을 불러오지 못했습니다.</p>';
    }
  }

  function modelFromFile(name) {
    return String(name || '').replace(/\.pdf$/i, '').replace(/_NAFTA_AEN$/i, '').replace(/_AEN$/i, '');
  }

  function specObject(manifest, meta) {
    const merged = {};
    const add = source => {
      if (!source) return;
      if (Array.isArray(source)) {
        source.forEach((value, index) => merged['Spec ' + (index + 1)] = value);
      } else if (typeof source === 'object') {
        Object.entries(source).forEach(([key, value]) => merged[key] = value);
      }
    };
    add(manifest && manifest.defaultSpecs);
    add(meta && meta.specs);
    return merged;
  }

  function productMeta(manifest, file, model) {
    const products = manifest && manifest.products || {};
    return products[model] || products[file.name] || {};
  }

  function shortCommScopeTitle(meta, model, specs) {
    const full = meta.name || model;
    const part = String(specs['Part Number'] || '').trim() || String(full).split('·')[0].trim() || model;
    const type = String(specs['Product Type'] || '').trim();
    return [part, type].filter(Boolean).join(' · ') || full;
  }

  function entryData(entry) {
    const file = entry.file;
    const manifest = entry.manifest;
    const group = entry.group || state.category;
    const rawModel = modelFromFile(file.name);
    const meta = productMeta(manifest, file, rawModel);
    const model = meta.characteristicsOnly ? 'Characteristics' : (meta.displayModel || rawModel);
    const specs = specObject(manifest, meta);
    const fullTitle = meta.name || model;
    const title = /COMMSCOPE/i.test(String(state.company||manifest?.company||'')) ? shortCommScopeTitle(meta, model, specs) : fullTitle;
    const description = meta.description || (manifest && manifest.productDescription) || '';
    const officialUrl = meta.officialUrl || (manifest && manifest.officialUrl) || '';
    const checked = meta.checked || (manifest && manifest.checked) || '';
    const localPdfUrl = meta.datasheetPath ? staticUrl(meta.datasheetPath) : '';
    const datasheetUrl = meta.datasheetUrl || localPdfUrl;
    const officialPdfUrl = meta.officialPdfUrl || '';
    const solutionOverviewUrl = meta.solutionOverviewUrl || '';
    const referenceOnly = !!(meta.referenceOnly || meta.vendorReferenceOnly);
    const searchText = [title, fullTitle, model, file.name, state.company, state.category, group, description]
      .concat(Object.entries(specs).flat()).join(' ').toLowerCase();
    return {file, manifest, group, model, specs, title, description, officialUrl, datasheetUrl, localPdfUrl, officialPdfUrl, solutionOverviewUrl, referenceOnly, checked, searchText};
  }

  function comparisonKeys(entries) {
    const keys = [], seen = new Set();
    entries.forEach(entry => {
      const declared = entry.manifest && Array.isArray(entry.manifest.comparisonFields) ? entry.manifest.comparisonFields : [];
      declared.forEach(key => { if (!seen.has(key)) { seen.add(key); keys.push(key); } });
    });
    if (keys.length) return keys;
    entries.forEach(entry => Object.keys(entryData(entry).specs).forEach(key => {
      if (!seen.has(key)) { seen.add(key); keys.push(key); }
    }));
    return keys;
  }

  function groupSchemas(entries) {
    const groups = new Map(), schemas = new Map();
    entries.forEach(entry => {
      if (!groups.has(entry.group)) groups.set(entry.group, []);
      groups.get(entry.group).push(entry);
    });
    groups.forEach((items, group) => schemas.set(group, comparisonKeys(items)));
    return schemas;
  }

  function card(entry, specKeys) {
    const d = entryData(entry);
    const keys = specKeys && specKeys.length ? specKeys : Object.keys(d.specs);
    const rows = keys.map(key => [key, d.specs[key] ?? '—']);
    const virtual = !!d.file.virtual;
    const primaryUrl = d.file.html_url || d.officialUrl;
    const primaryLabel = virtual ? (d.referenceOnly ? '공식 기술자료 보기 ↗' : '공식 제품 페이지 보기 ↗') : '제품 PDF 보기 ↗';
    const secondary = (!virtual && d.officialUrl && d.officialUrl !== d.file.html_url)
      ? '<a href="' + esc(d.officialUrl) + '" target="_blank" rel="noopener">공식 제품 페이지 ↗</a>' : '';
    const datasheet = d.datasheetUrl && d.datasheetUrl !== primaryUrl ? '<a href="' + esc(d.datasheetUrl) + '" target="_blank" rel="noopener">데이터시트 PDF ↗</a>' : '';
    const localPdf = d.localPdfUrl && d.localPdfUrl !== d.datasheetUrl ? '<a href="' + esc(d.localPdfUrl) + '" target="_blank" rel="noopener">업로드 PDF ↗</a>' : '';
    const overview = d.solutionOverviewUrl && d.solutionOverviewUrl !== primaryUrl ? '<a href="' + esc(d.solutionOverviewUrl) + '" target="_blank" rel="noopener">플랫폼 개요 ↗</a>' : '';
    const officialPdf = d.officialPdfUrl && d.officialPdfUrl !== primaryUrl ? '<a href="' + esc(d.officialPdfUrl) + '" target="_blank" rel="noopener">제조사 원문 PDF ↗</a>' : '';
    return '<article class="catalog-card" data-family="' + esc(d.group) + '" data-search="' + esc(d.searchText) + '">' +
      '<div class="catalog-card-top"><div><span class="catalog-vendor">' + esc(cleanLabel(state.company)) + '</span><span class="catalog-family">' + esc(d.group) + '</span><h4>' + esc(d.title) + '</h4><code>' + esc(d.model) + '</code></div><span class="catalog-file-size">' + (virtual ? 'URL' : esc(humanBytes(d.file.size))) + '</span></div>' +
      (d.description ? '<p class="catalog-description">' + esc(d.description) + '</p>' : '') +
      (rows.length ? '<dl class="catalog-specs">' + rows.map(([key,value]) => '<div><dt>' + esc(key) + '</dt><dd>' + esc(value) + '</dd></div>').join('') + '</dl>' :
        '<div class="catalog-no-spec">공통 비교 스펙 미등록 · catalog.json에 comparisonFields를 추가해야 합니다.</div>') +
      '<div class="catalog-card-actions">' + (primaryUrl ? '<a href="' + esc(primaryUrl) + '" target="_blank" rel="noopener">' + primaryLabel + '</a>' : '') + secondary + datasheet + localPdf + overview + officialPdf + '</div>' +
      (d.checked ? '<small class="catalog-checked">' + (d.referenceOnly ? '자료 등록일 ' : '사양 확인일 ') + esc(d.checked) + '</small>' : '') +
    '</article>';
  }
  function table(entries) {
    const groups = new Map();
    entries.forEach(entry => {
      if (!groups.has(entry.group)) groups.set(entry.group, []);
      groups.get(entry.group).push(entry);
    });
    return [...groups.entries()].map(([group, items]) => {
      const keys = comparisonKeys(items);
      const data = items.map(entryData);
      const headers = ['모델', ...keys, '자료', '공식 페이지', '데이터시트 / 기술자료'];
      const rows = data.map(d => '<tr class="catalog-table-row" data-family="' + esc(d.group) + '" data-search="' + esc(d.searchText) + '">' +
        '<td><b>' + esc(d.title) + '</b><small>' + esc(d.model) + '</small></td>' +
        keys.map(key => '<td>' + esc(d.specs[key] ?? '—') + '</td>').join('') +
        '<td>' + ((d.file.html_url || d.officialUrl) ? '<a href="' + esc(d.file.html_url || d.officialUrl) + '" target="_blank" rel="noopener">' + (d.file.virtual ? 'URL ↗' : 'PDF ↗') + '</a>' : '—') + '</td>' +
        '<td>' + (!d.file.virtual && d.officialUrl && d.officialUrl !== d.file.html_url ? '<a href="' + esc(d.officialUrl) + '" target="_blank" rel="noopener">공식 ↗</a>' : (d.file.virtual ? 'URL 제품' : '—')) + '</td>' +
        '<td>' + (d.datasheetUrl ? '<a href="' + esc(d.datasheetUrl) + '" target="_blank" rel="noopener">PDF ↗</a>' : '—') + (d.localPdfUrl && d.localPdfUrl !== d.datasheetUrl ? ' <a href="' + esc(d.localPdfUrl) + '" target="_blank" rel="noopener">보관 PDF ↗</a>' : '') + (d.solutionOverviewUrl ? ' <a href="' + esc(d.solutionOverviewUrl) + '" target="_blank" rel="noopener">개요 ↗</a>' : '') + (d.officialPdfUrl && d.officialPdfUrl !== d.datasheetUrl ? ' <a href="' + esc(d.officialPdfUrl) + '" target="_blank" rel="noopener">원문 PDF ↗</a>' : '') + '</td></tr>').join('');
      return '<section class="catalog-table-group" data-table-family="' + esc(group) + '">' +
        '<div class="catalog-table-group-head"><span class="catalog-family table-family">' + esc(group) + '</span> <b>' + items.length + '개 · 동일 스펙 기준 비교</b></div>' +
        (keys.length ? '<div class="catalog-sheet-wrap"><table class="catalog-sheet"><thead><tr>' + headers.map(h => '<th>' + esc(h) + '</th>').join('') + '</tr></thead><tbody>' + rows + '</tbody></table></div>' :
        '<div class="catalog-no-spec">이 제품군은 공통 비교 스펙이 아직 등록되지 않았습니다.</div>') +
        '</section>';
    }).join('');
  }

  function renderProducts(entries) {
    const host = $('catalogProducts');
    if (!host) return;
    host.classList.toggle('table-mode', state.viewMode === 'table');
    if (!entries.length) {
      host.innerHTML = '<p class="catalog-empty">이 부품군과 하위 폴더에 등록된 PDF/URL 제품이 없습니다.</p>';
      return;
    }
    if (state.viewMode === 'table') {
      host.innerHTML = table(entries) + '<p id="catalogSearchEmpty" class="catalog-empty" hidden>검색 조건과 일치하는 제품이 없습니다.</p>';
    } else {
      const schemas = groupSchemas(entries);
      host.innerHTML = entries.map(entry => card(entry, schemas.get(entry.group) || [])).join('') +
        '<p id="catalogSearchEmpty" class="catalog-empty" hidden>검색 조건과 일치하는 제품이 없습니다.</p>';
    }
    applySearch();
  }

  function renderFamilyFilters(entries) {
    const host = $('catalogFamilyFilters');
    if (!host) return;
    const counts = new Map();
    entries.forEach(entry => counts.set(entry.group, (counts.get(entry.group) || 0) + 1));
    const families = [...counts.keys()].sort((a,b) => a.localeCompare(b));
    state.selectedFamilies.clear();
    if (families.length <= 1) {
      host.hidden = true;
      host.innerHTML = '';
      return;
    }
    host.hidden = false;
    host.innerHTML =
      '<div class="catalog-filter-title"><b>제품군 필터</b><span>공통 분류별로 제품을 좁혀볼 수 있습니다.</span></div>' +
      '<div class="catalog-filter-checks">' +
        '<label class="catalog-filter-chip all active"><input type="checkbox" data-family-filter="__all__" checked><span>전체</span><em>' + entries.length + '</em></label>' +
        families.map(f => '<label class="catalog-filter-chip"><input type="checkbox" data-family-filter="' + esc(f) + '"><span>' + esc(f) + '</span><em>' + counts.get(f) + '</em></label>').join('') +
      '</div>';

    host.querySelectorAll('[data-family-filter]').forEach(input => {
      input.onchange = () => {
        const all = host.querySelector('[data-family-filter="__all__"]');
        const individual = [...host.querySelectorAll('[data-family-filter]:not([data-family-filter="__all__"])')];

        if (input.dataset.familyFilter === '__all__') {
          if (input.checked) {
            state.selectedFamilies.clear();
            individual.forEach(x => x.checked = false);
          } else if (!individual.some(x => x.checked)) {
            input.checked = true;
          }
        } else if (input.checked) {
          all.checked = false;
          state.selectedFamilies.clear();
          individual.filter(x => x.checked).forEach(x => state.selectedFamilies.add(x.dataset.familyFilter));
        } else {
          state.selectedFamilies.delete(input.dataset.familyFilter);
          if (!individual.some(x => x.checked)) {
            all.checked = true;
            state.selectedFamilies.clear();
          }
        }

        host.querySelectorAll('.catalog-filter-chip').forEach(label => {
          const box = label.querySelector('input');
          label.classList.toggle('active', !!box.checked);
        });
        applySearch();
      };
    });
  }

  function resetFamilyFilters() {
    state.selectedFamilies.clear();
    const host = $('catalogFamilyFilters');
    if (host) {
      host.hidden = true;
      host.innerHTML = '';
    }
  }

  function applySearch() {
    const q = ($('catalogSearch').value || '').trim().toLowerCase();
    const cards = Array.from(document.querySelectorAll('#catalogProducts .catalog-card, #catalogProducts .catalog-table-row'));
    let shown = 0;
    cards.forEach(el => {
      const searchMatch = !q || (el.dataset.search || '').includes(q);
      const familyMatch = state.selectedFamilies.size === 0 || state.selectedFamilies.has(el.dataset.family || '');
      const match = searchMatch && familyMatch;
      el.hidden = !match;
      if (match) shown++;
    });
    document.querySelectorAll('#catalogProducts .catalog-table-group').forEach(section => {
      const rows = [...section.querySelectorAll('.catalog-table-row')];
      section.hidden = rows.length > 0 && rows.every(row => row.hidden);
    });
    const empty = $('catalogSearchEmpty');
    if (empty) empty.hidden = shown !== 0;
    const result = $('catalogFilterResult');
    if (result) result.textContent = shown + '개 표시';
  }

  async function selectCategory(category) {
    state.category = category;
    setStep(3);
    document.querySelectorAll('#catalogCategories [data-category]').forEach(el => {
      const selected = el.dataset.category === category;
      el.classList.toggle('selected', selected);
      el.setAttribute('aria-pressed', String(selected));
    });
    $('catalogProducts').innerHTML = '<p class="catalog-empty">제품 자료와 스펙을 불러오는 중입니다.</p>';
    $('catalogSearch').value = '';
    $('catalogSearch').disabled = true;
    resetFamilyFilters();
    $('catalogContext').textContent = state.company + ' › ' + category;
    notice(state.company + ' · ' + category + ' 제품 자료를 불러오는 중입니다.');

    try {
      const loaded = await walkProducts(state.company, category, false);
      const entries = loaded.entries;
      entries.sort((a,b) => a.group.localeCompare(b.group) || a.file.name.localeCompare(b.file.name));
      state.products = entries;
      state.manifest = null;
      const families = [...new Set(entries.map(x => x.group))];
      const specCount = entries.filter(x => x.manifest).length;
      const urlOnlyCount = entries.filter(x => x.file && x.file.virtual).length;
      const pdfOnlyCount = entries.length - urlOnlyCount;
      $('catalogContext').innerHTML = esc(cleanLabel(state.company)) + ' › ' + esc(category) + ' · ' + entries.length + '개 제품 · ' + families.length + '개 제품군 · <span id="catalogFilterResult">' + entries.length + '개 표시</span>';
      renderProducts(entries);
      $('catalogSearch').disabled = !entries.length;
      renderFamilyFilters(entries);
      applySearch();
      notice(state.company + ' · ' + category + ' · 제품 ' + entries.length + '개 · 제품군 ' + families.length + '개 · 정적 카탈로그 조회' + (loaded.errors.length ? ' · 일부 파일 실패 ' + loaded.errors.length + '개' : ''), loaded.errors.length ? 'review' : 'ready');
    } catch (error) {
      notice(error.message, 'error');
      $('catalogProducts').innerHTML = '<p class="catalog-empty">제품 자료를 불러오지 못했습니다.</p>';
    }
  }

  function mount() {
    const view = $('view-catalog');
    if (!view || view.dataset.catalogMounted === 'true') return !!view;
    view.dataset.catalogMounted = 'true';
    view.innerHTML = shell();
    $('catalogRefresh').onclick = () => {
      catalogIndex = null;
      document.dispatchEvent(new CustomEvent("dc:catalog-refresh"));
      loadCompanies(true);
      loadDbMap(true);
    };
    $('catalogKoreaBtn').onclick = async () => {
      const panel = $('catalogKoreaPanel');
      const button = $('catalogKoreaBtn');
      const open = panel.hidden;
      panel.hidden = !open;
      button.setAttribute('aria-expanded', String(open));
      button.classList.toggle('active', open);
      if (open) await loadKoreaSuppliers();
    };
    $('catalogSearch').oninput = applySearch;
    document.querySelectorAll('[data-catalog-view]').forEach(button => {
      button.onclick = () => {
        state.viewMode = button.dataset.catalogView;
        document.querySelectorAll('[data-catalog-view]').forEach(b => {
          const active = b.dataset.catalogView === state.viewMode;
          b.classList.toggle('active', active);
          b.setAttribute('aria-pressed', String(active));
        });
        renderProducts(state.products);
      };
    });
    loadCompanies(false);
    return true;
  }

  function start() {
    if (mount()) return;
    let tries = 0;
    const timer = setInterval(() => {
      tries++;
      if (mount() || tries > 20) clearInterval(timer);
    }, 150);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start, {once:true});
  else start();
})();