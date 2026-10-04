(() => {
  'use strict';

  const REPO = 'kimheeseo/LSCNS';
  const REF = 'main';
  const BASE = 'DCI/DataCenter/product_catalog';
  const API = 'https://api.github.com/repos/' + REPO + '/contents/';
  const cache = new Map();
  const state = { company: '', category: '', products: [], manifest: null, selectedFamilies: new Set(), viewMode: 'cards' };
  const $ = id => document.getElementById(id);
  const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const pathUrl = path => path.split('/').map(encodeURIComponent).join('/');
  const humanBytes = value => {
    const n = Number(value) || 0;
    if (n < 1024) return n + ' B';
    if (n < 1024 * 1024) return (n / 1024).toFixed(0) + ' KB';
    return (n / 1024 / 1024).toFixed(1) + ' MB';
  };

  async function api(path, force) {
    if (!force && cache.has(path)) return cache.get(path);
    const response = await fetch(API + pathUrl(path) + '?ref=' + encodeURIComponent(REF), {
      headers: {Accept: 'application/vnd.github+json'},
      cache: 'no-store'
    });
    if (!response.ok) {
      if (response.status === 403) throw new Error('GitHub API 조회 한도에 도달했습니다. 잠시 후 새로고침해 주세요.');
      if (response.status === 404) throw new Error('product_catalog 폴더를 찾지 못했습니다.');
      throw new Error('제품 카탈로그를 불러오지 못했습니다. HTTP ' + response.status);
    }
    const data = await response.json();
    cache.set(path, data);
    return data;
  }

  async function loadManifest(items) {
    const item = (items || []).find(x => x.type === 'file' && x.name.toLowerCase() === 'catalog.json');
    if (!item || !item.download_url) return null;
    try {
      const response = await fetch(item.download_url, {cache:'no-store'});
      if (!response.ok) return null;
      return await response.json();
    } catch (_) {
      return null;
    }
  }

  async function walkProducts(path, relative, inheritedManifest, depth) {
    const level = Number(depth) || 0;
    if (level > 5) return [];
    const items = await api(path);
    const localManifest = await loadManifest(items);
    const manifest = localManifest || inheritedManifest || null;
    const group = (localManifest && localManifest.displayName) || (inheritedManifest && inheritedManifest.displayName) || relative || state.category;
    const direct = (items || [])
      .filter(x => x.type === 'file' && /\.pdf$/i.test(x.name))
      .map(file => ({file, manifest, group}));
    const pdfModels = new Set(direct.map(entry => modelFromFile(entry.file.name)));
    const virtual = localManifest ? Object.entries(localManifest.products || {})
      .filter(([key, meta]) => meta && (meta.officialUrl || meta.url) && !pdfModels.has(modelFromFile(key)))
      .map(([key, meta]) => ({
        file: {name:key, size:0, html_url:meta.officialUrl || meta.url, virtual:true},
        manifest: localManifest,
        group
      })) : [];
    const dirs = (items || []).filter(x => x.type === 'dir').sort((a,b) => a.name.localeCompare(b.name));
    if (!dirs.length) return direct.concat(virtual);
    const nested = await Promise.all(dirs.map(dir =>
      walkProducts(path + '/' + dir.name, relative ? relative + ' › ' + dir.name : dir.name, manifest, level + 1)
    ));
    return direct.concat(virtual, ...nested);
  }

  function shell() {
    return '<section class="panel catalogPanel">' +
      '<div class="catalog-head"><div><h2>업체별 부품 리스트</h2><p>GitHub의 <code>product_catalog</code> 폴더를 기준으로 업체 → 부품군 → 하위 제품군 → 제품 자료를 탐색합니다. 하위 폴더가 여러 단계여도 PDF를 자동 탐색하며, 각 폴더의 <code>catalog.json</code>에 등록된 핵심 사양은 제품 카드에 함께 표시됩니다.</p></div>' +
      '<div class="catalog-head-actions"><button type="button" id="catalogKoreaBtn" class="catalog-refresh" aria-expanded="false">한국 업체·대리점 조회</button><button type="button" id="catalogRefresh" class="catalog-refresh">카탈로그 새로고침</button></div></div>' +
      '<div class="catalog-steps"><span class="active">1 업체</span><span>2 부품군</span><span>3 제품·스펙</span></div>' +
      '<div id="catalogNotice" class="catalog-notice">카탈로그를 불러오는 중입니다.</div>' +
      '<section id="catalogKoreaPanel" class="catalog-korea-panel" hidden><div class="catalog-korea-head"><div><h3>한국 업체·대리점 / 국내 문의처</h3><p>업체별 부품 리스트 제조사의 한국 법인·공급 파트너·공식 문의 경로 포함 · 2026-10-04 확인</p></div></div><div id="catalogKoreaBody" class="catalog-korea-body"><p class="catalog-empty">업체 정보를 불러오는 중입니다.</p></div></section>' +
      '<div class="catalog-layout">' +
        '<section class="catalog-column"><div class="catalog-column-head"><h3>1. 업체</h3><span id="catalogCompanyCount">—</span></div><div id="catalogCompanies" class="catalog-list"></div></section>' +
        '<section class="catalog-column"><div class="catalog-column-head"><h3>2. 부품군</h3><span id="catalogCategoryCount">—</span></div><div id="catalogCategories" class="catalog-list"><p class="catalog-empty">업체를 선택하세요.</p></div></section>' +
        '<section class="catalog-products-column"><div class="catalog-products-head"><div><h3>3. 제품·간략 스펙</h3><p id="catalogContext">부품군을 선택하면 등록된 PDF 제품이 표시됩니다.</p></div><div class="catalog-products-tools"><div id="catalogViewMode" class="catalog-view-mode" role="group" aria-label="제품 정리 방식"><button type="button" data-catalog-view="cards" class="active" aria-pressed="true">카드형</button><button type="button" data-catalog-view="table" aria-pressed="false">표형</button></div><input id="catalogSearch" type="search" placeholder="제품명 / 모델 / 사양 검색" disabled></div></div><div id="catalogFamilyFilters" class="catalog-family-filters" hidden></div><div id="catalogProducts" class="catalog-products"><p class="catalog-empty">제품 자료를 선택하세요.</p></div></section>' +
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
    $('catalogContext').textContent = '부품군을 선택하면 등록된 PDF 제품이 표시됩니다.';

    try {
      const items = await api(BASE, force);
      const companies = (items || []).filter(x => x.type === 'dir').sort((a,b) => a.name.localeCompare(b.name));
      $('catalogCompanyCount').textContent = companies.length + '개';
      $('catalogCompanies').innerHTML = companies.length ? companies.map(x => button(x.name, 'company', false)).join('') : '<p class="catalog-empty">등록된 업체 폴더가 없습니다.</p>';
      $('catalogCategoryCount').textContent = '—';
      notice(companies.length + '개 업체가 등록되어 있습니다. 업체를 선택하세요.', 'ready');
      bindCompanyButtons();
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
      const items = await api(BASE + '/' + company);
      const categories = (items || []).filter(x => x.type === 'dir').sort((a,b) => a.name.localeCompare(b.name));
      $('catalogCategoryCount').textContent = categories.length + '개';
      $('catalogCategories').innerHTML = categories.length ? categories.map(x => button(x.name, 'category', false)).join('') : '<p class="catalog-empty">등록된 부품군 폴더가 없습니다.</p>';
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

  function entryData(entry) {
    const file = entry.file;
    const manifest = entry.manifest;
    const group = entry.group || state.category;
    const rawModel = modelFromFile(file.name);
    const meta = productMeta(manifest, file, rawModel);
    const model = meta.characteristicsOnly ? 'Characteristics' : rawModel;
    const specs = specObject(manifest, meta);
    const title = meta.name || model;
    const description = meta.description || (manifest && manifest.productDescription) || '';
    const officialUrl = meta.officialUrl || (manifest && manifest.officialUrl) || '';
    const checked = meta.checked || (manifest && manifest.checked) || '';
    const searchText = [title, model, file.name, state.company, state.category, group, description]
      .concat(Object.entries(specs).flat()).join(' ').toLowerCase();
    return {file, manifest, group, model, specs, title, description, officialUrl, checked, searchText};
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
    const primaryLabel = virtual ? '공식 제품 페이지 보기 ↗' : '제품 PDF 보기 ↗';
    const secondary = (!virtual && d.officialUrl && d.officialUrl !== d.file.html_url)
      ? '<a href="' + esc(d.officialUrl) + '" target="_blank" rel="noopener">공식 제품 페이지 ↗</a>' : '';
    return '<article class="catalog-card" data-family="' + esc(d.group) + '" data-search="' + esc(d.searchText) + '">' +
      '<div class="catalog-card-top"><div><span class="catalog-vendor">' + esc(cleanLabel(state.company)) + '</span><span class="catalog-family">' + esc(d.group) + '</span><h4>' + esc(d.title) + '</h4><code>' + esc(d.model) + '</code></div><span class="catalog-file-size">' + (virtual ? 'URL' : esc(humanBytes(d.file.size))) + '</span></div>' +
      (d.description ? '<p class="catalog-description">' + esc(d.description) + '</p>' : '') +
      (rows.length ? '<dl class="catalog-specs">' + rows.map(([key,value]) => '<div><dt>' + esc(key) + '</dt><dd>' + esc(value) + '</dd></div>').join('') + '</dl>' :
        '<div class="catalog-no-spec">공통 비교 스펙 미등록 · catalog.json에 comparisonFields를 추가해야 합니다.</div>') +
      '<div class="catalog-card-actions">' + (primaryUrl ? '<a href="' + esc(primaryUrl) + '" target="_blank" rel="noopener">' + primaryLabel + '</a>' : '') + secondary + '</div>' +
      (d.checked ? '<small class="catalog-checked">사양 확인일 ' + esc(d.checked) + '</small>' : '') +
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
      const headers = ['모델', ...keys, '자료', '공식 페이지'];
      const rows = data.map(d => '<tr class="catalog-table-row" data-family="' + esc(d.group) + '" data-search="' + esc(d.searchText) + '">' +
        '<td><b>' + esc(d.title) + '</b><small>' + esc(d.model) + '</small></td>' +
        keys.map(key => '<td>' + esc(d.specs[key] ?? '—') + '</td>').join('') +
        '<td>' + ((d.file.html_url || d.officialUrl) ? '<a href="' + esc(d.file.html_url || d.officialUrl) + '" target="_blank" rel="noopener">' + (d.file.virtual ? 'URL ↗' : 'PDF ↗') + '</a>' : '—') + '</td>' +
        '<td>' + (!d.file.virtual && d.officialUrl && d.officialUrl !== d.file.html_url ? '<a href="' + esc(d.officialUrl) + '" target="_blank" rel="noopener">공식 ↗</a>' : (d.file.virtual ? 'URL 제품' : '—')) + '</td></tr>').join('');
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
      const path = BASE + '/' + state.company + '/' + category;
      const entries = await walkProducts(path, '', null, 0);
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
      notice(state.company + ' · ' + category + ' · PDF ' + entries.length + '개 · 제품군 ' + families.length + '개' + (specCount ? ' · 스펙 연결 ' + specCount + '개' : ' · catalog.json 미등록'), specCount ? 'ready' : 'review');
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
      cache.clear();
      loadCompanies(true);
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