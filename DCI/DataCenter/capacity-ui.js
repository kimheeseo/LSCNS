(() => {
  'use strict';
  const M = window.DCCapacity, $ = id => document.getElementById(id);
  const esc = x => String(x ?? '').replace(/[&<>"']/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[c]));
  const fmt = x => typeof x === 'number' ? x.toLocaleString('ko-KR', {maximumFractionDigits: 2}) : String(x ?? '—');
  const number = id => $(id).value.trim() === '' ? undefined : Number($(id).value);
  const names = {compute: '컴퓨트/데이터', storage: '스토리지', management: '관리망', dci: '건물 간/DCI', shared: '공유 배선'};
  const labels = {rackKw: '랙당 IT 부하', pue: '목표 PUE', speed: '컴퓨트 링크 속도', oversub: '다운/업 대역폭 비율', gpuPerServer: 'GPU/서버', serverKw: '서버 전력', serverRu: '서버 RU', computePorts: '외부 컴퓨트 포트/기준', computePortBasis: '컴퓨트 포트 기준', storagePorts: '스토리지 포트/서버', managementPorts: '관리 포트/서버', computePowerShare: 'IT 부하 중 compute 비율', usableRackRu: 'Compute 설치 가능 RU', rackLimitRu: '전체 랙 RU', racksPerBlock: '랙/블록', blocksPerHall: '블록/홀', hallsPerBuilding: '홀/건물', buildingDistanceM: '건물 간 경로 거리', rowLengthM: '랙 열 경로 길이', rackDistanceM: '랙 내 평균 경로', hallDistanceM: '홀 내 평균 경로', fabricDistanceM: 'Leaf-Spine 평균 경로', superSpineDistanceM: 'Spine-Super-spine 평균 경로', endpointDistribution: '랙/열/홀 링크 분포', sparePct: '구매 예비품', uncertaintyPct: '계획 수량 범위 ±', slackPct: '배선 여유장', endAllowanceM: '양단 작업 여유장', panelTerminations: '패널당 종단그룹', cableFibers: 'DCI 케이블 심수', fiberReservePct: 'DCI 설치 예비심선', gbpsPerRack: 'DCI 요구 대역폭/랙', wdmChannels: '심선쌍당 WDM 채널', redundancy: '이중화', leafDown: 'Leaf 다운 포트', leafUp: 'Leaf 업 포트', spinePorts: 'Spine 전체 포트', spineUp: 'Spine 코어 예약 포트', coreOversub: '블록 간 오버서브', superSpinePorts: 'Super-spine 포트', minSpines: '블록당 최소 Spine', trayAreaMm2: '트레이 단면적', trayFill: '트레이 점유율', cableDiameterMm: '평균 케이블 외경', racksPerTrayRoute: '공유 트레이 경로당 랙', rackAreaM2: '랙당 IT 면적', siteAreaMultiplier: '총 부지 배수', powerModuleMw: '참고 전력 모듈 용량', coolingModuleKw: '참고 냉각 모듈 용량', airCoolingLimitKw: '공랭 경고 임계값', dacMaxM: 'DAC 최대 선택 길이', aocMaxM: 'AOC 최대 선택 길이', baseTMaxM: '관리망 copper 최대 길이', cpoPortsPerEngine: 'CPO engine당 포트', cpoExternalLasersPerEngine: 'CPO engine당 ELS', topology: '토폴로지', protocol: '프로토콜', cooling: '냉각', assemblySpeeds: 'DAC/AOC 계획 속도 목록', endsPerLink: '링크 양단 수'};
  let config, overrides = {}, customPresets = {}, active = false, latest, alternativeId = 'base', phaseDefinitions = [], timer, catalogCache, renderToken = 0;
  const key = 'dcCapacityDesign.v8';
  function save() {
    try { localStorage.setItem(key, JSON.stringify({active, config, input: readInput(), overrides, customPresets, phases: phaseDefinitions})); } catch (e) { console.warn('Capacity settings not saved', e); }
  }
  function options(object, current) { return Object.entries(object).map(([id, v]) => '<option value="' + esc(id) + '"' + (id === current ? ' selected' : '') + '>' + esc(v.name) + '</option>').join(''); }
  function table(headers, rows) { return '<div class="tableWrap cap-table"><table><thead><tr>' + headers.map(x => '<th>' + esc(x) + '</th>').join('') + '</tr></thead><tbody>' + rows.map(r => '<tr>' + r.map(x => '<td' + (/^[\d,.\s/–±%]+(?:MW|kW|G)?$/.test(String(x).split('<br>')[0]) ? ' class="cap-number"' : '') + '>' + x + '</td>').join('') + '</tr>').join('') + '</tbody></table></div>'; }
  function field(id, label, type = 'number', value = '', extra = '') {return '<label><span>' + esc(label) + '</span><input id="' + id + '" type="' + type + '" value="' + esc(value) + '" ' + extra + '></label>';}
  function select(id, label, items) {return '<label><span>' + label + '</span><select id="' + id + '">' + items.map(([v, t]) => '<option value="' + v + '">' + t + '</option>').join('') + '</select></label>';}
  function fillPowerDefaults(resetRack = false, resetPue = false) {
    const workloadPath = 'workloadPresets.' + $('cap-workload').value;
    const equipmentId = $('cap-equipment').value;
    const equipment = customPresets[equipmentId] || config.equipmentPresets[equipmentId];
    const valueAt = path => overrides[path] ?? path.split('.').reduce((value, part) => value[part], config).value;
    let textRackKw;
    try {textRackKw = M.parseScale($('cap-scale').value, $('cap-kind').value, config.units).rackKw;} catch {}
    const rackPath = equipment.rackScale ? 'equipmentPresets.' + equipmentId + '.rackKw' : workloadPath + '.rackKw';
    const rackValue = textRackKw ?? (equipment.rackScale && customPresets[equipmentId] ? overrides[rackPath] ?? equipment.rackKw.value : valueAt(rackPath));
    for (const [id, value, reset] of [['cap-rack', rackValue, resetRack], ['cap-pue', valueAt(workloadPath + '.pue'), resetPue]]) {
      const input = $(id);
      if (reset || input.dataset.presetDefault === 'true' || input.value === '') {input.value = value; input.dataset.presetDefault = 'true';}
    }
  }
  function powerInput(id) {return $(id).dataset.presetDefault === 'true' ? undefined : number(id);}
  function buildInput() {
    const form = document.querySelector('.formPanel');
    const launch = document.createElement('div'); launch.className = 'cap-launch';
    launch.innerHTML = '<button type="button" id="capacityButton" aria-expanded="false" aria-controls="capacity-input">용량</button><button type="button" id="capacityLegacy" hidden>기존 장비별 설계</button>';
    form.querySelector('h2').after(launch);
    const section = document.createElement('section'); section.id = 'capacity-input'; section.hidden = true;
    section.innerHTML = '<div class="cap-mode" role="group" aria-label="용량 설계 입력 모드"><button type="button" id="cap-quick" aria-pressed="true">빠른 견적</button><button type="button" id="cap-detail" aria-pressed="false">상세 설계</button></div>' +
      '<div class="cap-fields">' + field('cap-scale', '고객 요구 규모', 'text', 'GPU 10,000개', 'placeholder="15GW / 랙 2,000개, 랙당 10kW"') + select('cap-kind', '입력 단위', [['auto', '문장에서 자동 판별'], ['MW', 'MW'], ['GW', 'GW'], ['gpu', 'GPU 수'], ['racks', '랙 수'], ['servers', '서버 수']]) +
      select('cap-basis', '전력 입력 기준', [['it', 'IT 부하'], ['facility', '시설 전체 전력']]) + '<label><span>용도 프리셋</span><select id="cap-workload">' + options(config.workloadPresets, config.defaults.workload) + '</select></label><label><span>장비 프리셋</span><select id="cap-equipment">' + options(config.equipmentPresets, 'generic8') + '</select></label>' +
      field('cap-rack', '랙당 IT 부하 kW', 'number', '', 'min="0.001" step="any"') + field('cap-pue', '목표 PUE', 'number', '', 'min="1" step="any"') + '</div><p class="muted">용도·장비의 계획 기본값입니다. 실제 조건에 맞게 수정할 수 있습니다.</p><p id="cap-preview" class="muted" aria-live="polite"></p>' +
      '<details id="cap-advanced"><summary>제약조건·가정 편집</summary><div class="cap-fields">' + select('cap-tier', '요구 Tier', [['unspecified', '미정'], ['Tier I', 'Tier I'], ['Tier II', 'Tier II'], ['Tier III', 'Tier III'], ['Tier IV', 'Tier IV']]) +
      field('cap-power-limit', '전력 인입 한도 MW · 시설 전력', 'number', '', 'min="0.001" step="any"') + field('cap-site-area', '부지 면적 m²', 'number', '', 'min="1" step="any"') + field('cap-buildings', '건물 수 · 빈칸은 자동', 'number', '', 'min="1" step="1"') +
      select('cap-speed', '컴퓨트 링크 속도', [['', '프리셋 기본값'], ['400', '400G'], ['800', '800G'], ['1600', '1.6T · 검증 필요']]) + select('cap-media', '선호 매체', [['auto', '자동'], ['DAC', 'DAC'], ['AOC', 'AOC'], ['MMF', 'MMF'], ['SMF', 'SMF']]) +
      select('cap-cooling', '냉각 방식', [['', '프리셋 기본값'], ['air', '공랭'], ['liquid', '액체 냉각']]) + select('cap-protocol', '컴퓨트 프로토콜', [['', '프리셋 기본값'], ['InfiniBand', 'InfiniBand'], ['RoCE', 'RoCE'], ['Ethernet', 'Ethernet']]) + select('cap-topology', '토폴로지', [['', '프리셋 기본값'], ['two', '블록별 Leaf-Spine'], ['three', 'Leaf-Spine-Super-spine']]) +
      field('cap-budget', '예산 범위 · 단가 미정', 'text') + '<label class="cap-check"><input type="checkbox" id="cap-cpo"><span>CPO 적용 검토안 포함</span></label></div><h3>네트워크별 이중화</h3><div class="cap-fields">' + Object.entries(names).filter(([k]) => k !== 'shared').map(([k, name]) => select('cap-dual-' + k, name, [['1', '단일'], ['2', k === 'dci' ? '이중 경로' : '이중 패브릭']])).join('') +
      '</div><h3>단계별 증설 · 블록 범위</h3><div id="cap-phases"></div><button type="button" id="cap-add-phase">단계 범위 추가</button><p class="muted">미지정 블록은 Phase 1. 초기 단계에 공용 코어·DCI를 설치할 계획은 별도 확인 필요.</p>' +
      '<h3>설정 계수</h3><label><span>계수 그룹</span><select id="cap-coefficient-group"><option value="selected">현재 용도·장비</option><option value="coefficients">공통 BOM·전력</option><option value="layout">배치·거리</option><option value="networks">네트워크</option><option value="mediaRules">매체 선택</option></select></label><div id="cap-coefficients"></div>' +
      '<h3>장비 프리셋 관리</h3><div class="cap-fields">' + field('cap-preset-name', '새 장비 프리셋 이름', 'text', '', 'placeholder="예: 차세대 8-GPU 서버"') + '</div><button type="button" id="cap-add-preset">현재 장비를 복제하여 추가</button>' +
      '<div class="cap-actions"><button type="button" id="cap-settings-export">설정 JSON 내보내기</button><label class="cap-file">설정 가져오기<input type="file" id="cap-settings-import" accept=".json,application/json"></label><button type="button" id="cap-reset">가정 초기화</button></div>' +
      '<details><summary>전체 설정 JSON 편집 · 매체 규격 포함</summary><textarea id="cap-config-json" aria-label="전체 설정 JSON" spellcheck="false"></textarea><button type="button" id="cap-config-apply">설정 적용</button></details></details>' +
      '<button type="button" id="cap-run" class="primary">설계안 계산</button><p id="cap-error" class="error" role="alert"></p>';
    launch.after(section);
    fillPowerDefaults(true, true);
    $('capacityButton').onclick = () => {activate(true); recompute();};
    $('capacityLegacy').onclick = () => activate(false);
    $('cap-run').onclick = recompute;
    $('cap-quick').onclick = () => detailMode(false);
    $('cap-detail').onclick = () => detailMode(true);
    $('cap-advanced').addEventListener('toggle', () => {$('cap-quick').setAttribute('aria-pressed', String(!$('cap-advanced').open)); $('cap-detail').setAttribute('aria-pressed', String($('cap-advanced').open));});
    $('cap-workload').addEventListener('change', () => {$('cap-equipment').value = config.workloadPresets[$('cap-workload').value].equipment; fillPowerDefaults(true, true); renderCoefficients();});
    $('cap-equipment').addEventListener('change', () => {fillPowerDefaults(true); renderCoefficients();});
    $('cap-coefficient-group').onchange = renderCoefficients;
    section.addEventListener('input', event => {
      if (event.target.id === 'cap-config-json' || event.target.id === 'cap-preset-name' || event.target.type === 'file') return;
      if (event.target.id === 'cap-rack' || event.target.id === 'cap-pue') event.target.dataset.presetDefault = 'false';
      if (event.target.id === 'cap-scale') {
        try {const parsed = M.parseScale(event.target.value, $('cap-kind').value, config.units); if (parsed.detectedEquipment && parsed.detectedEquipment !== $('cap-equipment').value) {$('cap-equipment').value = parsed.detectedEquipment; fillPowerDefaults(true); renderCoefficients();} else fillPowerDefaults();}
        catch {}
      }
      if (event.target.dataset.coefficient) {
        try {
          const merged = {...config, equipmentPresets: {...config.equipmentPresets, ...customPresets}};
          const original = M.coefficientRows(merged).find(x => x.path === event.target.dataset.coefficient);
          overrides[original.path] = typeof original.value === 'number' ? Number(event.target.value) : typeof original.value === 'boolean' ? event.target.value === 'true' : typeof original.value === 'object' ? JSON.parse(event.target.value) : event.target.value;
          const network = original.path.match(/^networks\.(compute|storage|management|dci)\.redundancy$/);
          if (network) $('cap-dual-' + network[1]).value = String(overrides[original.path]);
          fillPowerDefaults();
        }
        catch (error) {invalidate(error); return;}
      }
      if (event.target.dataset.phaseField) readPhases();
      clearTimeout(timer); timer = setTimeout(recompute, 180);
    });
    $('cap-add-phase').onclick = () => {phaseDefinitions.push({start: 1, end: 1, phase: phaseDefinitions.length + 2}); renderPhases(); detailMode(true);};
    $('cap-add-preset').onclick = () => {
      const name = $('cap-preset-name').value.trim();
      if (!name) {invalidate(Error('새 장비 프리셋 이름을 입력하세요.')); return;}
      const id = 'custom-' + Date.now();
      const base = JSON.parse(JSON.stringify(customPresets[$('cap-equipment').value] || config.equipmentPresets[$('cap-equipment').value]));
      const allRows = M.coefficientRows({p: base});
      for (const row of allRows) {
        const fieldKey = row.path.slice(2), fullPath = 'equipmentPresets.' + $('cap-equipment').value + '.' + fieldKey;
        base[fieldKey] = {...base[fieldKey], value: overrides[fullPath] ?? row.value, source: '사용자 복제·수정; 실제 장비 사양 검증 필요', confidence: '낮음', status: '검증 필요', origin: '사용자 수정'};
      }
      base.name = name; customPresets[id] = base;
      $('cap-equipment').innerHTML = options({...config.equipmentPresets, ...customPresets}, id);
      renderCoefficients(); recompute();
    };
    $('cap-settings-export').onclick = () => download('DC-Capacity-Settings-v8.json', JSON.stringify({schemaVersion: 1, config, overrides, customPresets, input: readInput()}, null, 2), 'application/json');
    $('cap-settings-import').onchange = async event => {
      const file = event.target.files[0]; if (!file) return;
      try {
        if (file.size > 2000000) throw Error('설정 파일은 2MB 이하여야 합니다.');
        const imported = JSON.parse(await file.text());
        const nextConfig = imported.config || (imported.equipmentPresets ? imported : config);
        const nextInput = {...readInput(), ...(imported.input || {}), overrides: imported.overrides || {}, customPresets: imported.customPresets || {}};
        M.calculate(nextInput, nextConfig);
        config = nextConfig; overrides = nextInput.overrides; customPresets = nextInput.customPresets; phaseDefinitions = nextInput.phases || [];
        restoreInput(nextInput); renderCoefficients(); renderPhases(); $('cap-config-json').value = JSON.stringify(config, null, 2); catalogCache = null; recompute();
      } catch (error) {invalidate(error);}
      event.target.value = '';
    };
    $('cap-config-json').value = JSON.stringify(config, null, 2);
    $('cap-config-apply').onclick = () => {try {const next = JSON.parse($('cap-config-json').value); M.calculate(readInput(), next); config = next; catalogCache = null; fillPowerDefaults(); renderCoefficients(); recompute();} catch (error) {invalidate(error);}};
    $('cap-reset').onclick = () => {overrides = {}; phaseDefinitions = []; fillPowerDefaults(true, true); renderCoefficients(); renderPhases(); recompute();};
    renderCoefficients(); renderPhases();
  }
  function detailMode(show) {$('cap-advanced').open = show;}
  function renderPhases() {
    $('cap-phases').innerHTML = phaseDefinitions.map((p, i) => '<div class="cap-phase-row">' + ['start', 'end', 'phase'].map((f, n) => '<label><span>' + ['시작 블록', '끝 블록', 'Phase'][n] + '</span><input type="number" min="1" step="1" data-phase-index="' + i + '" data-phase-field="' + f + '" value="' + p[f] + '"></label>').join('') + '<button type="button" data-delete-phase="' + i + '" aria-label="단계 범위 삭제">×</button></div>').join('');
    $('cap-phases').querySelectorAll('[data-delete-phase]').forEach(b => b.onclick = () => {phaseDefinitions.splice(Number(b.dataset.deletePhase), 1); renderPhases(); recompute();});
  }
  function readPhases() {for (const input of $('cap-phases').querySelectorAll('input')) phaseDefinitions[Number(input.dataset.phaseIndex)][input.dataset.phaseField] = Number(input.value);}
  function renderCoefficients() {
    const merged = {...config, equipmentPresets: {...config.equipmentPresets, ...customPresets}};
    const group = $('cap-coefficient-group').value;
    const selected = ['workloadPresets.' + $('cap-workload').value + '.', 'equipmentPresets.' + $('cap-equipment').value + '.'];
    const rows = M.coefficientRows(merged).filter(x => group === 'selected' ? selected.some(p => x.path.startsWith(p)) : x.path.startsWith(group + '.'));
    $('cap-coefficients').innerHTML = rows.map(x => {
      const v = overrides[x.path] ?? x.value;
      return '<label class="cap-coefficient"><span>' + esc(labels[x.path.split('.').at(-1)] || x.path.split('.').at(-1)) + ' <small>' + esc(x.unit) + '</small></span><input data-coefficient="' + esc(x.path) + '" type="' + (typeof x.value === 'number' ? 'number' : 'text') + '" step="any" value="' + esc(typeof v === 'object' ? JSON.stringify(v) : v) + '"><small>' + esc(Object.hasOwn(overrides, x.path) ? '사용자 수정 · 검증 필요' : x.status + ' · 신뢰도 ' + x.confidence) + ' · ' + esc(x.source) + '</small></label>';
    }).join('');
  }
  function readInput() {return {scaleText: $('cap-scale').value, scaleKind: $('cap-kind').value, powerBasis: $('cap-basis').value, workload: $('cap-workload').value, equipment: $('cap-equipment').value, rackKw: powerInput('cap-rack'), pue: powerInput('cap-pue'), tier: $('cap-tier').value, powerLimitMw: number('cap-power-limit'), siteAreaM2: number('cap-site-area'), buildings: number('cap-buildings'), speed: number('cap-speed'), media: $('cap-media').value, cooling: $('cap-cooling').value || undefined, topology: $('cap-topology').value || undefined, protocol: $('cap-protocol').value || undefined, cpo: $('cap-cpo').checked, budget: $('cap-budget').value, redundancy: Object.fromEntries(['compute', 'storage', 'management', 'dci'].map(k => [k, Number($('cap-dual-' + k).value)])), phases: phaseDefinitions, overrides, customPresets};}
  function restoreInput(input) {
    $('cap-equipment').innerHTML = options({...config.equipmentPresets, ...customPresets}, input.equipment || 'generic8');
    const mapping = {scaleText: 'scale', scaleKind: 'kind', powerBasis: 'basis', workload: 'workload', equipment: 'equipment', rackKw: 'rack', pue: 'pue', tier: 'tier', powerLimitMw: 'power-limit', siteAreaM2: 'site-area', buildings: 'buildings', speed: 'speed', media: 'media', cooling: 'cooling', topology: 'topology', protocol: 'protocol', budget: 'budget'};
    for (const [k, id] of Object.entries(mapping)) if (input[k] != null) $('cap-' + id).value = input[k];
    $('cap-cpo').checked = !!input.cpo;
    for (const [key, id] of [['rackKw', 'cap-rack'], ['pue', 'cap-pue']]) $(id).dataset.presetDefault = String(input[key] == null);
    fillPowerDefaults();
    for (const [k, v] of Object.entries(input.redundancy || {})) if ($('cap-dual-' + k)) $('cap-dual-' + k).value = v;
  }
  function activate(enabled) {
    active = enabled; document.body.classList.toggle('capacity-active', enabled);
    $('capacityButton').setAttribute('aria-expanded', String(enabled)); $('capacityLegacy').hidden = !enabled; $('capacity-input').hidden = !enabled;
    document.querySelectorAll('#workbench-nav [data-view]').forEach(b => b.disabled = false);
    for (const id of ['result', 'optical', 'power', 'bom', 'evidence']) $('view-' + id)?.classList.add('cap-view');
    let used = $('capacity-assumptions');
    if (!used) {used = document.createElement('section'); used.id = 'capacity-assumptions'; used.className = 'cap-assumptions'; $('workbench-content').prepend(used);}
    used.hidden = !enabled;
    if (!enabled) {latest = null; window.DCCapacityDesign = null; if (window.DCDesign) document.dispatchEvent(new CustomEvent('dc:design', {detail: window.DCDesign}));}
    else $('tab-result').click();
    save();
  }
  function invalidate(error) {
    $('cap-error').textContent = error.message; latest = null; window.DCCapacityDesign = null;
    $('cap-preview').textContent = '입력 수정 필요';
    if (active) {for (const id of ['result', 'optical', 'power', 'bom', 'evidence']) content(id).innerHTML = '<p class="error">' + esc(error.message) + '</p><p class="muted">입력 수정 후 다시 계산합니다. 이전 결과는 내보내지 않습니다.</p>'; $('capacity-assumptions').innerHTML = '<b>유효한 설계 계산 대기</b>';}
  }
  function content(id) {
    const parent = $('view-' + id); let section = parent.querySelector('.cap-content');
    if (!section) {section = document.createElement('section'); section.className = 'cap-content'; parent.append(section);}
    return section;
  }
  function recompute() {
    if (!active) return;
    try {
      latest = M.calculate(readInput(), config); window.DCCapacityDesign = latest;
      if (!latest.alternatives.some(x => x.id === alternativeId)) alternativeId = 'base';
      $('cap-error').textContent = ''; const s = latest.summary;
      $('cap-preview').textContent = fmt(s.itKw / config.units.powerToKw.MW) + ' MW IT · ' + fmt(s.facilityKw / config.units.powerToKw.MW) + ' MW 시설 · ' + fmt(s.racks) + '랙 · ' + fmt(s.servers) + '서버 · ' + fmt(s.gpus) + 'GPU';
      document.body.classList.remove('design-stale'); $('exportWorkbook').disabled = false;
      render(); save();
    } catch (error) {invalidate(error);}
  }
  function selected() {return latest.alternatives.find(x => x.id === alternativeId);}
  function assumptionTable(rows) {return table(['가정', '사용값', '입력 출처', '근거', '신뢰도 / 상태'], rows.map(x => ['<button type="button" class="cap-assumption-edit" data-assumption="' + esc(x.path) + '">' + esc(labels[x.path.split('.').at(-1)] || x.path) + '</button>', esc(typeof x.value === 'object' ? JSON.stringify(x.value) : fmt(x.value)) + ' ' + esc(x.unit), esc(x.origin), esc(x.source) + (x.sourceUrl ? ' <a target="_blank" rel="noopener" href="' + esc(x.sourceUrl) + '">출처</a>' : ''), esc(x.confidence + ' / ' + x.status)]));}
  function editAssumption(path) {
    detailMode(true);
    const direct = {rackKw: 'cap-rack', pue: 'cap-pue', speed: 'cap-speed', cooling: 'cap-cooling', topology: 'cap-topology', protocol: 'cap-protocol', buildingCount: 'cap-buildings'};
    const field = path.split('.').at(-1);
    if (field === 'redundancy') {const el = $('cap-dual-' + path.split('.')[1]); if (el) {el.focus(); return;}}
    if (direct[field] && ['workloadPresets', 'equipmentPresets', 'layout'].includes(path.split('.')[0])) {const el = $(direct[field]); el.focus(); el.scrollIntoView({block: 'center', behavior: 'smooth'}); return;}
    const prefix = path.split('.')[0];
    if (prefix === 'workloadPresets' || prefix === 'equipmentPresets') $('cap-coefficient-group').value = 'selected';
    else $('cap-coefficient-group').value = prefix;
    renderCoefficients();
    const input = [...$('cap-coefficients').querySelectorAll('input')].find(x => x.dataset.coefficient === path);
    if (input) {input.focus(); input.scrollIntoView({block: 'center', behavior: 'smooth'});}
    else {$('cap-config-json').closest('details').open = true; $('cap-config-json').focus();}
  }
  function bomTable(rows) {return table(['구간 / 망', 'Phase', '품목', '매체', '규격', '설치', '구매 범위', '단위', '비고'], rows.map(x => [esc(x.segment + ' / ' + names[x.network]), fmt(x.phase), esc(x.item), esc(x.media), esc(x.spec), fmt(x.installedQty), fmt(x.low) + '–' + fmt(x.high) + '<br><small>기준 ' + fmt(x.purchaseQty) + ' · ±' + x.uncertaintyPct + '%</small>', esc(x.unit), esc(x.note)]));}
  function render() {
    const r = latest, s = r.summary, a = selected(), token = ++renderToken;
    const uncertainty = r.assumptions.find(x => x.path === 'coefficients.uncertaintyPct').value;
    const quantityRange = value => fmt(Math.floor(value * (1 - uncertainty / 100))) + '–' + fmt(Math.ceil(value * (1 + uncertainty / 100)));
    const keyPaths = ['rackKw', 'pue', 'gpuPerServer', 'serverKw', 'computePorts', 'computePowerShare', 'racksPerBlock', 'endpointDistribution', 'uncertaintyPct'];
    const used = r.assumptions.filter(x => keyPaths.includes(x.path.split('.').at(-1)));
    $('capacity-assumptions').innerHTML = '<div class="cap-heading"><b>이번 계산의 가정</b><span>v' + esc(r.version) + ' · 비용: 산정 불가</span></div><div class="cap-used">' + used.map(x => '<button type="button" data-assumption="' + esc(x.path) + '"><span>' + esc(labels[x.path.split('.').at(-1)] || x.path) + '</span><b>' + esc(typeof x.value === 'object' ? Object.values(x.value).map(fmt).join(' / ') : fmt(x.value)) + ' ' + esc(x.unit) + '</b><small>' + esc(x.origin + ' · ' + x.status) + '</small></button>').join('') + '</div><details><summary>사용된 모든 가정·출처 (' + r.assumptions.length + ')</summary>' + assumptionTable(r.assumptions) + '</details>';
    $('capacity-assumptions').querySelectorAll('[data-assumption]').forEach(b => b.onclick = () => editAssumption(b.dataset.assumption));
    const kpis = [['IT 부하', fmt(s.itKw / config.units.powerToKw.MW) + ' MW'], ['시설 전체 전력', fmt(s.facilityKw / config.units.powerToKw.MW) + ' MW'], ['랙 / 서버', fmt(s.racks) + ' / ' + fmt(s.servers)], ['GPU 수용', fmt(s.gpus)], ['건물 / 홀 / 블록', fmt(s.buildings) + ' / ' + fmt(s.halls) + ' / ' + fmt(s.blocks)], ['포트 / 링크', fmt(s.endpointPorts) + ' / ' + fmt(s.links)]];
    content('result').innerHTML = '<h2>용량 기반 설계안</h2><div class="cap-kpis">' + kpis.map(([name, v]) => '<div><span>' + esc(name) + '</span><b>' + esc(v) + '</b></div>').join('') + '</div>' +
      '<p>' + esc(config.workloadPresets[r.workload].name) + ' · ' + esc((r.input.customPresets?.[r.equipmentId] || config.equipmentPresets[r.equipmentId]).name) + ' · ' + esc(s.cooling === 'air' ? '공랭' : '액체 냉각') + '</p><p class="muted">IT 부하에는 compute ' + fmt(s.computeKw / config.units.powerToKw.MW) + ' MW와 네트워크·스토리지·여유 ' + fmt(s.otherItKw / config.units.powerToKw.MW) + ' MW가 포함됩니다. 랙 수용 전력 ' + fmt(s.rackCapacityKw / config.units.powerToKw.MW) + ' MW는 요청 전력과 구분됩니다.</p>' +
      '<h3>설계 대안 비교</h3>' + table(['안', '구매 케이블 assembly/트렁크 조수', '광모듈', '패널', '트레이/덕트 m', '장점', '검토 사항'], r.alternatives.map(x => ['<button type="button" data-alternative="' + x.id + '" aria-pressed="' + (x.id === alternativeId) + '">' + esc(x.name) + '</button>', quantityRange(x.totals.cables), quantityRange(x.totals.transceivers), quantityRange(x.totals.panels), quantityRange(x.totals.trayM), esc(x.pros), esc(x.cons)])) +
      '<p class="muted">수량은 설치량에 예비품을 더한 비교 기준값입니다. 상세 BOM에는 ±범위를 표시합니다. 커넥터 종단과 CPO 참고 수량은 케이블 구매량에 중복 합산하지 않습니다.</p><h3>전력 → 랙 → GPU → 포트 → 링크 → 케이블</h3>' +
      table(['중간 결과', '값', '산식 / 해석'], [['IT 전력', fmt(s.itKw) + ' kW', '전력 기준 및 PUE로 환산'], ['랙', fmt(s.racks), '전력 입력: ceil(IT kW / 랙 kW)'], ['서버 / GPU', fmt(s.servers) + ' / ' + fmt(s.gpus), '전력 배분·RU·장비 프리셋; 랙 스케일은 완전한 랙 기준'], ['서버 endpoint 포트', fmt(s.endpointPorts), '각 망의 서버 수 × 포트 × 패브릭 복제'], ['설치 링크', fmt(s.links), 'endpoint + leaf-spine + super-spine + DCI'], ['구매 케이블 조수', fmt(a.totals.cables), '일체형/광 assembly + DCI trunk + 예비품; 패치리드 별도'], ['내부 산식 검증', Object.values(r.validation).every(Boolean) ? '통과' : '실패', '수량·전력·포트·Phase 합계; 실제 시공 검증 아님']]) +
      '<h3>다음에 확인할 질문 3개</h3><ol>' + r.questions.map(q => '<li>' + esc(q) + '</li>').join('') + '</ol><details open><summary>검토 사항 (' + r.warnings.length + ')</summary><ul class="cap-warnings">' + r.warnings.map(w => '<li>' + esc(w) + '</li>').join('') + '</ul></details>';
    content('result').querySelectorAll('[data-alternative]').forEach(b => b.onclick = () => {alternativeId = b.dataset.alternative; render();});
    content('optical').innerHTML = '<h2>' + esc(a.name) + ' · 구간별 연결</h2>' + table(['네트워크', 'Phase', '구간', '경로/유효 길이 m', '링크', '매체', '규격', '사용/설치 심수 · 링크당'], a.selections.map(x => [esc(names[x.network]), fmt(x.phase), esc(x.tier + ' / ' + x.distanceClass), fmt(x.distanceM) + ' / ' + fmt(x.lengthM), fmt(x.links), esc(x.media), esc(x.standard + ' · ' + x.connector), x.network === 'dci' ? 'DCI 별도 표 참조' : fmt(x.activeFibers) + ' / ' + fmt(x.installedFibers)])) +
      '<h3>건물 간 / DCI · SMF</h3>' + (r.dci.rows.length ? table(['건물', 'Phase', '속도', '링크', '사용심선', '설치심선', '예비심선', '케이블 조수', '채널/심선쌍'], r.dci.rows.map(x => [fmt(x.building), fmt(x.phase), fmt(config.networks.dci.speed.value) + 'G', fmt(x.links), fmt(x.activeFibers), fmt(x.installedFibers), fmt(x.installedFibers - x.activeFibers), fmt(x.cableCount), fmt(x.channels)])) : '<p class="muted">단일 건물: 건물 간 DCI 수량 0. 외부 통신사 연결은 별도 요구조건 필요.</p>') + '<h3>랙 내부 참고</h3><p>' + esc(r.internalReference) + '</p>';
    content('power').innerHTML = '<h2>전력·냉각 / 배치 계획</h2>' + table(['항목', '계획값', '비고'], [['PUE', fmt(s.pue), '목표 입력; 측정값 아님'], ['시설 전력', fmt(s.facilityKw / config.units.powerToKw.MW) + ' MW', 'IT × PUE; 인입 한도 비교 기준'], ['인입 한도', r.input.powerLimitMw == null ? '미정' : fmt(r.input.powerLimitMw) + ' MW', r.input.powerLimitMw != null && s.facilityKw > r.input.powerLimitMw * config.units.powerToKw.MW ? '초과' : '입력 조건 기준'], ['냉각', esc(s.cooling), 'IT 열부하 ' + fmt(s.itKw) + ' kW'], ['Tier 요구', esc(r.tier), '인증 충족 판정 아님'], ['계획 부지 면적', fmt(s.siteAreaM2) + ' m²', '랙 면적 × 총 부지 배수; 임의 가정']]) +
      '<h3>Phase별 수량</h3>' + table(['Phase', '블록', '랙', '서버', 'GPU', 'IT MW', '누적 IT MW', '링크', '구매 케이블'], r.phaseRows.map((x, i) => [fmt(x.phase), fmt(x.blocks), fmt(x.racks), fmt(x.servers), fmt(x.gpus), fmt(x.itKw / config.units.powerToKw.MW), fmt(r.phaseRows.slice(0, i + 1).reduce((n, p) => n + p.itKw, 0) / config.units.powerToKw.MW), fmt(x.links), fmt(x.cables)])) + '<h3>블록 배치 · 일부 표시</h3><div class="cap-actions">' + field('cap-block-search', '블록 번호', 'number', '', 'min="1" max="' + s.blocks + '"') + '<button type="button" id="cap-find-block">블록 조회</button></div><div id="cap-block-table"></div>';
    const blocksTable = rows => table(['블록', '건물', '홀', 'Phase', '랙', '서버', 'GPU', 'IT kW'], rows.map(b => [b.id, b.building, b.hall, b.phase, fmt(b.racks), fmt(b.servers), fmt(b.gpus), fmt(b.itKw)]));
    $('cap-block-table').innerHTML = blocksTable(r.blocks.slice(0, 20));
    $('cap-find-block').onclick = () => {const id = number('cap-block-search'); $('cap-block-table').innerHTML = blocksTable(id ? r.blocks.filter(x => x.id === id) : r.blocks.slice(0, 20));};
    content('bom').innerHTML = '<div class="cap-heading"><h2>' + esc(a.name) + ' · BOM</h2><div class="cap-actions"><button type="button" id="cap-csv">CSV</button><button type="button" id="cap-xlsx">Excel</button></div></div><p class="muted">비용: 산정 불가 · 가격 자료 미등록 · ±' + esc(a.bom[0]?.uncertaintyPct ?? 0) + '%는 계획 수량 범위입니다.</p>' + bomTable(a.bom.filter(x => !x.referenceOnly)) + '<h3>설계 참고 수량</h3>' + table(['품목', '수량', '단위', '비고'], r.references.map(x => [esc(x.item), fmt(x.qty), esc(x.unit), esc(x.note)])) + '<h3>종단·CPO 참고 · 중복 구매 제외</h3>' + bomTable(a.bom.filter(x => x.referenceOnly));
    $('cap-csv').onclick = exportCsv; $('cap-xlsx').onclick = exportXlsx;
    content('evidence').innerHTML = '<h2>규격 기반 제품 후보 / 검증</h2><p>필수 속도·규격·패키지·커넥터·거리 근거가 모두 확인된 광모듈만 후보로 연결합니다. SKU 확정에는 host/FEC/광손실·극성 검증이 필요합니다.</p><div id="cap-candidates" aria-live="polite">관련 카탈로그 조회 중…</div><h3>계산 검증</h3>' + table(['검증 항목', '결과'], Object.entries(r.validation).map(([k, v]) => [esc(k), v ? '통과' : '실패'])) + '<h3>논리 토폴로지 집계</h3>' + table(['망', '프로토콜', '서버 endpoint', 'Leaf', 'Spine', 'Leaf-Spine 링크', 'Spine-Core 링크'], ['compute', 'storage', 'management'].map(k => {const xs = r.networks.filter(x => x.network === k), total = f => xs.reduce((n, x) => n + x[f], 0); return [names[k], esc(xs[0]?.protocol || '없음'), fmt(total('endpoints')), fmt(total('leaf')), fmt(total('spine')), fmt(total('fabricLinks')), fmt(total('coreLinks'))];}));
    loadCandidates(a, token);
  }
  async function loadCandidates(a, token) {
    try {
      catalogCache ||= Promise.all(config.catalogPaths.map(async path => {try {const res = await fetch(new URL(path, location.href)); if (!res.ok) throw Error('HTTP ' + res.status); return {path, catalog: await res.json()};} catch (e) {return {path, error: e.message};}}));
      const catalogs = await catalogCache;
      if (token !== renderToken || !latest) return;
      const rows = [], evidenceRows = [];
      for (const bom of a.bom.filter(x => x.transceiver)) {
        const candidates = catalogs.filter(x => x.catalog).flatMap(x => M.matchCatalog(bom.requirement, x.catalog));
        evidenceRows.push({segment: names[bom.network] + ' / ' + bom.segment, spec: bom.spec, candidates});
        rows.push([esc(names[bom.network] + ' / ' + bom.segment), esc(bom.spec), candidates.length ? candidates.slice(0, 3).map(c => '<a href="' + esc(c.source) + '" target="_blank" rel="noopener">' + esc(c.vendor + ' · ' + c.name) + '</a><br><small>' + esc(c.evidence + ' · ' + c.status) + '</small>').join('<br>') : '검증 가능한 후보 없음 · RFQ']);
      }
      if ($('cap-candidates')) $('cap-candidates').innerHTML = table(['구간', '요구 규격', '카탈로그 후보 / 근거'], rows) + '<p class="muted">케이블·커넥터·패널은 정확한 길이·극성·핀·용량 정보 미확정 시 규격만 표시합니다.</p>' + catalogs.filter(x => x.error).map(x => '<p class="error">' + esc(x.path + ': ' + x.error) + '</p>').join('');
      latest.catalogCandidates = evidenceRows;
    } catch (error) {if (token === renderToken && $('cap-candidates')) $('cap-candidates').textContent = '카탈로그 조회 실패: ' + error.message;}
  }
  function download(name, content, type) {const url = URL.createObjectURL(new Blob([content], {type})), anchor = document.createElement('a'); anchor.href = url; anchor.download = name; anchor.click(); setTimeout(() => URL.revokeObjectURL(url), 1000);}
  function exportRows() {return selected().bom.map(x => [selected().name, names[x.network], x.phase, x.segment, x.item, x.media, x.spec, x.installedQty, x.spareQty, x.purchaseQty, x.low, x.high, x.uncertaintyPct, x.unit, x.referenceOnly ? '설계 참고' : '구매 계획', x.note]);}
  const bomHeaders = ['설계안', '네트워크', 'Phase', '구간', '품목', '매체', '규격', '설치', '예비품', '구매 기준', '하한', '상한', '±%', '단위', '범위', '비고'];
  function exportCsv() {
    if (!latest) return;
    const rows = [bomHeaders, ...exportRows(), [], ['중간 계산', '값'], ...Object.entries(latest.summary), [], ['대안', '케이블 조수 기준', '광모듈 기준', '패널 기준', '트레이 m 기준', '비용'], ...latest.alternatives.map(a => [a.name, a.totals.cables, a.totals.transceivers, a.totals.panels, a.totals.trayM, '산정 불가']), [], ['Phase', '블록', '랙', '서버', 'GPU', 'IT kW', '링크'], ...latest.phaseRows.map(p => [p.phase, p.blocks, p.racks, p.servers, p.gpus, p.itKw, p.links]), [], ['사용 가정', '값', '단위', '입력 출처', '근거', '신뢰도', '상태'], ...latest.assumptions.map(x => [x.path, typeof x.value === 'object' ? JSON.stringify(x.value) : x.value, x.unit, x.origin, x.source, x.confidence, x.status]), [], ['검토 사항'], ...latest.warnings.map(x => [x]), [], ['다음 확인 질문'], ...latest.questions.map(x => [x])];
    const csv = rows.map(r => r.map(x => {const text = String(x ?? ''); return '"' + (/^[=+@-]/.test(text) && typeof x !== 'number' ? "'" : '') + text.replace(/"/g, '""') + '"';}).join(',')).join('\r\n');
    download('DC-Capacity-BOM-v8-' + alternativeId + '.csv', '\uFEFF' + csv, 'text/csv;charset=utf-8');
  }
  async function exportXlsx() {
    if (!latest) return;
    try {
      const wb = new ExcelJS.Workbook(); wb.creator = 'DataCenter Capacity Designer';
      const add = (name, headers, rows) => {const ws = wb.addWorksheet(name); ws.addRow(headers); ws.addRows(rows); ws.views = [{state: 'frozen', ySplit: 1}]; ws.getRow(1).font = {bold: true}; ws.columns = headers.map(() => ({width: 24})); ws.eachRow(row => row.eachCell(cell => {cell.alignment = {wrapText: true, vertical: 'top'};}));};
      add('1_Requirements', ['입력', '값'], Object.entries(latest.input).map(([k, v]) => [k, typeof v === 'object' ? JSON.stringify(v) : v ?? '프리셋 기본값']));
      add('2_Calculation', ['중간 결과', '값'], Object.entries(latest.summary));
      add('3_BOM', bomHeaders, exportRows());
      add('4_Assumptions', ['계수', '값', '단위', '입력 출처', '근거', '공식 URL', '신뢰도', '상태'], latest.assumptions.map(x => [x.path, typeof x.value === 'object' ? JSON.stringify(x.value) : x.value, x.unit, x.origin, x.source, x.sourceUrl || '', x.confidence, x.status]));
      add('5_Alternatives', ['설계안', '케이블 조수', '광모듈', '패널', '트레이 m', '장점', '검토 사항', '비용'], latest.alternatives.map(x => [x.name, x.totals.cables, x.totals.transceivers, x.totals.panels, x.totals.trayM, x.pros, x.cons, '산정 불가']));
      add('6_Phases', ['Phase', '블록', '랙', '서버', 'GPU', 'IT kW', '링크', '구매 케이블'], latest.phaseRows.map(x => [x.phase, x.blocks, x.racks, x.servers, x.gpus, x.itKw, x.links, x.cables]));
      add('7_Blocks', ['블록', '건물', '홀', 'Phase', '랙', '서버', 'GPU', 'IT kW'], latest.blocks.map(x => [x.id, x.building, x.hall, x.phase, x.racks, x.servers, x.gpus, x.itKw]));
      add('8_DCI', ['건물', 'Phase', '링크', '사용심선', '설치심선', '케이블 조수', 'WDM 채널', '토폴로지'], latest.dci.rows.map(x => [x.building, x.phase, x.links, x.activeFibers, x.installedFibers, x.cableCount, x.channels, x.topology]));
      add('9_References', ['품목', '수량', '단위', '비고'], latest.references.map(x => [x.item, x.qty, x.unit, x.note]));
      add('10_Review', ['종류', '내용'], [...latest.warnings.map(x => ['검토 사항', x]), ...latest.questions.map(x => ['다음 확인 질문', x]), ...Object.entries(latest.validation).map(([k, v]) => ['계산 보존 검증', k + ': ' + v])]);
      add('11_Catalog_Candidates', ['구간', '요구 규격', '벤더', '후보', '근거', 'URL', '상태'], (latest.catalogCandidates || []).flatMap(x => x.candidates.length ? x.candidates.map(c => [x.segment, x.spec, c.vendor, c.name, c.evidence, c.source, c.status]) : [[x.segment, x.spec, '', '', '', '', '검증 가능한 후보 없음 / RFQ']]));
      for (const a of latest.alternatives) add('BOM_' + a.id, bomHeaders, a.bom.map(x => [a.name, names[x.network], x.phase, x.segment, x.item, x.media, x.spec, x.installedQty, x.spareQty, x.purchaseQty, x.low, x.high, x.uncertaintyPct, x.unit, x.referenceOnly ? '설계 참고' : '구매 계획', x.note]));
      download('DC-Capacity-BOM-v8-' + alternativeId + '.xlsx', await wb.xlsx.writeBuffer(), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    } catch (error) {$('cap-error').textContent = 'Excel 생성 실패: ' + error.message;}
  }
  async function init() {
    try {
      const response = await fetch(new URL('capacity-design.config.json?v=8.0.0', location.href), {cache: 'no-store'});
      if (!response.ok) throw Error('설정 HTTP ' + response.status);
      config = await response.json();
      let saved; try {saved = JSON.parse(localStorage.getItem(key) || 'null');} catch {}
      if (saved?.config?.version === config.version) config = saved.config;
      buildInput();
      document.addEventListener('click', event => {
        if (!active) return;
        if (event.target.closest('#exportWorkbook')) {event.preventDefault(); event.stopImmediatePropagation(); exportXlsx();}
        if (event.target.closest('#run')) {event.preventDefault(); event.stopImmediatePropagation(); recompute();}
      }, true);
      document.addEventListener('dc:design', () => {if (active && latest) requestAnimationFrame(render);});
      document.addEventListener('change', event => {if (active && event.target.closest('#capacity-input')) requestAnimationFrame(() => {document.body.classList.remove('design-stale'); $('exportWorkbook').disabled = !latest;});});
      if (saved) {overrides = saved.overrides || {}; customPresets = saved.customPresets || {}; phaseDefinitions = saved.phases || []; restoreInput(saved.input || {}); renderPhases(); renderCoefficients(); if (saved.active) {activate(true); recompute();}}
      window.DCCapacityUI = {activate, recompute, exportXlsx, exportCsv, getConfig: () => config};
    } catch (error) {const form = document.querySelector('.formPanel'); if (form) form.insertAdjacentHTML('afterbegin', '<p class="error">용량 설계 초기화 실패: ' + esc(error.message) + '</p>');}
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, {once: true}); else init();
})();
