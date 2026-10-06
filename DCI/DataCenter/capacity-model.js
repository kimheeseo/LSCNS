(function (root) {
  'use strict';
  const clone = x => JSON.parse(JSON.stringify(x));
  const ceil = x => Math.ceil(x - 1e-10);
  const sum = xs => xs.reduce((a, b) => a + b, 0);
  const blockedKeys = new Set(['__proto__', 'constructor', 'prototype']);
  function setPath(object, path, value) {
    const keys = path.split('.');
    if (keys.some(k => blockedKeys.has(k))) throw Error('허용되지 않는 설정 경로');
    let node = object;
    for (const key of keys.slice(0, -1)) {
      if (!Object.hasOwn(node, key)) throw Error('알 수 없는 설정: ' + path);
      node = node[key];
    }
    const key = keys.at(-1), previous = node[key];
    if (!previous || !Object.hasOwn(previous, 'value')) throw Error('계수 경로가 아닙니다: ' + path);
    node[key] = {...previous, value, origin: '사용자 수정', source: '사용자 입력; 기존 근거: ' + previous.source, confidence: '낮음', status: '검증 필요'};
  }
  function coefficientRows(config, prefix = '') {
    const rows = [];
    for (const [key, value] of Object.entries(config)) {
      const path = prefix ? prefix + '.' + key : key;
      if (value && typeof value === 'object' && !Array.isArray(value)) {
        if (Object.hasOwn(value, 'value')) rows.push({path, ...value});
        else rows.push(...coefficientRows(value, path));
      }
    }
    return rows;
  }
  function numberToken(value) {
    const match = String(value).replace(/,/g, '').match(/^\s*(\d+(?:\.\d+)?)\s*([만천]?)\s*$/);
    if (!match) throw Error('규모 숫자를 확인하세요.');
    return Number(match[1]) * ({'만': 10000, '천': 1000}[match[2]] || 1);
  }
  function parseScale(text, kind = 'auto', units) {
    const raw = String(text).replace(/,/g, '').trim();
    const token = '(\\d+(?:\\.\\d+)?\\s*[만천]?)';
    const rackKw = raw.match(/(?:랙당|rack\s*(?:power|당)|per\s*rack)\s*(\d+(?:\.\d+)?)\s*kW/i);
    if (kind !== 'auto') {
      const n = raw.match(new RegExp(token));
      if (!n) throw Error('규모 숫자를 입력하세요.');
      return {kind: units.powerToKw[kind] ? 'power' : kind, value: numberToken(n[1]), unit: units.powerToKw[kind] ? kind : kind, rackKw: rackKw ? Number(rackKw[1]) : undefined};
    }
    const rack = raw.match(new RegExp('(?:랙|racks?)\\s*' + token, 'i')) || raw.match(new RegExp(token + '\\s*(?:개\\s*)?(?:랙|racks?)', 'i'));
    const gpu = raw.match(new RegExp('GPU(?:s)?\\s*' + token, 'i')) || raw.match(new RegExp(token + '\\s*(?:개의?\\s*)?GPU', 'i'));
    const server = raw.match(new RegExp('(?:서버|servers?)\\s*' + token, 'i')) || raw.match(new RegExp(token + '\\s*(?:대\\s*)?(?:서버|servers?)', 'i'));
    const power = raw.match(/(\d+(?:\.\d+)?)\s*(GW|MW|kW)/i);
    let result;
    if (rack) result = {kind: 'racks', value: numberToken(rack[1]), unit: 'racks'};
    else if (gpu) result = {kind: 'gpu', value: numberToken(gpu[1]), unit: 'gpu'};
    else if (server) result = {kind: 'servers', value: numberToken(server[1]), unit: 'servers'};
    else if (power) result = {kind: 'power', value: Number(power[1]), unit: Object.keys(units.powerToKw).find(k => k.toLowerCase() === power[2].toLowerCase())};
    else throw Error('15GW, GPU 10,000개, 랙 2,000개, 서버 500대처럼 입력하세요.');
    return {...result, rackKw: rackKw ? Number(rackKw[1]) : undefined, detectedEquipment: /GB200\s*NVL72/i.test(raw) ? 'nvl72' : undefined};
  }
  function distribute(total, fractions) {
    const entries = Object.entries(fractions);
    if (entries.some(([, x]) => !Number.isFinite(x) || x < 0) || Math.abs(sum(entries.map(([, x]) => x)) - 1) > 1e-9) throw Error('거리 분포는 음수가 아닌 값이며 합계가 1이어야 합니다.');
    const result = entries.map(([key, f]) => ({key, qty: Math.floor(total * f), fraction: total * f % 1}));
    let remain = total - sum(result.map(x => x.qty));
    const ranked = [...result].sort((a, b) => b.fraction - a.fraction);
    for (let i = 0; i < remain; i++) ranked[i % ranked.length].qty++;
    return result;
  }
  function calculate(input, original) {
    const config = clone(original);
    if (input.customPresets) {
      for (const [key, preset] of Object.entries(input.customPresets)) {
        if (blockedKeys.has(key)) throw Error('잘못된 프리셋 이름');
        config.equipmentPresets[key] = clone(preset);
      }
    }
    for (const [path, value] of Object.entries(input.overrides || {})) setPath(config, path, value);
    if (!Array.isArray(config.mediaRules?.profiles) || !config.mediaRules.profiles.length) throw Error('매체 규격 프로파일이 필요합니다.');
    for (const p of config.mediaRules.profiles) {
      if (!p.id || !p.source || !p.confidence || !p.status || !Array.isArray(p.speeds) || p.speeds.some(x => !(x > 0)) || !(p.reachM > 0) || ![p.activeFibers, p.installedFibers, p.connectorGroups].every(x => Number.isSafeInteger(x) && x >= 0) || p.installedFibers < p.activeFibers) throw Error('매체 규격·심수·출처를 확인하세요: ' + p.id);
    }
    const assumptions = new Map(), warnings = new Set();
    const get = path => {
      let node = config;
      for (const key of path.split('.')) {
        if (blockedKeys.has(key) || !Object.hasOwn(node, key)) throw Error('설정 누락: ' + path);
        node = node[key];
      }
      if (!node || !Object.hasOwn(node, 'value') || !node.source || !node.confidence || !node.status) throw Error('출처·신뢰도·상태가 없는 계수: ' + path);
      assumptions.set(path, {path, ...node, origin: node.origin || '프리셋 기본값'});
      if (typeof node.value === 'number' && !Number.isFinite(node.value)) throw Error('유한한 계수 필요: ' + path);
      return node.value;
    };
    const numeric = (path, min = 0, integer = false, max = Infinity) => {
      const value = get(path);
      if (typeof value !== 'number' || value < min || value > max || (integer && !Number.isInteger(value))) throw Error('계수 범위 오류: ' + path);
      return value;
    };
    const c = key => numeric('coefficients.' + key);
    const l = key => numeric('layout.' + key, 0.0000001, /Per/.test(key));
    const workload = input.workload || config.defaults.workload, wPath = 'workloadPresets.' + workload;
    if (!config.workloadPresets[workload]) throw Error('알 수 없는 용도 프리셋');
    const scale = parseScale(input.scaleText, input.scaleKind || 'auto', config.units);
    if (!(scale.value > 0) || !Number.isFinite(scale.value)) throw Error('규모는 0보다 커야 합니다.');
    if (scale.kind !== 'power' && !Number.isInteger(scale.value)) throw Error('GPU·랙·서버 수는 정수로 입력하세요.');
    const equipmentId = input.equipment || scale.detectedEquipment || config.workloadPresets[workload].equipment;
    const equipment = config.equipmentPresets[equipmentId];
    if (!equipment) throw Error('장비 프리셋을 찾을 수 없습니다.');
    const ePath = 'equipmentPresets.' + equipmentId;
    const e = key => get(ePath + '.' + key);
    let rackKw = input.rackKw ?? scale.rackKw ?? (equipment.rackScale ? e('rackKw') : get(wPath + '.rackKw'));
    let pue = input.pue ?? get(wPath + '.pue');
    if (!(rackKw > 0) || !Number.isFinite(rackKw) || !(pue >= 1) || !Number.isFinite(pue)) throw Error('랙 전력은 양수, PUE는 1 이상이어야 합니다.');
    for (const [key, value, path] of [['rackKw', rackKw, equipment.rackScale ? ePath + '.rackKw' : wPath + '.rackKw'], ['pue', pue, wPath + '.pue']]) {
      if (input[key] != null || (key === 'rackKw' && scale.rackKw != null)) assumptions.set(path, {path, value, unit: key === 'pue' ? 'ratio' : 'kW/rack', source: '고객 입력', confidence: '낮음', status: '검증 필요', origin: '고객 입력'});
    }
    const gpuPerServer = numeric(ePath + '.gpuPerServer', 0, true), serverKw = numeric(ePath + '.serverKw', 0.0000001), serverRu = numeric(ePath + '.serverRu', 0.0000001);
    const share = numeric('coefficients.computePowerShare', 0.0000001, false, 1);
    const rackRu = c('usableRackRu');
    if (rackRu > c('rackLimitRu')) throw Error('사용 가능 RU가 전체 랙 높이를 초과합니다.');
    let serversPerRack = Math.min(Math.floor(rackKw * share / serverKw + 1e-10), Math.floor(rackRu / serverRu));
    if (equipment.rackScale) {
      serversPerRack = numeric(ePath + '.serversPerRack', 1, true);
      if (serversPerRack * gpuPerServer !== numeric(ePath + '.gpuPerRack', 1, true)) throw Error('랙 스케일 GPU/랙과 tray 수가 일치하지 않습니다.');
      if (serversPerRack * serverKw > rackKw || serversPerRack * serverRu > rackRu) warnings.add('랙 스케일 tray 구성과 전력/RU 한도가 충돌합니다. OEM 전체 랙 구성을 확인하세요.');
    }
    if (serversPerRack < 1) throw Error('현재 장비를 한 대도 수용할 수 없는 랙 전력/RU입니다. 장비 또는 랙 가정을 수정하세요.');
    let racks, servers, itKw, requestedGpu = null;
    if (scale.kind === 'power') {
      const kw = scale.value * config.units.powerToKw[scale.unit];
      if (!['it', 'facility'].includes(input.powerBasis || config.defaults.powerBasis)) throw Error('전력 기준을 확인하세요.');
      itKw = (input.powerBasis || config.defaults.powerBasis) === 'facility' ? kw / pue : kw;
      racks = ceil(itKw / rackKw);
      servers = equipment.rackScale ? Math.floor(itKw / rackKw + 1e-10) * serversPerRack : Math.min(racks * serversPerRack, Math.floor(itKw * share / serverKw + 1e-10));
    } else if (scale.kind === 'racks') {
      racks = scale.value; servers = racks * serversPerRack; itKw = racks * rackKw;
    } else {
      if (scale.kind === 'gpu' && gpuPerServer === 0) throw Error('GPU 입력에는 GPU 장비 프리셋을 선택하세요.');
      requestedGpu = scale.kind === 'gpu' ? scale.value : null;
      servers = scale.kind === 'gpu' ? ceil(scale.value / gpuPerServer) : scale.value;
      racks = ceil(servers / serversPerRack);
      if (equipment.rackScale) servers = racks * serversPerRack;
      itKw = racks * rackKw;
    }
    const gpus = servers * gpuPerServer, facilityKw = itKw * pue;
    if (![racks, servers, gpus].every(Number.isSafeInteger)) throw Error('안전한 정수 범위를 초과한 규모입니다.');
    if (servers === 0) warnings.add('입력 전력으로 완전한 서버/랙 스케일 장비를 수용할 수 없습니다. 네트워크 BOM은 0이며 장비·규모 확인이 필요합니다.');
    const rememberInput = (path, value, unit) => assumptions.set(path, {path, value, unit, source: '고객 입력', confidence: '낮음', status: '검증 필요', origin: '고객 입력'});
    const cooling = input.cooling || (equipment.cooling ? e('cooling') : get(wPath + '.cooling'));
    if (!['air', 'liquid'].includes(cooling)) throw Error('냉각 방식은 air 또는 liquid입니다.');
    if (input.cooling) rememberInput(equipment.cooling ? ePath + '.cooling' : wPath + '.cooling', cooling, 'mode');
    if (cooling === 'air' && rackKw > c('airCoolingLimitKw')) warnings.add('공랭 선택과 랙 밀도가 계획 경고 임계값을 초과합니다. 냉각 방식·공급 온도·풍량 검증 필요.');
    const tier = input.tier || config.defaults.tier;
    if (tier !== 'unspecified') warnings.add(tier + '는 요구 등급입니다. 네트워크 이중화 선택으로 Tier 인증 충족을 판정하지 않습니다.');
    if (input.powerLimitMw != null && (!Number.isFinite(input.powerLimitMw) || input.powerLimitMw <= 0)) throw Error('전력 인입 한도는 양수여야 합니다.');
    if (input.powerLimitMw != null && facilityKw > input.powerLimitMw * config.units.powerToKw.MW) warnings.add('시설 전체 전력이 인입 한도를 ' + ((facilityKw / config.units.powerToKw.MW) - input.powerLimitMw).toFixed(3) + ' MW 초과합니다.');
    const siteAreaM2 = racks * c('rackAreaM2') * c('siteAreaMultiplier');
    if (input.siteAreaM2 != null && (!Number.isFinite(input.siteAreaM2) || input.siteAreaM2 <= 0)) throw Error('부지 면적은 양수여야 합니다.');
    if (input.siteAreaM2 != null && siteAreaM2 > input.siteAreaM2) warnings.add('계획 면적이 입력 부지를 초과합니다. 면적 계수는 임의 가정이며 배치 검증 필요.');
    const computeBasis = e('computePortBasis');
    if (!['gpu', 'server'].includes(computeBasis)) throw Error('포트 기준은 gpu 또는 server여야 합니다.');
    const perServerPorts = {compute: numeric(ePath + '.computePorts', 0, true) * (computeBasis === 'gpu' ? gpuPerServer : 1), storage: numeric(ePath + '.storagePorts', 0, true), management: numeric(ePath + '.managementPorts', 0, true)};
    const settings = {};
    for (const key of Object.keys(perServerPorts)) {
      const p = 'networks.' + key;
      const enabled = get(p + '.enabled');
      if (typeof enabled !== 'boolean') throw Error('네트워크 사용 여부는 true/false여야 합니다.');
      const redundancy = input.redundancy?.[key] ?? numeric(p + '.redundancy', 1, true, 2);
      if (![1, 2].includes(redundancy)) throw Error('네트워크 이중화는 1 또는 2여야 합니다.');
      const topology = key === 'compute' ? (input.topology || get(wPath + '.topology')) : 'two';
      const speed = key === 'compute' ? (input.speed ?? get(wPath + '.speed')) : numeric(p + '.speed', 0.0000001);
      const oversub = key === 'compute' ? get(wPath + '.oversub') : numeric(p + '.oversub', 1);
      if (!['two', 'three'].includes(topology) || speed <= 0 || oversub < 1) throw Error('토폴로지·속도·오버서브 계수를 확인하세요.');
      settings[key] = {key, name: config.networks[key].name, enabled, redundancy, topology, speed, oversub, protocol: key === 'compute' ? (input.protocol || get(wPath + '.protocol')) : get(p + '.protocol'), leafDown: numeric(p + '.leafDown', 1, true), leafUp: numeric(p + '.leafUp', 1, true), spinePorts: numeric(p + '.spinePorts', 1, true), spineUp: key === 'compute' && topology === 'three' ? numeric(p + '.spineUp', 1, true) : 0};
      if (key === 'compute') {
        if (input.speed != null) rememberInput(wPath + '.speed', speed, 'Gbps');
        if (input.protocol) rememberInput(wPath + '.protocol', input.protocol, 'protocol');
        if (input.topology) rememberInput(wPath + '.topology', topology, 'mode');
      }
      if (settings[key].spinePorts <= settings[key].spineUp) throw Error('Spine 다운링크 포트가 없습니다.');
      if (ceil(settings[key].leafDown / oversub) > settings[key].leafUp) throw Error('Leaf 업링크 포트가 오버서브 목표를 수용하지 못합니다.');
      if (input.redundancy?.[key] != null) assumptions.set(p + '.redundancy', {path: p + '.redundancy', value: redundancy, unit: 'fabrics', source: '고객 입력', confidence: '낮음', status: '검증 필요', origin: '고객 입력'});
    }
    let racksPerBlock = l('racksPerBlock');
    for (const n of Object.values(settings)) if (n.enabled && perServerPorts[n.key]) {
      const bound = Math.floor(n.leafDown * (n.spinePorts - n.spineUp) / (serversPerRack * perServerPorts[n.key]));
      if (bound < 1) throw Error('한 랙의 포트 수가 블록 패브릭 한도를 초과합니다. 스위치 포트 또는 장비 구성을 수정하세요.');
      racksPerBlock = Math.min(racksPerBlock, bound);
    }
    if (racksPerBlock !== l('racksPerBlock')) warnings.add('Spine 포트 한도에 맞춰 실제 블록 크기를 ' + racksPerBlock + '랙으로 줄였습니다.');
    const blockCount = ceil(racks / racksPerBlock), blocksPerHall = l('blocksPerHall'), hallsPerBuilding = l('hallsPerBuilding');
    if (blockCount > 100000) throw Error('100,000블록 이상은 현재 브라우저 계산 범위 밖입니다. 블록 계수를 조정하세요.');
    let buildings = input.buildings ?? ceil(blockCount / (blocksPerHall * hallsPerBuilding));
    if (!Number.isSafeInteger(buildings) || buildings < 1 || buildings > Math.max(1, blockCount)) throw Error('건물 수는 1 이상, 블록 수 이하여야 합니다.');
    const blocksPerBuilding = ceil(blockCount / buildings);
    rememberInput('layout.buildingCount', buildings, 'buildings');
    if (input.buildings == null) assumptions.set('layout.buildingCount', {path: 'layout.buildingCount', value: buildings, unit: 'buildings', source: '블록·홀·건물 용량 기본값으로 자동 분할', confidence: '낮음', status: '임의 가정', origin: '자동 환산'});
    if (blocksPerBuilding > blocksPerHall * hallsPerBuilding) warnings.add('지정 건물 수에 필요한 홀 수가 기본 홀/건물 가정을 초과합니다. 건물·홀 면적 검증 필요.');
    const phases = input.phases || [];
    const assigned = new Map();
    for (const phase of phases) {
      if (![phase.start, phase.end, phase.phase].every(Number.isSafeInteger) || phase.start < 1 || phase.end < phase.start || phase.end > blockCount || phase.phase < 1) throw Error('Phase는 유효한 블록 시작·끝 번호와 양의 정수 단계로 지정하세요.');
      for (let b = phase.start; b <= phase.end; b++) {
        if (assigned.has(b)) throw Error('Phase 블록 범위가 중복됩니다.');
        assigned.set(b, phase.phase);
      }
    }
    const blocks = [], fractions = get('layout.endpointDistribution');
    let serversRemaining = servers;
    for (let i = 0; i < blockCount; i++) {
      const count = Math.min(racksPerBlock, racks - i * racksPerBlock);
      const countServers = Math.min(serversRemaining, count * serversPerRack);
      serversRemaining -= countServers;
      const building = Math.floor(i / blocksPerBuilding) + 1, local = i % blocksPerBuilding;
      blocks.push({id: i + 1, building, hall: Math.floor(local / blocksPerHall) + 1, phase: assigned.get(i + 1) || config.defaults.phase, racks: count, servers: countServers, gpus: countServers * gpuPerServer, itKw: scale.kind === 'power' ? Math.min(count * rackKw, itKw - i * racksPerBlock * rackKw) : count * rackKw});
    }
    const networkRows = [], segmentMap = new Map(), coreByBuilding = new Map();
    function segment(network, block, tierName, distanceClass, distanceM, links, speed, protocol) {
      if (!links) return;
      const id = [network, block.phase, tierName, distanceClass, distanceM, speed].join('|');
      const existing = segmentMap.get(id);
      if (existing) existing.links += links;
      else segmentMap.set(id, {network, phase: block.phase, tier: tierName, distanceClass, distanceM, links, speed, protocol});
    }
    for (const block of blocks) for (const n of Object.values(settings)) {
      if (!n.enabled) continue;
      const endpoints = block.servers * perServerPorts[n.key];
      if (!endpoints) continue;
      const leaves = ceil(endpoints / n.leafDown), full = Math.floor(endpoints / n.leafDown), partial = endpoints % n.leafDown;
      const fullUp = ceil(n.leafDown / n.oversub), partialUp = partial ? ceil(partial / n.oversub) : 0;
      const fabricLinks = full * fullUp + partialUp;
      const spineDown = n.spinePorts - n.spineUp;
      const effectiveSpineDown = n.topology === 'three' ? Math.min(spineDown, n.spineUp * numeric('networks.compute.coreOversub', 1)) : spineDown;
      const spineCount = Math.max(Math.min(numeric('coefficients.minSpines', 1, true), fabricLinks), ceil(fabricLinks / effectiveSpineDown));
      if (ceil(fabricLinks / spineCount) > spineDown) throw Error('Spine 포트 보존 실패');
      const replicas = n.redundancy;
      const row = {block: block.id, building: block.building, phase: block.phase, network: n.key, endpoints: endpoints * replicas, leaf: leaves * replicas, spine: spineCount * replicas, fabricLinks: fabricLinks * replicas, coreLinks: 0, speed: n.speed, protocol: n.protocol, leafPortCapacity: leaves * (n.leafDown + n.leafUp) * replicas, leafPortsUsed: (endpoints + fabricLinks) * replicas, spineDownCapacity: spineCount * spineDown * replicas};
      networkRows.push(row);
      for (const part of distribute(row.endpoints, fractions)) segment(n.key, block, 'endpoint', part.key, part.key === 'rack' ? l('rackDistanceM') : part.key === 'row' ? l('rowLengthM') : l('hallDistanceM'), part.qty, n.speed, n.protocol);
      segment(n.key, block, 'leaf-spine', 'hall', l('fabricDistanceM'), row.fabricLinks, n.speed, n.protocol);
      if (n.topology === 'three') {
        const coreRatio = numeric('networks.compute.coreOversub', 1), available = numeric('networks.compute.superSpinePorts', 1, true);
        const fullSpineDown = Math.floor(fabricLinks / spineCount), bigger = fabricLinks % spineCount;
        const upSmall = ceil(fullSpineDown / coreRatio), upBig = ceil((fullSpineDown + 1) / coreRatio);
        if (Math.max(upSmall, bigger ? upBig : 0) > n.spineUp) throw Error('Spine 예약 업링크로 블록 간 오버서브 목표를 수용할 수 없습니다.');
        row.coreLinks = ((spineCount - bigger) * upSmall + bigger * upBig) * replicas;
        segment(n.key, block, 'spine-super-spine', 'hall', l('superSpineDistanceM'), row.coreLinks, n.speed, n.protocol);
        const key = block.building + '|' + block.phase;
        const previous = coreByBuilding.get(key) || {building: block.building, phase: block.phase, links: 0, ports: available, redundancy: replicas};
        previous.links += row.coreLinks; coreByBuilding.set(key, previous);
      }
    }
    const coreRows = [...coreByBuilding.values()].map(x => ({...x, count: ceil(x.links / x.redundancy / x.ports) * x.redundancy}));
    if (coreRows.some(x => x.count / x.redundancy > settings.compute.spineUp)) warnings.add('Super-spine은 포트 수 기준 집계 수량입니다. 다수 그룹의 완전 연결·rail 배치·블록 간 도달성은 상세 설계 검증 필요.');
    const dci = {links: 0, activeFibers: 0, installedFibers: 0, cableCount: 0, rows: [], demandGbps: 0};
    const dciReplicas = input.redundancy?.dci ?? numeric('networks.dci.redundancy', 1, true, 2);
    if (![1, 2].includes(dciReplicas)) throw Error('DCI 이중화는 1 또는 2여야 합니다.');
    if (input.redundancy?.dci != null) assumptions.set('networks.dci.redundancy', {path: 'networks.dci.redundancy', value: dciReplicas, unit: 'routes', source: '고객 입력', confidence: '낮음', status: '검증 필요', origin: '고객 입력'});
    const dciSpeed = numeric('networks.dci.speed', 0.0000001), dciTopology = get('networks.dci.topology');
    if (!['star', 'ring'].includes(dciTopology)) throw Error('DCI 토폴로지는 star 또는 ring입니다.');
    const channels = numeric('networks.dci.wdmChannels', 1, true), cableFibers = numeric('networks.dci.cableFibers', 2, true);
    const reserve = numeric('networks.dci.fiberReservePct', 0, false, 99) / 100;
    const gbpsPerRack = numeric('networks.dci.gbpsPerRack', 0);
    if (buildings > 1) {
      const grouped = new Map();
      for (const block of blocks) {
        const key = block.building + '|' + block.phase;
        const g = grouped.get(key) || {building: block.building, phase: block.phase, racks: 0};
        g.racks += block.racks; grouped.set(key, g);
      }
      for (const group of grouped.values()) {
        if (dciTopology === 'star' && group.building === 1) continue;
        if (dciTopology === 'ring' && buildings === 2 && group.building === 1) continue;
        const links = ceil(group.racks * gbpsPerRack / dciSpeed) * dciReplicas;
        const pairs = ceil(links / dciReplicas / channels) * dciReplicas;
        const dciLength = l('buildingDistanceM') * (1 + c('slackPct') / 100) + c('endsPerLink') * c('endAllowanceM');
        const dciProfile = config.mediaRules.profiles.find(p => p.media === 'SMF' && p.speeds.includes(dciSpeed) && p.reachM >= dciLength);
        const activeFibers = channels > 1 ? pairs * c('endsPerLink') : links * (dciProfile?.activeFibers || c('endsPerLink'));
        if (!dciProfile) warnings.add('DCI 심수는 임시 duplex 가정이며 reach/모듈 미확정입니다. 해당 구간의 실제 심수 검증 필요.');
        const cableCount = ceil((activeFibers / dciReplicas) / (cableFibers * (1 - reserve))) * dciReplicas;
        const row = {...group, links, activeFibers, cableCount, installedFibers: cableCount * cableFibers, channels, demandGbps: group.racks * gbpsPerRack, topology: dciTopology};
        dci.rows.push(row);
        for (const key of ['links', 'activeFibers', 'cableCount', 'installedFibers', 'demandGbps']) dci[key] += row[key];
        segment('dci', group, 'dci', 'building', l('buildingDistanceM'), links, dciSpeed, 'Ethernet');
      }
      warnings.add('DCI 대역폭은 랙당 수요 가정으로 산출합니다. 건물 간 학습 패브릭 전체 대역폭을 보장하지 않습니다. 트래픽·gateway·물리 경로 분리 검증 필요.');
      if (dciTopology === 'ring') warnings.add('Ring은 건물별 회선 수요의 계획 합계입니다. 양방향 트래픽·통과 트래픽·장애 시 수용량은 별도 검증 필요.');
    }
    if (phases.length) warnings.add('Phase별 코어·DCI는 단계마다 별도 설치하는 보수적 수량입니다. 기존 공용 장비 재사용 및 선행 공사 일정은 확인 필요.');
    if (settings.compute.topology === 'two' && blockCount > 1) warnings.add('2계층안은 블록 내부 연결 수량입니다. 블록 간 aggregation/gateway 장비와 uplink는 별도 RFQ입니다.');
    warnings.add('스토리지 서버/head 및 스위치 관리 포트는 OEM·스토리지 규모 미정으로 별도 RFQ입니다. 현재 endpoint는 계산된 서버 기준입니다.');
    warnings.add('토폴로지는 논리 포트·대역폭 계획입니다. cage/breakout, rail 매핑, 손실 예산과 실제 배선은 검증 필요.');
    const segments = [...segmentMap.values()];
    const effectiveLength = d => d * (1 + c('slackPct') / 100) + c('endsPerLink') * c('endAllowanceM');
    const sparePct = numeric('coefficients.sparePct', 0), uncertaintyPct = numeric('coefficients.uncertaintyPct', 0, false, 99);
    function mediaFor(s, policy) {
      const distance = effectiveLength(s.distanceM);
      const profiles = config.mediaRules.profiles;
      const use = p => {assumptions.set('mediaRules.profiles.' + p.id, {path: 'mediaRules.profiles.' + p.id, value: {speed: s.speed, reachM: p.reachM, activeFibers: p.activeFibers, installedFibers: p.installedFibers, connectorGroups: p.connectorGroups, connector: p.connector}, unit: 'profile', source: p.source, sourceUrl: p.sourceUrl, confidence: p.confidence, status: p.status, origin: '프리셋 기본값', readonly: true}); return {...p, lengthM: distance};};
      if (s.network !== 'dci' && policy !== 'SMF') {
        const base = profiles.find(p => p.id === 'baseT' && p.speeds.includes(s.speed) && p.reachM >= distance && distance <= numeric('mediaRules.baseTMaxM', 0));
        if (base) return use(base);
        const speeds = get('mediaRules.assemblySpeeds');
        if (speeds.includes(s.speed)) {
          const dac = numeric('mediaRules.dacMaxM', 0), aoc = numeric('mediaRules.aocMaxM', 0);
          const medium = ['auto', 'DAC'].includes(policy) && distance <= dac ? 'DAC' : ['auto', 'AOC'].includes(policy) && distance <= aoc ? 'AOC' : null;
          if (medium) return {media: medium, fiber: medium === 'DAC' ? 'copper' : 'integrated', connector: '장비 cage 미확정', standard: s.speed + 'G ' + medium + ' 계획', activeFibers: 0, installedFibers: 0, connectorGroups: 0, lengthM: distance, package: '미확정', status: '검증 필요'};
        }
      }
      const preferred = s.network === 'dci' ? 'SMF' : ['MMF', 'SMF'].includes(policy) ? policy : 'MMF';
      const profile = profiles.find(p => p.media === preferred && p.speeds.includes(s.speed) && p.reachM >= distance) || profiles.find(p => p.media === 'SMF' && p.speeds.includes(s.speed) && p.reachM >= distance);
      if (!profile) {warnings.add(s.speed + 'G / ' + distance.toFixed(1) + 'm 링크를 지원하는 검증 가능한 계획 프로파일이 없습니다. 해당 구간은 RFQ로 유지합니다.'); return {media: 'RFQ', fiber: '미확정', connector: '미확정', standard: s.speed + 'G reach 검증 필요', lengthM: distance, activeFibers: 0, installedFibers: 0, connectorGroups: 0, status: '검증 필요'};}
      if (preferred !== profile.media) warnings.add(preferred + ' 선호 매체가 거리/속도 조건을 만족하지 않아 ' + profile.media + ' 계획안을 사용합니다.');
      if (s.protocol === 'InfiniBand') warnings.add('Ethernet 광 규격으로 표시된 항목은 InfiniBand 물리 매체 계획 참고입니다. IB 지원·FEC·모듈 호환 검증 전 제품 확정 불가.');
      return use(profile);
    }
    function makeAlternative(definition) {
      const bom = [], add = (s, item, media, spec, installedQty, unit, note, extra = {}) => {
        if (!installedQty) return;
        const spareQty = unit === 'm' ? installedQty * sparePct / 100 : ceil(installedQty * sparePct / 100);
        const purchaseQty = installedQty + spareQty;
        bom.push({id: bom.length + 1, network: s.network, phase: s.phase, segment: s.tier + ' / ' + s.distanceClass, item, media, spec, installedQty, spareQty, purchaseQty, low: unit === 'm' ? purchaseQty * (1 - uncertaintyPct / 100) : Math.floor(purchaseQty * (1 - uncertaintyPct / 100)), high: unit === 'm' ? purchaseQty * (1 + uncertaintyPct / 100) : ceil(purchaseQty * (1 + uncertaintyPct / 100)), uncertaintyPct, unit, note, ...extra});
      };
      const selections = [], cpoEndpoints = new Map();
      for (const s of segments) {
        const policy = definition.id === 'base' ? (input.media || config.defaults.media) : definition.policy;
        if (!['auto', 'DAC', 'AOC', 'MMF', 'SMF'].includes(policy)) throw Error('매체 선택을 확인하세요.');
        const originalProfile = mediaFor(s, policy);
        const p = s.network === 'dci' && channels > 1 ? {...originalProfile, standard: 'DWDM 채널 광모듈 RFQ', package: '미확정', status: '검증 필요'} : originalProfile;
        const spec = s.speed + 'G · ' + p.standard + ' · ' + p.fiber + ' · ' + p.connector + ' · ' + p.lengthM.toFixed(1) + 'm';
        if (['DAC', 'AOC'].includes(policy) && p.media !== policy && s.network !== 'dci' && p.media !== 'BASE-T') warnings.add(policy + ' 선호가 유효 길이/속도 조건을 만족하지 않아 ' + p.media + '로 대체된 구간이 있습니다.');
        selections.push({...s, ...p});
        if (p.media === 'RFQ') {add(s, '연결 매체 RFQ', 'RFQ', spec, s.links, '링크', '거리·속도 호환 프로파일 미확정; BOM 미완성'); continue;}
        const optical = ['MMF', 'SMF'].includes(p.media);
        if (!optical) {add(s, '연결 케이블', p.media, spec, s.links, '개', p.media === 'AOC' ? '광송수신부 일체형; 별도 플러그형 트랜시버 0개' : '양단 일체형 또는 RJ45 패치 케이블; 광모듈 0개', {lengthM: p.lengthM, cable: true, requirement: {...p, speed: s.speed, protocol: s.protocol}}); continue;}
        if (s.network !== 'dci') add(s, '광 케이블 assembly', p.media, spec, s.links, '개', '양단 패널 접속 계획; 양단 패치리드 별도', {lengthM: p.lengthM, cable: true, requirement: {...p, speed: s.speed, protocol: s.protocol}});
        const useCpo = definition.id === 'cpo' && s.network === 'compute';
        const pluggables = useCpo ? (s.tier === 'endpoint' ? s.links : 0) : c('endsPerLink') * s.links;
        add(s, '플러그형 트랜시버', p.media, s.speed + 'G · ' + p.standard + ' · ' + p.package + ' · ' + p.connector, pluggables, '개', useCpo ? '서버 측 pluggable만 포함; 스위치 측 CPO 별도' : '링크 양단 모듈', {requirement: {...p, speed: s.speed, protocol: s.protocol}, transceiver: true});
        if (useCpo) {
          const key = s.phase + '|' + s.speed;
          const group = cpoEndpoints.get(key) || {...s, ports: 0}; group.ports += s.links * (s.tier === 'endpoint' ? 1 : c('endsPerLink')); cpoEndpoints.set(key, group);
        }
        const groups = c('endsPerLink') * s.links * p.connectorGroups;
        add(s, '커넥터 종단 그룹', p.media, p.connector + ' · 1그룹은 LC duplex 또는 MPO 종단 1개', groups, '그룹', 'assembly에 포함된 종단 참고; 별도 구매 수량으로 중복 합산 금지', {referenceOnly: true});
        add(s, '패치패널', p.media, p.connector + ' · ' + c('panelTerminations') + ' 종단그룹/패널', c('endsPerLink') * ceil(s.links * p.connectorGroups / c('panelTerminations')), '개', '구간 양단 용량 합산 최소 계획; 랙·패널별 단편화 및 adapter SKU 확인 필요');
        if (s.network !== 'dci') add(s, '패치리드', p.media, p.connector + ' · ' + p.fiber + ' · 양단 작업 여유 ' + c('endAllowanceM') + 'm 참고', c('endsPerLink') * s.links * p.connectorGroups, '개', '서버/스위치 ↔ 패널; 단면 수량 포함');
      }
      for (const s of cpoEndpoints.values()) {
        const engines = ceil(s.ports / numeric('mediaRules.cpoPortsPerEngine', 1, true));
        add(s, 'CPO optical engine 참고', 'CPO', s.speed + 'G · ports/engine 설정 기반', engines, '개', '패키지별 포트 단편화 미반영; 실제 engine 수량 검증 필요', {referenceOnly: true});
        add(s, 'CPO ELS 참고', 'CPO', '외부 레이저 계획', engines * numeric('mediaRules.cpoExternalLasersPerEngine', 0, true), '개', 'ELS redundancy/fiber attach 확정 필요', {referenceOnly: true});
      }
      for (const row of dci.rows) {
        const s = {network: 'dci', phase: row.phase, tier: 'dci', distanceClass: 'building'};
        add(s, 'DCI SMF 트렁크', 'SMF', cableFibers + 'F · OS2 · ' + effectiveLength(l('buildingDistanceM')).toFixed(1) + 'm · ' + row.activeFibers + ' 사용심선 / ' + row.installedFibers + ' 설치심선', row.cableCount, '조', '경로별 여유심선 ' + reserve * 100 + '%; 건물 ' + row.building, {lengthM: effectiveLength(l('buildingDistanceM')), cable: true});
        if (channels > 1) add(s, 'WDM terminal/mux 참고', 'SMF', channels + ' channels/fiber-pair', row.activeFibers, '단', '양단 단말 기준; 광파워·파장계획 RFQ', {referenceOnly: true});
      }
      const byPhase = new Map();
      for (const b of blocks) {const p = byPhase.get(b.phase) || {racks: 0, phase: b.phase}; p.racks += b.racks; byPhase.set(b.phase, p);}
      for (const group of byPhase.values()) {
        const cableLength = sum(bom.filter(x => x.phase === group.phase && x.cable).map(x => x.installedQty * x.lengthM));
        const routes = ceil(group.racks / numeric('coefficients.racksPerTrayRoute', 1, true));
        const baseRouteM = routes * l('rowLengthM');
        const averageCables = baseRouteM ? cableLength / baseRouteM : 0;
        const layers = Math.max(1, ceil(averageCables * Math.PI * (c('cableDiameterMm') / 2) ** 2 / (numeric('coefficients.trayAreaMm2', 0.0000001) * numeric('coefficients.trayFill', 0.0000001, false, 1))));
        add({network: 'shared', phase: group.phase, tier: 'route', distanceClass: 'row/hall'}, '트레이/덕트 계획 길이', '경로', '공유 경로 · 평균 점유 기반 ' + layers + ' 병렬단면', baseRouteM * layers, 'm', '평균 OD·경로로 산정; 이중 경로 분리·굽힘·국소 점유 검증 필요');
      }
      // CPU가 서버 내부에 포함되더라도, BOM/견적 검토에서 누락되지 않도록
      // 별도 계획 행으로 표시합니다. 정확한 CPU SKU·소켓·TDP는 OEM 확정 전 RFQ입니다.
      const cpuSpec = equipmentId === 'genericCpu' ? 'CPU-only 서버 · CPU SKU 미확정' : '호스트 CPU · OEM/SKU 미확정';
      add({network: 'compute', phase: 1, tier: 'compute', distanceClass: 'server'}, 'CPU / Host Processor', 'Compute', cpuSpec, servers, '개', '서버당 1개 계획 수량. 실제 소켓 수·CPU 모델·TDP·가격은 서버 OEM/견적서로 확정 필요', {computeComponent: 'cpu', referenceOnly: true});
      const qty = item => sum(bom.filter(x => x.item === item).map(x => x.purchaseQty));
      return {...definition, bom, selections, totals: {cables: sum(bom.filter(x => x.cable).map(x => x.purchaseQty)), cableM: sum(bom.filter(x => x.cable).map(x => x.purchaseQty * x.lengthM)), transceivers: qty('플러그형 트랜시버'), panels: qty('패치패널'), connectorGroups: qty('커넥터 종단 그룹'), trayM: qty('트레이/덕트 계획 길이'), rfqLinks: qty('연결 매체 RFQ')}, priceStatus: '산정 불가'};
    }
    const alternatives = config.alternatives.filter(x => x.id !== 'cpo' || input.cpo).map(makeAlternative);
    const references = [{item: equipment.rackScale ? 'Compute tray' : '서버', qty: servers, unit: '대', note: equipment.name}, {item: 'IT 랙', qty: racks, unit: '개', note: 'IT 전체 부하 기준'}, {item: 'Compute NIC 계획', qty: settings.compute.enabled ? ceil(perServerPorts.compute / numeric('coefficients.computePortsPerNic', 1, true)) * servers * settings.compute.redundancy : 0, unit: '개', note: '포트/NIC 설정 기준; 기존 온보드 포트·분기·이중 패브릭 실제 NIC 구성 확인 필요'}, {item: 'Leaf', qty: sum(networkRows.map(x => x.leaf)), unit: '대', note: '네트워크별 독립 패브릭 합계'}, {item: 'Spine', qty: sum(networkRows.map(x => x.spine)), unit: '대', note: '네트워크별 독립 패브릭 합계'}, {item: 'Super-spine', qty: sum(coreRows.map(x => x.count)), unit: '대', note: '포트 집계 계획; 완전 연결 검증 필요'}, {item: '시설 전력 모듈', qty: ceil(facilityKw / (numeric('coefficients.powerModuleMw', 0.0000001) * config.units.powerToKw.MW)), unit: '모듈', note: '시설 전력 기준 참고; UPS/변압기/발전기 개별 수량 아님'}, {item: '냉각 모듈', qty: ceil(itKw / numeric('coefficients.coolingModuleKw', 0.0000001)), unit: '모듈', note: 'IT 열부하 기준 참고; 실제 CDU/CRAH 용량 검증 필요'}];
    const phaseRows = [...new Set(blocks.map(x => x.phase))].sort((a, b) => a - b).map(phase => {const xs = blocks.filter(x => x.phase === phase), bom = alternatives[0].bom.filter(x => x.phase === phase);return {phase, blocks: xs.length, racks: sum(xs.map(x => x.racks)), servers: sum(xs.map(x => x.servers)), gpus: sum(xs.map(x => x.gpus)), itKw: sum(xs.map(x => x.itKw)), links: sum(segments.filter(x => x.phase === phase).map(x => x.links)), cables: sum(bom.filter(x => x.cable).map(x => x.purchaseQty))};});
    const validation = {racks: sum(blocks.map(x => x.racks)) === racks, servers: sum(blocks.map(x => x.servers)) === servers, gpu: sum(blocks.map(x => x.gpus)) === gpus, power: Math.abs(sum(blocks.map(x => x.itKw)) - itKw) < 1e-6, endpointLinks: sum(segments.filter(x => x.tier === 'endpoint').map(x => x.links)) === sum(networkRows.map(x => x.endpoints)), fabricLinks: sum(segments.filter(x => x.tier === 'leaf-spine').map(x => x.links)) === sum(networkRows.map(x => x.fabricLinks)), leafPorts: networkRows.every(x => x.leafPortsUsed <= x.leafPortCapacity), spinePorts: networkRows.every(x => x.fabricLinks <= x.spineDownCapacity), phaseLinks: sum(phaseRows.map(x => x.links)) === sum(segments.map(x => x.links))};
    if (Object.values(validation).some(x => !x)) throw Error('내부 수량 보존 검증에 실패했습니다.');
    const questions = ['랙당 IT 부하 ' + rackKw + ' kW와 ' + equipment.name + '의 실제 전력·RU·외부 포트 구성이 확정됐나요?', phaseRows.length > 1 ? '각 Phase의 블록 가동 시점과 기존 건물/DCI 공용 경로 재사용 계획은 무엇인가요?' : '건물·홀 배치와 단계별 증설 블록 범위를 확정할 수 있나요?', '선호 벤더와 cage/FEC·커넥터 극성·실제 배선 길이·광손실 예산을 확인할 수 있나요?'];
    return {version: config.version, input: clone(input), scale, workload, equipmentId, summary: {itKw, facilityKw, pue, rackKw, racks, servers, serversPerRack, gpus, requestedGpu, gpuExcess: requestedGpu == null ? null : gpus - requestedGpu, computeKw: servers * serverKw, otherItKw: Math.max(0, itKw - servers * serverKw), rackCapacityKw: racks * rackKw, siteAreaM2, cooling, buildings, halls: sum(Array.from({length: buildings}, (_, i) => ceil(blocks.filter(x => x.building === i + 1).length / blocksPerHall))), blocks: blockCount, racksPerBlock, endpointPorts: sum(networkRows.map(x => x.endpoints)), links: sum(segments.map(x => x.links))}, assumptions: [...assumptions.values()], blocks, networks: networkRows, coreRows, segments, dci, phaseRows, alternatives, references, questions, warnings: [...warnings], validation, internalReference: equipment.internalReference || 'OEM 내부 연결 구성 미확정; 외부 네트워크 BOM에서 제외', tier, budget: input.budget || '', priceStatus: '산정 불가'};
  }
  function matchCatalog(requirement, catalog) {
    if (!requirement || requirement.protocol === 'InfiniBand') return [];
    const connector = x => String(x).toUpperCase().replace(/\s+/g, '');
    return Object.entries(catalog.products || {}).flatMap(([key, meta]) => {
      const s = {...catalog.defaultSpecs, ...meta.specs};
      const rate = String(s['Data Rate'] || '').match(/^(\d+(?:\.\d+)?)\s*G/i);
      const reach = String(s.Reach || '').match(/(\d+(?:\.\d+)?)\s*(km|m)/i);
      const remark = String(s.Remark || '');
      const standard = requirement.standard.replace(/BASE-/i, '').replace(/\s+/g, '');
      const canonicalRemark = remark.replace(/BASE-/i, '').replace(/\s+/g, '');
      const uri = meta.officialUrl || catalog.officialUrl;
      if (!rate || Number(rate[1]) !== requirement.speed || !reach || !uri || !/^https:\/\//.test(uri)) return [];
      if (Number(reach[1]) * (reach[2].toLowerCase() === 'km' ? 1000 : 1) < requirement.lengthM || connector(s.Connector) !== connector(requirement.connector) || s.Package !== requirement.package || !canonicalRemark.toUpperCase().includes(standard.toUpperCase())) return [];
      return [{key, name: meta.name || key, vendor: catalog.company, source: uri, evidence: '카탈로그 속도·규격·패키지·커넥터·reach 일치', status: '규격 후보; host/FEC/광손실·극성 호환 검증 필요'}];
    });
  }
  const api = {calculate, parseScale, coefficientRows, distribute, matchCatalog};
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.DCCapacity = api;
})(typeof window !== 'undefined' ? window : globalThis);
