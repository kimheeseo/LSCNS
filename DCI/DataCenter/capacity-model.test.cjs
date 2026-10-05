'use strict';
const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const M = require('./capacity-model.js');
const config = require('./capacity-design.config.json');
const run = input => M.calculate(input, config);
const total = (r, field, network) => r.networks.filter(x => !network || x.network === network).reduce((n, x) => n + x[field], 0);
const scenarios = [
  {name: 'A: 10,000 GPUs', input: {scaleText: 'GPU 10,000개'}, expected: {itKw: 18780, facilityKw: 24414, racks: 313, servers: 1250, gpus: 10000, endpointPorts: 13750, links: 27817, blocks: 10, buildings: 1}},
  {name: 'B: 2,000 racks at 10 kW', input: {scaleText: '랙 2,000개, 랙당 10kW', workload: 'colo'}, expected: {itKw: 20000, facilityKw: 28000, racks: 2000, servers: 32000, gpus: 0, endpointPorts: 128000, links: 168025, blocks: 63, buildings: 2}},
  {name: 'C: 100 MW, one building', input: {scaleText: '100MW', buildings: 1}, expected: {itKw: 100000, facilityKw: 130000, racks: 1667, servers: 6666, gpus: 53328, endpointPorts: 73326, links: 148320, blocks: 53, buildings: 1}}
];
for (const scenario of scenarios) test(scenario.name, () => {
  const r = run(scenario.input);
  for (const [key, expected] of Object.entries(scenario.expected)) assert.equal(r.summary[key], expected, key);
  assert.ok(Object.values(r.validation).every(Boolean));
  assert.equal(r.priceStatus, '산정 불가');
  for (const a of r.alternatives) for (const row of a.bom) {assert.ok(Number.isFinite(row.purchaseQty) && row.purchaseQty >= 0); assert.ok(row.low <= row.purchaseQty && row.high >= row.purchaseQty);}
  if (scenario.name.startsWith('A')) {
    // Independent block arithmetic: 9 full blocks and one 98-server block.
    assert.equal(total(r, 'fabricLinks', 'compute'), 9 * 1024 + 784);
    assert.equal(total(r, 'fabricLinks', 'storage'), 9 * 128 + 98);
    assert.equal(total(r, 'fabricLinks', 'management'), 9 * 32 + 25);
    assert.equal(total(r, 'coreLinks', 'compute'), 9 * 256 + 200);
  }
  if (scenario.name.startsWith('B')) {assert.equal(r.dci.links, Math.ceil(976 * 10 / 400)); assert.equal(r.dci.activeFibers, 50); assert.equal(r.dci.installedFibers, 144);}
  if (scenario.name.startsWith('C')) {assert.equal(r.summary.halls, 14); assert.ok(r.warnings.some(x => x.includes('홀 수'))); assert.equal(total(r, 'coreLinks', 'compute'), 52 * 256 + 3 * 7);}
});
test('IT/facility conversion and inlet limit', () => {
  const a = run({scaleText: '100MW'}), b = run({scaleText: '130MW', powerBasis: 'facility'});
  assert.equal(b.summary.itKw, a.summary.itKw); assert.equal(b.summary.racks, a.summary.racks);
  assert.ok(run({scaleText: '100MW', powerLimitMw: 120}).warnings.some(x => x.includes('인입 한도')));
  assert.ok(!run({scaleText: '100MW', powerLimitMw: 130}).warnings.some(x => x.includes('인입 한도')));
});
test('independent redundancy doubles only the selected fabric', () => {
  const a = run({scaleText: 'GPU 10000'}), b = run({scaleText: 'GPU 10000', redundancy: {compute: 2}});
  for (const f of ['endpoints', 'leaf', 'spine', 'fabricLinks', 'coreLinks']) assert.equal(total(b, f, 'compute'), total(a, f, 'compute') * 2);
  for (const k of ['storage', 'management']) assert.equal(total(b, 'endpoints', k), total(a, 'endpoints', k));
  const c = run({scaleText: '랙 2000개', workload: 'colo', redundancy: {dci: 2}});
  const d = run({scaleText: '랙 2000개', workload: 'colo'});
  assert.equal(c.dci.links, d.dci.links * 2); assert.equal(c.dci.cableCount, d.dci.cableCount * 2);
});
test('phases conserve scale, links, ports and partial blocks', () => {
  const a = run({scaleText: 'GPU 10000'}), b = run({scaleText: 'GPU 10000', phases: [{start: 6, end: 10, phase: 2}]});
  assert.equal(b.phaseRows.length, 2); assert.equal(b.summary.links, a.summary.links);
  assert.equal(b.phaseRows.reduce((n, x) => n + x.gpus, 0), 10000);
  assert.throws(() => run({scaleText: 'GPU 10000', phases: [{start: 1, end: 2, phase: 1}, {start: 2, end: 3, phase: 2}]}), /중복/);
});
test('media, integrated optics and CPO avoid duplicate pluggables', () => {
  const r = run({scaleText: 'GPU 10000', cpo: true});
  for (const a of r.alternatives) {
    for (const s of a.selections.filter(x => ['DAC', 'AOC', 'BASE-T'].includes(x.media))) assert.equal(a.bom.filter(x => x.transceiver && x.network === s.network && x.segment === s.tier + ' / ' + s.distanceClass).length, 0);
  }
  const cpo = r.alternatives.find(x => x.id === 'cpo'), smf = r.alternatives.find(x => x.id === 'smf');
  assert.ok(cpo.totals.transceivers < smf.totals.transceivers);
  assert.ok(cpo.bom.some(x => x.item === 'CPO optical engine 참고' && x.referenceOnly));
});
test('overrides are attributed and JSON-only custom presets work', () => {
  const custom = JSON.parse(JSON.stringify(config.equipmentPresets.generic8)); custom.name = 'Custom'; custom.gpuPerServer.value = 4;
  const a = run({scaleText: 'GPU 10000', overrides: {'equipmentPresets.generic8.computePorts': 2}});
  assert.equal(total(a, 'endpoints', 'compute'), 20000);
  assert.equal(a.assumptions.find(x => x.path === 'equipmentPresets.generic8.computePorts').origin, '사용자 수정');
  const b = run({scaleText: 'GPU 10000', equipment: 'custom4', customPresets: {custom4: custom}});
  assert.equal(b.summary.servers, 2500); assert.equal(b.summary.gpus, 10000);
});
test('rack scale uses whole racks and excludes internal NVLink from BOM', () => {
  const r = run({scaleText: 'GB200 NVL72 랙 500개'});
  assert.equal(r.summary.gpus, 500 * 72); assert.equal(r.summary.racks, 500);
  assert.ok(r.internalReference.includes('NVLink'));
  assert.ok(!r.alternatives[0].bom.some(x => /NVLink/.test(x.item)));
  assert.equal(run({scaleText: 'GPU 73개', equipment: 'nvl72'}).summary.gpus, 144);
});
test('warnings and invalid configuration do not produce false results', () => {
  assert.ok(run({scaleText: '랙 10개', rackKw: 200, cooling: 'air'}).warnings.some(x => x.includes('공랭')));
  assert.throws(() => run({scaleText: '100MW', pue: 0.9}), /PUE/);
  assert.throws(() => run({scaleText: 'GPU -5개'}));
  assert.throws(() => run({scaleText: 'GPU 10000', overrides: {'layout.endpointDistribution': {rack: 0.5, hall: 0.7}}}), /분포/);
  assert.throws(() => run({scaleText: 'GPU 10000', overrides: {'__proto__.polluted': 1}}), /허용/);
  const broken = JSON.parse(JSON.stringify(config)); broken.mediaRules.profiles[1].activeFibers = -8;
  assert.throws(() => M.calculate({scaleText: '100MW'}, broken), /매체 규격/);
});
test('unknown reach remains RFQ; WDM physical fibers differ from channels', () => {
  const r = run({scaleText: '랙 2000개', workload: 'colo', overrides: {'layout.buildingDistanceM': 100000}});
  assert.ok(r.alternatives[0].bom.some(x => x.media === 'RFQ'));
  const a = run({scaleText: '랙 2000개', workload: 'colo'}), b = run({scaleText: '랙 2000개', workload: 'colo', overrides: {'networks.dci.wdmChannels': 8}});
  assert.equal(a.dci.links, b.dci.links); assert.ok(b.dci.activeFibers < a.dci.activeFibers);
});
test('catalog candidates require all published dimensions and protocol fit', () => {
  const cat = JSON.parse(fs.readFileSync(path.join(__dirname, 'product_catalog/YOFC/Transceiver/400G/catalog.json'), 'utf8'));
  const req = {protocol: 'Ethernet', speed: 400, connector: 'LC', package: 'QSFP-DD', standard: '400GBASE-FR4', lengthM: 1102};
  assert.ok(M.matchCatalog(req, cat).length > 0);
  assert.equal(M.matchCatalog({...req, lengthM: 3000}, cat).length, 0);
  assert.equal(M.matchCatalog({...req, connector: 'MPO-16'}, cat).length, 0);
  assert.equal(M.matchCatalog({...req, protocol: 'InfiniBand'}, cat).length, 0);
  assert.equal(M.matchCatalog({...req, speed: 800}, cat).length, 0);
});
test('15 GW uses bounded block aggregation', () => {
  const t = performance.now(), r = run({scaleText: '15GW'});
  assert.equal(r.summary.itKw, 15000000); assert.equal(r.summary.facilityKw, 19500000);
  assert.equal(r.summary.racks, 250000); assert.equal(r.summary.blocks, 7813);
  assert.ok(r.segments.length < 100); assert.ok(performance.now() - t < 5000);
});
