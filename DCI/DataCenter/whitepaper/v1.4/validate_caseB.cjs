'use strict';
/**
 * Reproducible Case B / Phase 7-A cross-check for whitepaper v1.4.
 * Run from any working directory: node DCI/DataCenter/whitepaper/v1.4/validate_caseB.cjs
 * This code tests the shared BOM COUNT engine, not a live web UI or PSU/cooling approval.
 */
const assert = require('node:assert/strict');
const { calculate } = require('../../phase7-bom-engine.js');
const links = Array.from({length: 64}, (_, i) => ({
  id: `B-fiber-${String(i + 1).padStart(2, '0')}`,
  from: `GPU-${Math.floor(i / 8) + 1}`,
  to: 'SW-01',
  type: 'SMF',
  media: 'SMF',
  lengthM: 30,
  speed: '400G',
  down: false
}));
const packet = {
  schemaVersion: 'lsdc-twin-bom/1.0',
  facility: {profile: 'aidc', rackCount: 2, gpuServerRatio: 1},
  compute: {gpuServerCount: 8, gpuPerServer: 8},
  network: {nicPerServer: 1},
  optical: {cpoMode: 'pluggable', topology: {
    nodes: [{id: 'SW-01', layer: 'Leaf'}], links
  }}
};
const baseline = calculate(packet);
const c = baseline.counts;
let count = 0;
function check(actual, expected, name) {
  assert.deepEqual(actual, expected, name);
  count++;
}
check(baseline.schema, 'lsdc/phase7-bom-engine/1.0', 'engine schema');
check(c.rackCount, 2, 'racks (input value)');
check(c.gpuServerCount, 8, 'GPU servers (explicit input)');
check(c.gpuCount, 64, 'GPUs');
check(c.nicCount, 8, 'NICs: default 1/server, not physical port cages');
check(c.leafCount, 1, 'Leaf nodes: fixture assumption');
check(c.spineCount, 0, 'Spine nodes');
check(c.coreCount, 0, 'Core nodes');
check(c.opticalLinkCount, 64, 'active logical optical links');
check(c.totalOpticalLengthM, 1920, 'sum of 64 links x 30m');
check(baseline.rows.length, 9, 'summary row count');
check(baseline.rows.find(x => x.id === 'rack').quantity, 2, 'rack BOM row');
check(baseline.rows.find(x => x.id === 'gpu-server').quantity, 8, 'GPU-server BOM row');
check(baseline.rows.find(x => x.id === 'gpu').quantity, 64, 'GPU BOM row');
check(baseline.rows.find(x => x.id === 'optical-link').quantity, 64, 'optical-link BOM row');
check(baseline.rows.find(x => x.id === 'optical-length').quantity, 1920, 'optical length BOM row');
packet.optical.topology.links[0].down = true;
const afterDown = calculate(packet).counts;
check(afterDown.opticalLinkCount, 63, 'single link failure');
check(afterDown.totalOpticalLengthM, 1890, 'active length after failure');
check(Number((4 * 10.2 + 1).toFixed(1)), 41.8, 'independent 1-rack IT power, kW');
assert.throws(() => calculate({schemaVersion:'invalid'}), TypeError);
console.log(`PASS: ${count} value assertions + 1 invalid schema rejection`);
console.log('LIMITATION: no live GitHub Pages UI output and no validated power/cooling/SKU/connector inference.');
