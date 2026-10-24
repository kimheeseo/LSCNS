import test from 'node:test';import assert from 'node:assert/strict';
import {describeArchitecturePart} from './architecture.mjs';
test('gold power tray is identified as cable support, not busway',()=>{
 const p=describeArchitecturePart('tray','hall-power-tray-rail--0.38','hall');
 assert.equal(p.label,'전력 케이블 트레이');assert.equal(p.component,'측면 레일');assert.match(p.description,/버스웨이는 아닙니다/);
 assert.equal(p.id,describeArchitecturePart('tray','hall-power-tray-rung-4','hall').id);
});
test('fiber tray, hanger and structural beam have distinct labels',()=>{
 assert.equal(describeArchitecturePart('tray','hall-fiber-tray-base','hall').label,'광케이블 트레이');
 assert.equal(describeArchitecturePart('tray','hall-power-tray-hanger-left-2','hall').component,'천장 행거');
 assert.equal(describeArchitecturePart('beam','cross-beam-x-1','hall').label,'천장 구조 보');
});
test('walls and floors do not become click blockers; zones have separate IDs',()=>{
 assert.equal(describeArchitecturePart('wall','wall','hall'),null);
 assert.notEqual(describeArchitecturePart('tray','power-tray-base','hall').id,describeArchitecturePart('tray','power-tray-base','network').id);
});

test('tray channel click surface shares rail identity and detailed product guidance',()=>{
 const surface=describeArchitecturePart('tray','hall-power-tray-pick-surface','hall');
 const rail=describeArchitecturePart('tray','hall-power-tray-rail--0.38','hall');
 assert.equal(surface.id,rail.id);assert.equal(surface.component,'트레이 본체');
 for(const word of ['Cable Tray','허용 하중','접지·본딩','천장 행거'])assert.ok(surface.description.includes(word));
});
