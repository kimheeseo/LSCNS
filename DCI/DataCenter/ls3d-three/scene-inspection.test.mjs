import test from 'node:test';
import assert from 'node:assert/strict';
import {isVisiblePickObject,describeEquipment} from './scene-inspection.mjs';
test('hidden LOD and hidden parent geometry cannot intercept scene clicks',()=>{
 const root={visible:true},hidden={visible:false,parent:root};
 assert.equal(isVisiblePickObject({visible:true,parent:hidden,userData:{}}),false);
 assert.equal(isVisiblePickObject({visible:false,parent:root,userData:{}}),false);
 assert.equal(isVisiblePickObject({visible:true,parent:root,userData:{}}),true);
 assert.equal(isVisiblePickObject({visible:true,parent:root,userData:{noPick:true}}),false);
});
test('busduct has an independent product card and does not claim a manufacturer rating',()=>{
 const p=describeEquipment({id:'cable-tray',name:'Power Busway',zone:'gray'},'busway');
 assert.equal(p.id,'equipment:cable-tray');assert.equal(p.assetId,'cable-tray');
 assert.match(p.label,/부스덕트/);assert.match(p.description,/탭오프/);assert.match(p.description,/트레이/);assert.match(p.description,/특정 제품.*지정하지 않습니다/);
});
test('other equipment keeps its own description and racks keep existing part inspection',()=>{
 const a={id:'ups',name:'UPS',zone:'gray',desc:'무정전 전원 공급',details:'배터리 연계'};
 assert.match(describeEquipment(a,'ups').description,/무정전 전원 공급 배터리 연계/);
 assert.equal(describeEquipment(a,'rack'),null);
});
