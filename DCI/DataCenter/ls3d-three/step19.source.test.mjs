// STEP 19 source/compatibility regressions for all hosted modules and protected pages.
// node --test DCI/DataCenter/ls3d-three/step19.source.test.mjs
import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync,existsSync} from 'node:fs';
import {createHash} from 'node:crypto';
const root=new URL('../../../',import.meta.url);
const read=path=>readFileSync(new URL(path,root),'utf8');
const html=read('DCI/DataCenter/LS_Datacenter_3D.html');
const adapter=read('DCI/DataCenter/ls3d-three/scene-adapter.mjs');
const bootstrap=read('DCI/DataCenter/ls3d-three/bootstrap.mjs');
const specialized=read('DCI/DataCenter/ls3d-three/zone-specializations.mjs');
const mep=read('DCI/DataCenter/ls3d-three/campus-mep-routes.mjs');
const policy=read('DCI/DataCenter/ls3d-three/scene-viewport-policy.mjs');

test('seven canonical zones, physical asset anchors and pinned import map survive',()=>{
  for(const zone of ['campus','utility','cooling','gray','hall','network','ops'])
    assert.match(html,new RegExp("id:'"+zone+"'"));
  assert.ok(html.includes('gpu-${r}-${i}'),'GPU rack IDs are generated from row/column loops');
  for(const id of ['tor-a','carrier','mmr','odf','core',
    'xfmr-a','switchgear','ups-a','pdu','tower-a','chiller',
    'colo-rack-a','colo-rack-b','colo-rack-c'])
    assert.ok(html.includes("'"+id+"'"),id);
  assert.ok(html.includes('three@0.186.1/build/three.module.js'));
  assert.match(html,/bootstrap\.mjs\?v=[A-Za-z0-9-]+/,'Three bootstrap must be cache-versioned');
});

test('STEP03-18 Three scene adapter still wires specialized 3D modules and QA diagnostics',()=>{
  for(const imported of ['rack-network-cabling.mjs','mmr-room.mjs','colocation-cages.mjs',
    'campus-mep-routes.mjs','scene-viewport-policy.mjs']){
    assert.ok((adapter+specialized).includes(imported),imported);
  }
  for(const token of ['createCameraOcclusion','layoutSceneOverlays',
    'cameraOcclusion.dispose()','viewportPerformance','mepController','rackCabling'])
    assert.ok(adapter.includes(token),token);
  assert.ok(bootstrap.includes('lastPinch'));
  assert.ok(bootstrap.includes('multiTouch'));
  assert.ok(bootstrap.includes("window.LS3D_THREE"));
});

test('MMR/colo/MEP retain separate optical, power A/B and cooling route identities',()=>{
  assert.ok(specialized.includes("pickTenantCircuit"));
  assert.ok(specialized.includes("coloCircuit"));
  assert.ok(specialized.includes("mmrLink"));
  for(const id of ['carrier-to-tor','hall-to-customer-A','hall-to-customer-B',
    'hall-to-customer-C','coolant-to-gpu','power-to-colo-A'])
    assert.ok(mep.includes(id),id);
  assert.ok(mep.includes('verified:false'));
  assert.ok(mep.includes('lengthMeters:null'));
  assert.ok(policy.includes('sample(deltaSeconds)'));
});

test('all five preceding stage Node test files and STEP19 acceptance artifacts exist',()=>{
  for(const stage of [14,15,16,17,18])
    assert.ok(existsSync(new URL(
      'DCI/DataCenter/ls3d-three/step'+stage+'.unit.test.mjs',root)),stage);
  for(const file of [
    'DCI/DataCenter/ls3d-three/step19.integration.playwright.spec.mjs',
    'DCI/DataCenter/ls3d-three/step19.playwright.config.mjs',
    '.github/workflows/ls3d-step19-quality.yml'
  ])assert.ok(existsSync(new URL(file,root)),file);
});

function gitBlobSha(path){
  const data=readFileSync(new URL(path,root));
  return createHash('sha1').update(Buffer.from('blob '+data.length+'\0')).update(data).digest('hex');
}
test('previous 2D Gold Pixel Tour remains unchanged and BOM entrypoints remain wired',()=>{
  assert.equal(gitBlobSha('DCI/DataCenter/LS_Datacenter_Campus.html'),
    '9c84f24f2101043d7c113ce2d0da50e7245bc318');
  const bom=read('DCI/DataCenter/index.html');
  for(const marker of ['id="trunkFiberCount"','engineering-v82.js?v=','engineering-v82-ui.js?v=','id="run"'])
    assert.ok(bom.includes(marker),marker);
});
