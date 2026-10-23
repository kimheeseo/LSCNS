// LS3D STEP20 release documentation + protected baseline validation.
// Run from repository root:
// node --test DCI/DataCenter/ls3d-three/step20.release.test.mjs
import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync,existsSync} from 'node:fs';
import {createHash} from 'node:crypto';
import {resolve} from 'node:path';

const root=process.cwd();
const prefix='DCI/DataCenter/';
const load=relative=>readFileSync(resolve(root,relative),'utf8');
const html=load(prefix+'LS_Datacenter_3D.html');
const rootReadme=load('README.md');
const guide=load(prefix+'README.md');
const version=load(prefix+'LS3D_VERSION_UPGRADE_HISTORY.txt');
const report=load(prefix+'architecture/STEP20_RELEASE_REPORT.md');

test('Three preview version appears in real HTML title and sidebar while core baseline remains declared',()=>{
  assert.match(html,/<title>LS Datacenter Campus · 3D Digital Twin v4\.8\.1<\/title>/);
  assert.match(html,/v4\.7\.0 계산 코어 · Three\.js v4\.8\.1/);
  assert.match(html,/bootstrap\.mjs\?v=20261023-ls-rack-detail/);
  assert.match(html,/three@0\.186\.1\/build\/three\.module\.js/);
});
test('root README and DataCenter guide link to the correct Three and rollback URLs',()=>{
  for(const text of [rootReadme,guide]){
    assert.match(text,/LS_Datacenter_3D\.html\?engine=three/);
    assert.match(text,/LS_Datacenter_3D\.html\?engine=legacy/);
    assert.match(text,/LS3D_VERSION_UPGRADE_HISTORY\.txt/);
    assert.match(text,/STEP20_RELEASE_REPORT\.md/);
    assert.match(text,/v4\.8\.1/);
  }
});
test('separate UTF-8 TXT describes all 20 stages, version strategy and unverified QA',()=>{
  for(let n=1;n<=20;n++){
    const marker='STEP '+String(n).padStart(2,'0');
    assert.ok(version.includes(marker),'Missing '+marker+' in upgrade TXT');
  }
  assert.match(version,/v4\.7\.0/);
  assert.match(version,/v4\.8\.1/);
  assert.match(version,/캐시/);
  assert.match(version,/롤백/);
  assert.match(report,/Chromium\/WebGL2/);
  assert.match(report,/실측 성능/);
});
test('protected 2D Gold Pixel Tour remains unchanged and BOM entrypoints remain wired',()=>{
  const file='LS_Datacenter_Campus.html';
  const bytes=readFileSync(resolve(root,prefix,file));
  const header=Buffer.from('blob '+bytes.length+'\0');
  const sha=createHash('sha1').update(header).update(bytes).digest('hex');
  assert.equal(sha,'9c84f24f2101043d7c113ce2d0da50e7245bc318',
    'Protected original changed: '+file);
  const bom=load(prefix+'index.html');
  for(const marker of ['id="trunkFiberCount"','engineering-v82.js?v=',
    'engineering-v82-ui.js?v=','id="run"'])
    assert.ok(bom.includes(marker),'Missing BOM entrypoint: '+marker);
});
test('roadmap, progress log and browser QA workflow all exist',()=>{
  for(const name of [
    prefix+'LS3D_V4_20_TASK_ROADMAP.md',
    prefix+'LS3D_PROGRESS_STATUS.md',
    prefix+'ls3d-three/step19.integration.playwright.spec.mjs',
    '.github/workflows/ls3d-step19-quality.yml'
  ])assert.equal(existsSync(resolve(root,name)),true,name);
});
