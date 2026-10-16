'use strict';
const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const source=name=>fs.readFileSync(path.join(__dirname,name),'utf8');

test('landing displays an explicitly labelled example before results arrive',()=>{
 const html=source('index.html');
 assert.match(html,/id="entry-sample-result"/);
 assert.match(html,/샘플 · 실시간 결과 아님/);
 assert.match(html,/1,937\.4 kW/);
 assert.match(html,/2,324\.88 kW/);
 assert.equal((html.match(/id="entry-sample-result"/g)||[]).length,1);
});
test('default/preset input mount preserves core choices and the colocation branch',()=>{
 const ui=source('design-entry.js');
 assert.match(ui,/\['targetGPU','systemId','topology'\]/);
 assert.match(ui,/1k GPU · B200 Rail/);
 assert.match(ui,/256 GPU · H200 Dual/);
 assert.match(ui,/4랙 · GB200 NVL72/);
 assert.match(ui,/const colo=form\.querySelector\(':scope>details\.coloInputs'\)/);
 assert.match(ui,/d!==advanced&&d!==colo/);
 assert.match(ui,/document\.addEventListener\('dc:inputs-ready',mount/);
});
test('stored preview is hidden only upon successful usable calculation',()=>{
 const ui=source('design-entry.js'),css=source('design-entry.css');
 assert.match(ui,/preview\.hidden=Boolean\(c\.valid\)/);
 assert.match(ui,/현재 입력값의 결과가 아니며/);
 assert.match(css,/#entry-sample-result\[hidden\]\{display:none!important\}/);
 assert.match(css,/grid-template-columns:repeat\(2,minmax\(0,1fr\)\)/);
});
test('original engine loads before the non-mutating UX layer',()=>{
 const html=source('index.html');
 assert(html.indexOf('engineering-v82-ui.js')<html.indexOf('design-entry.js'));
 assert(html.indexOf('design-checks.js')<html.indexOf('design-entry.js'));
 assert.match(source('design-checks.js'),/\['power','전력'/);
 assert.match(source('design-checks.js'),/\['cooling','냉각'/);
 assert.match(source('design-checks.js'),/\['ru','랙 공간'/);
});
