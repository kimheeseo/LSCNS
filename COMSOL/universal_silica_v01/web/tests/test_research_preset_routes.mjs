import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
const base=new URL('../',import.meta.url);
const html=readFileSync(new URL('index.html',base),'utf8');
const app=readFileSync(new URL('app.mjs',base),'utf8');
const viewer=readFileSync(new URL('vector_mode_ui.mjs',base),'utf8');
const hcf=readFileSync(new URL('hcf_designer_ui.mjs',base),'utf8');
const fem=readFileSync(new URL('hcf_complex_fem_ui.mjs',base),'utf8');
for(const text of ['value="hcf"','value="vector"','id="hcf-sidebar-controls"','id="vector-sidebar-controls"','id="standard-fiber-options"','data-hcf-target="hcf-coreDiameterUm"'])
 assert.ok(html.includes(text),'Missing HTML research control '+text);
assert.ok(html.includes('body.hcf-selection main.workspace >'),'HCF-specific workspace CSS');
assert.ok(html.includes('body.vector-selection main.workspace >'),'Vector-specific workspace CSS');
assert.doesNotMatch(html,/@media\(max-width:850px\)\{@media\(max-width:520px\)\{/,'Dangling media rules');
for(const text of ["if(name==='hcf'||name==='vector')","target.dispatchEvent(new Event('input'","syncHcfSidebar()","if(el.closest('.special-design-controls'))continue"])
 assert.ok(app.includes(text),'Missing routing handler '+text);
assert.ok(viewer.includes('LSCNS_vector_mode_v1'),'Vector schema must be checked');
assert.ok(viewer.includes("id=\"fv-json-file\""),'Import vector results JSON');
assert.ok(viewer.includes("id=\"fv-neff\""),'Present vector neff');
assert.ok(hcf.includes("id=\"hcf-coreDiameterUm\""),'HCF field exists');
assert.ok(fem.includes('Scalar P1 FEM'),'HCF scalar solver limitations');
console.log('PASS: research preset HCF + vector routing and result import contracts');
