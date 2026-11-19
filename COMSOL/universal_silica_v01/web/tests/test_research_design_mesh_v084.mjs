import test from 'node:test';
import assert from 'node:assert/strict';
import {validateMeshPreview,inputsMatch} from '../research_design_mesh_v084.mjs';

const preview=()=>({schema:'LSCNS_mesh_preview_v084',input:{mesh:{symmetry:'full'}},
  quality:{triangles:1},triangles:[{vertices_um:[[0,0],[1,0],[0,1]],material:'core'}],
  boundary_edges:{OUTER_PEC:[[[0,0],[1,0]]]}});
test('mesh import accepts a linear triangle but rejects nonfinite or absent boundaries',()=>{
  assert.equal(validateMeshPreview(preview()).quality.triangles,1);
  const p=preview();p.boundary_edges.OUTER_PEC[0][0][0]=Infinity;
  assert.throws(()=>validateMeshPreview(p));
  delete p.boundary_edges;assert.throws(()=>validateMeshPreview(p));
});
test('matching ignores JSON key order and detects changed geometry or mesh',()=>{
  assert.equal(inputsMatch({mesh:{bulk:1},a:2},{a:2,mesh:{bulk:1}}),true);
  assert.equal(inputsMatch({mesh:{bulk:1}},{mesh:{bulk:.5}}),false);
});
