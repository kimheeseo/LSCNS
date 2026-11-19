import test from 'node:test';
import assert from 'node:assert/strict';
import {validateModeResult} from '../research_step3_v085.mjs';
import {inputsMatch} from '../research_design_mesh_v084.mjs';
const mode=()=>({neff_real:1.4,neff_imag:0,residual_relative:1e-12,weak_gauss_relative:1e-12,
 field_samples:{coordinates_um:[[0,0]],components:Object.fromEntries(['Ex','Ey','Ez','Hx','Hy','Hz'].map(k=>[k,[[1,0]]]))}});
const data=()=>({schema:'LSCNS_fiber_research_vector_v1',input:{wavelength_nm:1550},modes:[mode(),mode()]});
test('retains independent or degenerate mode records for selector',()=>assert.equal(validateModeResult(data()).modes.length,2));
test('rejects nonfinite H field in any returned mode',()=>{let d=data();d.modes[1].field_samples.components.Hz[0][1]=Infinity;assert.throws(()=>validateModeResult(d),/Hz/);});
test('rejects sample coordinate and field size mismatches',()=>{let d=data();d.modes[0].field_samples.components.Ez=[];assert.throws(()=>validateModeResult(d),/Ez/);});
test('result must carry the exact design; key order and integral formatting are irrelevant',()=>{assert(inputsMatch({b:1,a:2},{a:2.0,b:1.0}));assert(!inputsMatch(data().input,{wavelength_nm:1500}));let d=data();delete d.input;assert.throws(()=>validateModeResult(d));});
