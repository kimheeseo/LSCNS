/**
 * COMSOL public Step-Index Fiber (NOT a G.654.E product model) numeric benchmark.
 * The production web solver fixes the cladding index to silica Sellmeier.
 * This validation adapter substitutes COMSOL's constant n_clad=1.4378 ONLY
 * in validate() and solve(), leaving radial FEM matrices/integrals unchanged.
 * This permits an honest like-for-like neff comparison without changing the
 * existing user's production web preset.
 *
 * Official model: https://doc.comsol.com/6.4/doc/com.comsol.help.models.rf.step_index_fiber/step_index_fiber.html
 * a=8 um, R=40 um, n_core=1.4457, n_clad=1.4378, lambda=1550 nm.
 * COMSOL rounded neff 1.4444, V=4.895.
 * Independent analytical scalar LP01 Bessel root: neff=1.4444018444092133.
 */
import fs from 'node:fs';
import assert from 'node:assert/strict';
const sourceFile=new URL('../../universal_silica_v01/web/solver.mjs',import.meta.url);
const original=fs.readFileSync(sourceFile,'utf8');
const token='silicaIndex(m.wavelength_nm)';
const found=original.split(token).length-1;
assert.equal(found,2,'Only validate() and solve() may refer to the baseline index using this exact pattern');
const benchmarkSource=original.replaceAll(token,'1.4378');
const testSolver=await import('data:text/javascript;base64,'+Buffer.from(benchmarkSource).toString('base64'));
const p={...testSolver.PRESETS.g654,profile:'step',wavelength_nm:1550,core_radius_um:8,
  delta_n_core:1.4457-1.4378,outer_radius_um:40,elements:750};
const points=[200,350,500,750,1000,1400,1800].map(elements=>{
 const solution=testSolver.solve({...p,elements});
 assert.equal(solution.guided_proxy,true,'guided HE11-like step index scalar fundamental expected');
 assert.ok(solution.n_eff>1.4378&&solution.n_eff<1.4457);
 return {elements:solution.elements,neff:solution.n_eff,MFD_PII_um:solution.MFD_PII_um,
  Aeff_um2:solution.Aeff_um2,relative_residual:solution.residual};
});
const actual=points.find(p=>p.elements===750);
assert.ok(actual,'750-element baseline missing');
assert.ok(Math.abs(actual.neff-1.4444)<5e-5,'published COMSOL 4-decimal value must agree');
const exactLP01=1.4444018444092133;
const out={
 kind:'COMSOL_PUBLIC_STEP_INDEX_NOT_G654E',comsol_url:'https://doc.comsol.com/6.4/doc/com.comsol.help.models.rf.step_index_fiber/step_index_fiber.html',
 wavelength_nm:1550,core_radius_um:8,cladding_radius_um:40,n_core:1.4457,n_clad:1.4378,
 material_adapter:'ONLY two occurrences of silicaIndex(m.wavelength_nm) substituted with the COMSOL 1.4378 cladding constant: original radial FEM solver otherwise unchanged',
 comsol_neff_published_4_decimals:1.4444,analytical_scalar_LP01_neff:exactLP01,
 numerical_radial_FEM_neff:actual.neff,
 diff_from_COMSOL_reported_value:actual.neff-1.4444,
 diff_from_exact_scalar_LP01:actual.neff-exactLP01,
 mesh_study:points,
 limitations:[
   'Public COMSOL example is generic step-index fiber, not G.654.E or LE125',
   'COMSOL published neff is rounded to four decimals; cannot infer sub-1e-4 COMSOL accuracy',
   'Scalar radial FEM differs from COMSOL vector Maxwell FEM',
   'Material adapter only used for validation; production browser model is not modified',
   'No manufacturer RIP has been entered: nominal G.654.E MFD fit is not product validation'
 ]
};
console.log('COMSOL_STEP_INDEX_BENCHMARK_JSON='+JSON.stringify(out));
console.log('MESH_TABLE_CSV=elements,neff,MFD_PII_um,Aeff_um2,residual');
for(const q of points)console.log('MESH_TABLE_CSV='+[q.elements,q.neff,q.MFD_PII_um,q.Aeff_um2,q.relative_residual].join(','));
console.log('COMSOL_BENCHMARK_PASS');
