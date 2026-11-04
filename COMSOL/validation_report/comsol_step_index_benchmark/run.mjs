/**
 * Public COMSOL Step-Index Fiber neff benchmark, DIRECTLY using production LSCNS solver.
 * No reference tuning: COMSOL's reference neff is NOT passed to solve().
 * This tutorial is NOT a measured G.654.E manufacturer RIP.
 * https://doc.comsol.com/6.4/doc/com.comsol.help.models.rf.step_index_fiber/step_index_fiber.html
 */
import assert from 'node:assert/strict';
import {PRESETS,claddingIndex,solve,meshStudy,validate} from '../../universal_silica_v01/web/solver.mjs';
const p=structuredClone(PRESETS.comsol_step);
assert.equal(p.profile,'step');
assert.equal(p.wavelength_nm,1550);
assert.equal(p.core_radius_um,8);
assert.equal(p.outer_radius_um,40);
assert.equal(p.clad_index_override,1.4378);
assert.ok(Math.abs((p.clad_index_override+p.delta_n_core)-1.4457)<1e-12);
validate(p);
assert.equal(claddingIndex(p),1.4378);
assert.ok(Math.abs(claddingIndex(PRESETS.g654)-1.4378)>0.001,'standard silica material must be unchanged');
assert.throws(()=>validate({...p,clad_index_override:-1}),/고정 클래딩/);
const referencePrinted=1.4444;
const analyticScalarLP01=1.4444018444092133;
const points=[200,350,500,750,1000,1400,1800].map(elements=>{
 const s=solve({...p,elements});
 assert.equal(s.guided_proxy,true);
 assert.ok(s.residual<1e-7);
 assert.ok(s.n_eff>p.clad_index_override&&s.n_eff<p.clad_index_override+p.delta_n_core);
 return {elements:s.elements,neff:s.n_eff,MFD_PII_um:s.MFD_PII_um,Aeff_um2:s.Aeff_um2,
  relative_residual:s.residual};
});
const actual=points.find(x=>x.elements===750);
const absoluteError=Math.abs(actual.neff-referencePrinted);
const relativeErrorPct=100*absoluteError/referencePrinted;
assert.ok(relativeErrorPct<5,'user-requested COMSOL reference neff relative error <5%');
assert.ok(absoluteError<0.00005,'COMSOL published 4-decimal rounding precision');
assert.ok(Math.abs(actual.neff-analyticScalarLP01)<1e-6,'independent scalar Bessel LP01 crosscheck');
assert.ok(Math.abs(points.at(-1).neff-analyticScalarLP01)<Math.abs(points[0].neff-analyticScalarLP01),'mesh-refinement improvement');
const originalG654=solve({...PRESETS.g654,elements:750});
assert.ok(Math.abs(originalG654.n_eff-1.445718968)<1e-7,'legacy G654 preset regression');
const out={
 schema:'LSCNS_COMSOL_public_step_index_v2',wavelength_nm:p.wavelength_nm,
 core_radius_um:p.core_radius_um,outer_radius_um:p.outer_radius_um,
 n_core:p.clad_index_override+p.delta_n_core,n_clad:p.clad_index_override,
 reference_comsol_neff_4dp:referencePrinted,
 lscns_scalar_fem_neff:actual.neff,
 neff_absolute_error: absoluteError,
 neff_relative_error_percent:relativeErrorPct,
 requested_threshold_percent:5,
 status: relativeErrorPct<5?'PASS':'FAIL',
 reference_rounding_compatible:absoluteError<0.00005,
 independent_analytical_scalar_lp01_neff:analyticScalarLP01,
 mesh_study:points,
 limitations:['Published COMSOL neff has 4 decimals only','Model is generic step-index and not commercial G.654.E','Scalar FEM vs COMSOL full-vector electromagnetic formulation','Fixed n at 1550 nm is not a material dispersion model']
};
console.log('COMSOL_STEP_INDEX_BENCHMARK_JSON='+JSON.stringify(out));
console.log('MESH_TABLE_CSV=elements,neff,MFD_PII_um,Aeff_um2,residual');
for(const x of points)console.log('MESH_TABLE_CSV='+[x.elements,x.neff,x.MFD_PII_um,x.Aeff_um2,x.relative_residual].join(','));
console.log('COMSOL_BENCHMARK_PASS');
