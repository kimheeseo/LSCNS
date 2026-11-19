import test from 'node:test';import assert from 'node:assert/strict';import {validateStep6Result} from '../research_step6_v088.mjs';
test('STEP6 rejects foreign schema, nonfinite results and invalid efficiency',()=>{
 for(const r of [{schema:'LSCNS_step6_fake_v088',certified:false},{schema:'LSCNS_step6_splice_v088',certified:false,eta_Et:1.1},{schema:'LSCNS_step6_mcf_v088',certified:false,modes:[{neff_real:NaN}]}])assert.throws(()=>validateStep6Result(r));
});
test('STEP6 imports explicit candidate schemas',()=>{
 for(const r of [{schema:'LSCNS_step6_bend_v088',certified:false,cases:[]},{schema:'LSCNS_step6_splice_v088',validation:{certified:false},eta_Et:.9},{schema:'LSCNS_step6_pm_v088',certified:false,modal_birefringence:{B:null}}])assert.equal(validateStep6Result(r),r);
});
