import assert from 'node:assert/strict';
import {run3D,DEFAULT_3D} from '../fiber_3d_core.mjs';
import {silicaIndex} from '../solver.mjs';
let passed=0;
function ok(c,m){assert.ok(c,m);passed++;}
const h=run3D({...DEFAULT_3D,profile:'homogeneous',grid:128,half_width_um:45,length_um:90,steps:90,waist_um:6,offset_um:0,snapshots:5});
const zR=Math.PI*silicaIndex(1550)*36/1.55,oracle=12*Math.sqrt(1+(90/zR)**2);
ok(Math.abs(h.history.at(-1).MFD_2nd_um/oracle-1)<.015,'analytic Gaussian diffraction');
ok(h.max_power_error<1e-10,'paraxial FFT power conservation');
ok(Math.abs(h.history[0].MFD_2nd_um-12)<.02,'initial Gaussian MFD');
for(const profile of ['step','graded','trench','multilayer']){const r=run3D({...DEFAULT_3D,profile,grid:64,steps:35,length_um:200,snapshots:5});
ok(r.max_power_error<1e-9,profile+' power conservation');
ok(r.history.at(-1).Aeff_um2>0,profile+' positive Aeff');
ok(r.slices.length===5,profile+' snapshots');
ok(r.confinementLoss_dB_km===null,profile+' no invented loss');
}
const c=run3D({...DEFAULT_3D,grid:64,steps:50,length_um:200,offset_um:0});
const free=run3D({...DEFAULT_3D,profile:'homogeneous',grid:64,steps:50,length_um:200,offset_um:0});
ok(c.history.at(-1).MFD_2nd_um<free.history.at(-1).MFD_2nd_um,'guided vs homogeneous');
for(const b of [{grid:70},{profile:'hcf'},{steps:0},{half_width_um:5}]){
 assert.throws(()=>run3D({...DEFAULT_3D,...b}));passed++;
}
console.log(JSON.stringify({tests_passed:passed,gaussian_mfd_analytic:oracle,gaussian_mfd_BPM:h.history.at(-1).MFD_2nd_um,max_power_error:h.max_power_error},null,2));