import assert from 'node:assert/strict';
import {PRESETS,solve,dispersion,spectrum,meshStudy,validate} from '../solver.mjs';
const reference={
 g654:[1.445718967909122,12.500052522436938,121.92389530338414,19.825664345368367],
 g652:[1.4461047281745436,10.182785332301778,79.29336128699885,16.99288790683873],
 graded:[1.4458288150125134,10.192939527752662,76.88413268997199,14.919171294740607],
 trench:[1.4460435091028094,9.839323631520045,74.82460399998054,17.8038792615834],
 multi:[1.447001785935375,9.582447746912562,71.68879956793322,19.16257095579606]
};
const ranges=[1e-8,1e-4,1e-4,.005];let checks=0;
for(const [name,ref] of Object.entries(reference)){
 const sol=solve(PRESETS[name]),got=[sol.n_eff,sol.MFD_PII_um,sol.Aeff_um2,dispersion(PRESETS[name])];
 for(let i=0;i<ref.length;i++){const err=Math.abs(got[i]-ref[i]);assert.ok(err<ranges[i],`${name} field ${i} error ${err} > ${ranges[i]}`);checks++;}
 assert.ok(sol.residual<1e-8,`${name} residual`);assert.ok(sol.guided_proxy,`${name} guiding`);checks+=2;
}
const mesh=meshStudy(PRESETS.g654,[200,350,550,850]);assert.ok(Math.abs(mesh[3].n_eff-reference.g654[0])<Math.abs(mesh[0].n_eff-reference.g654[0]));checks++;
const sw=spectrum(PRESETS.g654,1530,1620,10);assert.equal(sw.length,10);assert.ok(sw[0].D_ps_nm_km<sw.at(-1).D_ps_nm_km);checks+=2;
assert.throws(()=>validate({...PRESETS.g654,outer_radius_um:5}));assert.throws(()=>validate({...PRESETS.multi,shells:[[6,0.003],[4,0.001]]}));checks+=2;
console.log(`PASS ${checks} numerical / validation assertions. 5 profiles compared against Python radial FEM; mesh and sweep checked.`);
