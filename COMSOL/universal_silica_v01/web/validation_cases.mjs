/**
 * M1 independent public references and pre-update numeric regression baselines.
 * The COMSOL value is rounded to 4 decimals; never use it as a solver input.
 */
export const COMSOL_PUBLIC = Object.freeze({
 id:"comsol-step-index-145",name:"COMSOL Step-Index Fiber · Application ID 145",
 source:"https://doc.comsol.com/6.4/doc/com.comsol.help.models.rf.step_index_fiber/step_index_fiber.html",
 wavelength_nm:1550,core_radius_um:8,outer_radius_um:40,
 n_core:1.4457,n_clad:1.4378,reference_neff_rounded:1.4444,
 roundingHalfWidth:0.00005,
 comment:"Published full-vector HE11 result rounded to 4 decimals; not a high-precision FEM golden value."
});
export const LP01_CASES=Object.freeze([
 Object.freeze({
  id:"comsol_step",name:"COMSOL matching geometry & indices",
  wavelength_nm:1550,core_radius_um:8,outer_radius_um:40,n_core:1.4457,n_clad:1.4378,
  elements:750,expectedAnalytic:1.4444018444092133,expectedV:4.894537730479192,
  source:"COMSOL geometry + independent SciPy Bessel reference"
 }),
 Object.freeze({
  id:"single_mode",name:"Independent single-mode step-index (V < 2.405)",
  wavelength_nm:1550,core_radius_um:4.1,outer_radius_um:40,n_core:1.4472,n_clad:1.444,
  elements:750,expectedAnalytic:1.444863128995567,expectedV:1.598623098047904,
  source:"Independent SciPy scipy.special + brentq, not a COMSOL published number"
 })
]);
export const REGRESSION=Object.freeze([
 Object.freeze({id:"g654",neff:1.445718967909122,mfd:12.500052522436938,aeff:121.92389530338414,D:19.825664345368367}),
 Object.freeze({id:"g652",neff:1.4461047281745436,mfd:10.182785332301778,aeff:79.29336128699885,D:16.99288790683873}),
 Object.freeze({id:"graded",neff:1.4458288150125134,mfd:10.192939527752662,aeff:76.88413268997199,D:14.919171294740607}),
 Object.freeze({id:"trench",neff:1.4460435091028094,mfd:9.839323631520045,aeff:74.82460399998054,D:17.8038792615834}),
 Object.freeze({id:"multi",neff:1.447001785935375,mfd:9.582447746912562,aeff:71.68879956793322,D:19.16257095579606})
]);
export const TOL=Object.freeze({
 lp01Neff:1e-6,comsolRoundedNeff:5e-5,analyticVsSciPy:1e-10,
 neff:1e-8,mfdUm:1e-4,aeffUm2:1e-4,dispersion:0.005,residual:1e-7
});
