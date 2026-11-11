/**
 * M1 v0.5.1 quantitative verification fixtures. User-configured a=4.1 um,
 * delta n=0.005 is NOT COMSOL's published App 145 geometry (a=8 um).
 * Fixed n_clad=1.4440 in user cases is an explicit benchmark assumption,
 * not a claimed COMSOL value or manufacturer measurement.
 * Analytic baseline values are independently reproduced with SciPy J0/J1/K0/K1
 * + Brent root solve; they are not injected into the browser eigensolver.
 */
export const M1Q_VERSION="0.5.1";
export const M1Q_SCHEMA=1;
export const M1Q_COMSOL_URL="https://doc.comsol.com/6.4/doc/com.comsol.help.models.rf.step_index_fiber/step_index_fiber.html";
export const M1Q_TOLERANCES=Object.freeze({
  absNeff:1e-6,
  relativePercent:1e-4, // Display-only reference: |FEM-analytic| / |analytic| * 100
  analyticSciPy:1e-10,
  eigenResidual:1e-7,
  comsolRounded:5e-5,
  regressionNeff:1e-10,
  regressionMfdUm:1e-6,
  regressionAeffUm2:1e-6,
  regressionD:1e-6
});
export const M1Q_CASES=Object.freeze([
 Object.freeze({id:"official_comsol_145",name:"COMSOL 공식 App 145",
  sourceType:"COMSOL 공식 입력 + SciPy LP01 해석해",sourceUrl:M1Q_COMSOL_URL,
  wavelength_nm:1550,core_radius_um:8,n_clad:1.4378,delta_n_core:0.0079,
  outer_radius_um:40,elements:750,
  scipy_neff:1.4444018444092133,comsol_neff_rounded:1.4444}),
 Object.freeze({id:"requested_41",name:"사용자 지정 4.1 µm / Δn 0.005",
  sourceType:"사용자 지정 조건 + SciPy LP01 해석해; COMSOL 공식 결과 아님",
  wavelength_nm:1550,core_radius_um:4.1,n_clad:1.4440,delta_n_core:0.005,
  outer_radius_um:40,elements:750,scipy_neff:1.446081132401304}),
 Object.freeze({id:"low_v",name:"저V 단일모드 · V < 2.405",
  sourceType:"사용자 검증용 고정 재료 + SciPy LP01 해석해",
  wavelength_nm:1550,core_radius_um:4.1,n_clad:1.4440,delta_n_core:0.0032,
  outer_radius_um:40,elements:750,scipy_neff:1.444863128995567}),
 Object.freeze({id:"different_wavelength",name:"다른 파장 · 1310 nm",
  sourceType:"고정 굴절률 1310 nm 비교 + SciPy LP01 해석해; 실제 재료 분산 미포함",
  wavelength_nm:1310,core_radius_um:4.1,n_clad:1.4440,delta_n_core:0.005,
  outer_radius_um:40,elements:750,scipy_neff:1.446608620045317})
]);
/**
 * Captured BEFORE M1 changes from v0.5.0 solver.mjs
 * blob SHA: 6d774b556b83829c77a3b6f3d208eb8ea8da6a3a.
 * Same copy/preset, same execution/runtime; these are snapshots, not
 * independent physical-material experimental references.
 */
export const M1Q_PREUPDATE=Object.freeze([
 Object.freeze({id:"g654",neff:1.4457189679091202,mfd:12.500052522796532,aeff:121.9238953090936,D:19.825664231522584}),
 Object.freeze({id:"g652",neff:1.446104728174543,mfd:10.182785332592127,aeff:79.29336129096376,D:16.9928875566912}),
 Object.freeze({id:"graded",neff:1.445828815012516,mfd:10.192939528309598,aeff:76.8841326978277,D:14.919171784564464}),
 Object.freeze({id:"trench",neff:1.4460435091028088,mfd:9.839323631718809,aeff:74.82460400259428,D:17.80387925775665}),
 Object.freeze({id:"multi",neff:1.447001785935374,mfd:9.582447747061527,aeff:71.68879956968254,D:19.162570615215405})
]);
