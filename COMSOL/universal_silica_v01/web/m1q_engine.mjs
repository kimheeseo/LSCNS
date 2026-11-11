/**
 * M1 v0.5.1 read-only quantitative verification. No shared UI/model mutation.
 * FEM uses the *unchanged* original scalar P1 radial solver.
 * Independent LP01 Bessel oracle is imported from validation_lp01.mjs.
 * Oracle numerical method:
 *   F(u)=u J1(u)/J0(u) - w K1(w)/K0(w), w=sqrt(V²-u²).
 *   Bracket: 0<u<min(V, first zero J0=2.4048255577).
 *   Bisection: at most 75 iterations, stop when |upper-lower|<4e-14
 *   (dimensionless u). J functions use convergent series, K ratio uses
 *   1024-interval Simpson quadrature of the positive Bessel-K integral.
 *   Check independent SciPy Brent reference at |delta neff|<=1e-10 and
 *   characteristic F(u) residual <=1e-8 (not a FEM residual).
 */
import {PRESETS,solve,dispersion} from "./solver.mjs";
import {lp01Characteristic} from "./validation_lp01.mjs";
import {M1Q_VERSION,M1Q_SCHEMA,M1Q_CASES,M1Q_PREUPDATE,M1Q_TOLERANCES as T,M1Q_COMSOL_URL} from "./m1q_cases.mjs";
function finite(n){return typeof n==="number"&&Number.isFinite(n);}
export function m1qFixtureModel(c){
 return {...PRESETS.comsol_step,profile:"step",wavelength_nm:c.wavelength_nm,
  core_radius_um:c.core_radius_um,delta_n_core:c.delta_n_core,
  clad_index_override:c.n_clad,outer_radius_um:c.outer_radius_um,elements:c.elements};
}
export function verifyM1QCases(){
 return M1Q_CASES.map(c=>{
  const base={
   id:c.id,name:c.name,sourceType:c.sourceType,sourceUrl:c.sourceUrl??null,
   inputs:{wavelength_nm:c.wavelength_nm,core_radius_um:c.core_radius_um,
    delta_n_core:c.delta_n_core,n_clad:c.n_clad,n_core:c.n_clad+c.delta_n_core,
    outer_radius_um:c.outer_radius_um,elements:c.elements},
   referenceNeffSciPy:c.scipy_neff,toleranceAbsNeff:T.absNeff,
   comsolPublishedNeff:c.comsol_neff_rounded??null,
   comsolStatus:c.comsol_neff_rounded==null?"미검증":"미검증",
   analyticNeff:null,femNeff:null,absError:null,relativePercent:null,
   V:null,analyticResidual:null,femEigenResidual:null,
   passed:false,status:"FAIL",error:null
  };
  try{
   const anal=lp01Characteristic({
    wavelength_nm:c.wavelength_nm,core_radius_um:c.core_radius_um,
    n_core:c.n_clad+c.delta_n_core,n_clad:c.n_clad
   });
   const fem=solve(m1qFixtureModel(c));
   const absError=Math.abs(fem.n_eff-anal.n_eff);
   const relativePercent=absError/Math.abs(anal.n_eff)*100;
   const sciPyAbs=Math.abs(anal.n_eff-c.scipy_neff);
   const valid=finite(anal.n_eff)&&finite(fem.n_eff)&&finite(fem.residual)&&
    Math.abs(anal.residual)<=1e-8&&sciPyAbs<=T.analyticSciPy&&
    fem.residual<=T.eigenResidual&&fem.guided_proxy===true;
   const passed=valid&&absError<=T.absNeff;
   let comsolStatus="미검증";
   if(c.comsol_neff_rounded!=null){
    comsolStatus=Math.abs(fem.n_eff-c.comsol_neff_rounded)<=T.comsolRounded?"반올림 범위 내":"반올림 범위 외";
   }
   return {...base,analyticNeff:anal.n_eff,femNeff:fem.n_eff,absError,relativePercent,
    V:anal.V,analyticResidual:anal.residual,femEigenResidual:fem.residual,
    sciPyAbsError:sciPyAbs,comsolStatus,status:passed?"PASS":"FAIL",passed};
  }catch(error){return {...base,error:String(error?.message||error)};}
 });
}
export function verifyM1QLegacy(){
 return M1Q_PREUPDATE.map(c=>{
  const sol=solve({...PRESETS[c.id]}),D=dispersion({...PRESETS[c.id]});
  const now={neff:sol.n_eff,mfd:sol.MFD_PII_um,aeff:sol.Aeff_um2,D};
  const before={neff:c.neff,mfd:c.mfd,aeff:c.aeff,D:c.D};
  const deviations={
   neff:Math.abs(now.neff-before.neff),
   mfd:Math.abs(now.mfd-before.mfd),
   aeff:Math.abs(now.aeff-before.aeff),
   D:Math.abs(now.D-before.D)
  };
  const tolerances={
   neff:T.regressionNeff,mfd:T.regressionMfdUm,
   aeff:T.regressionAeffUm2,D:T.regressionD
  };
  const passed=Object.keys(tolerances).every(key=>finite(now[key])&&deviations[key]<=tolerances[key]);
  return {id:c.id,before,after:now,absError:deviations,tolerances,passed,status:passed?"PASS":"FAIL"};
 });
}
export function runM1Quantitative({onProgress=()=>{}}={}){
 onProgress("Step-index LP01 / 독립 SciPy 기준 4개");
 const cases=verifyM1QCases();
 onProgress("기존 5개 프리셋 계산 결과 비교");
 const regression=verifyM1QLegacy();
 const allPass=cases.every(x=>x.passed)&&regression.every(x=>x.passed);
 return {
  schemaVersion:M1Q_SCHEMA,version:M1Q_VERSION,kind:"LSWaveM1QuantitativeVerification",
  generatedAt:new Date().toISOString(),
  source:{comsol:M1Q_COMSOL_URL,
   analytic:"step-index scalar LP01 Bessel characteristic equation; SciPy scipy.special + scipy.optimize.brentq independent fixture",
   legacy:"v0.5.0 solver.mjs blob SHA 6d774b556b83829c77a3b6f3d208eb8ea8da6a3a"},
  passed:allPass,passCount:cases.filter(x=>x.passed).length,
  caseCount:cases.length,regressionPassCount:regression.filter(x=>x.passed).length,
  regressionCount:regression.length,cases,regression,tolerances:{...T},
  unverified:[
   "사용자 지정 a=4.1 µm, Δn=0.005는 COMSOL 공식 공개 예제값이 아니므로 COMSOL 직접 결과 미검증",
   "LP01 스칼라 해석해는 무한 클래딩·약유도 가정. 유한 R 경계와 FEM 메시 오차 존재",
   "1310 nm 케이스의 n_clad=1.4440 및 Δn=0.005는 고정 상수; 재료 분산 미검증",
   "벡터 HE11 편광, PML, HCF, MCF, PMD, 실측 감쇠, 제조사 규격은 이 M1 사례로 검증되지 않음"
  ]
 };
}
