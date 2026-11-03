/**
 * Velamuri et al., COMSOL Conference 2018 bend-loss paper: engineering bridge.
 * https://www.comsol.com/paper/study-of-bending-losses-in-optical-fibers-using-comsol-multiphysics-61111
 *
 * Important: neither this file nor existing scalar radial FEM computes curved 2D
 * complex eigenmodes/PML or paper's Geometrically Exact Beam Theory (GEBT).
 *
 * Supported (separately, explicitly):
 *   1) conformal straight-guide equivalent index n_eq(x) = n_stress(x)(1+x/R)
 *   2) optional simple linear elastic-beam strain/Poisson proxy, NOT GEBT
 *   3) COMSOL complex mode n_eff imaginary part -> dB/m and dB/turn
 *   4) convergence assessment on USER-SUPPLIED PML-sweep imaginary indices
 */
import {silicaIndex,deltaIndex} from './solver.mjs?v=20261103-slope-v1';
export const VELAMURI_REFERENCE=Object.freeze({
 title:'Study of Bending Losses in Optical Fibers using COMSOL',
 authors:'A. Velamuri et al.',year:2018,p11:0.113,p12:0.252,
 pmlStudySuggestedLambda:7,exampleBendDiameterMm:10,exampleWavelengthNm:1550,
 reportedBifSimLossDbTurn:0.0119,reportedBifMeasuredMeanLossDbTurn:0.012,
 source:'https://www.comsol.com/paper/study-of-bending-losses-in-optical-fibers-using-comsol-multiphysics-61111'
});
const TOL=1e-10;
function real(value,label){
 if(!Number.isFinite(value))throw Error(label+' 값은 유한한 숫자여야 합니다.');
 return value;
}
function validateDiameter(d){real(d,'굽힘 직경');if(!(d>0&&d<=10000))throw Error('굽힘 직경은 0보다 크고 10,000 mm 이하여야 합니다.');}
function validateWavelength(nm){real(nm,'파장');if(!(nm>210&&nm<3700))throw Error('파장이 실리카 Sellmeier 모델 범위 밖입니다.');}
export function stressOpticIndex(n,epsilon1,epsilon2,epsilon3,p11=0.113,p12=0.252){
 [n,epsilon1,epsilon2,epsilon3,p11,p12].forEach((v,i)=>real(v,['n','epsilon1','epsilon2','epsilon3','P11','P12'][i]));
 if(n<=0||Math.max(Math.abs(epsilon1),Math.abs(epsilon2),Math.abs(epsilon3))>.2)
  throw Error('굴절률은 양수이고 스트레인은 절댓값 0.2 이하여야 합니다.');
 const nStress=n*(1-0.5*n*n*(p11*epsilon1+p12*(epsilon2+epsilon3)));
 if(nStress<=0)throw Error('응력-광학 보정 결과의 굴절률이 양수가 아닙니다.');
 return nStress;
}
export function equivalentBendIndex({straightIndex,xUm,bendDiameterMm,
 stressMode='none',poissonRatio=0.17}){
 real(straightIndex,'직선 광섬유 굴절률');real(xUm,'횡좌표 x');validateDiameter(bendDiameterMm);
 if(!(straightIndex>0))throw Error('직선 광섬유 굴절률은 양수여야 합니다.');
 if(!['none','beam_proxy'].includes(stressMode))throw Error('지원하지 않는 응력 근사 모드입니다.');
 if(!Number.isFinite(poissonRatio)||poissonRatio<0||poissonRatio>=.5)throw Error('Poisson ratio는 0~0.5 미만이어야 합니다.');
 const RUm=bendDiameterMm*500;
 if(Math.abs(xUm)>=RUm)throw Error('굽힘 중립축의 거리보다 큰 위치는 본 근사 범위를 벗어납니다.');
 const longitudinalStrain=xUm/RUm;
 let ns=straightIndex;
 if(stressMode==='beam_proxy'){
   const ep1=longitudinalStrain,ep2=-poissonRatio*ep1,ep3=ep2;
   ns=stressOpticIndex(straightIndex,ep1,ep2,ep3);
 }
 const geometricFactor=1+longitudinalStrain;
 return {nStraight:straightIndex,nStress:ns,nEquivalent:ns*geometricFactor,
  nGeometryOnly:straightIndex*geometricFactor,epsilonLongitudinal:longitudinalStrain,
  bendRadiusUm:RUm,stressMode,assumption:stressMode==='beam_proxy'?
   'first-order linear beam strain + isotropic Poisson proxy, NOT GEBT':
   'conformal mapping only, stress optical correction disabled'};
}
export function bentIndexCut(model,{wavelengthNm=1550,bendDiameterMm=10,claddingRadiusUm=62.5,
 stressMode='none',poissonRatio=.17,samples=301}={}){
 validateWavelength(wavelengthNm);validateDiameter(bendDiameterMm);
 if(!Number.isFinite(claddingRadiusUm)||claddingRadiusUm<=0||claddingRadiusUm>1000)
  throw Error('클래딩 반경은 0~1000 µm 이내 양수여야 합니다.');
 if(!Number.isInteger(samples)||samples<11||samples>5001)throw Error('표시 샘플은 11~5001개여야 합니다.');
 if(claddingRadiusUm>=bendDiameterMm*500)throw Error('클래딩이 굽힘 반경보다 클 수 없습니다.');
 if(model.profile==='mcf')throw Error('기존 MCF 보간 모델의 2D 실제 인덱스 프로파일은 제공되지 않습니다.');
 const rows=[];const nClad=silicaIndex(wavelengthNm);
 for(let j=0;j<samples;j++){
   const xUm=claddingRadiusUm*(-1+2*j/(samples-1));
   const n=nClad+deltaIndex(model,Math.abs(xUm));
   const v=equivalentBendIndex({straightIndex:n,xUm,bendDiameterMm,stressMode,poissonRatio});
   rows.push({xUm,nStraight:v.nStraight,nGeometryOnly:v.nGeometryOnly,nEquivalent:v.nEquivalent,
    epsilonLongitudinal:v.epsilonLongitudinal});
 }
 return {rows,wavelengthNm,bendDiameterMm,bendRadiusMm:bendDiameterMm/2,claddingRadiusUm,
  stressMode,source:'simulated geometry + first-order conformal map, NOT computed bend loss',
  caveats:['Index cross section is not a complex eigenmode or loss prediction',
   'For actual GEBT supply principal strain fields from structural mechanics',
   'Real manufacturer RIP, coating, and stress coefficients must be calibrated']};
}
export function bendLossFromComplexNeff({wavelengthNm,bendDiameterMm,neffImagMagnitude,
 turns=1,neffReal=null}){
 validateWavelength(wavelengthNm);validateDiameter(bendDiameterMm);
 real(neffImagMagnitude,'복소 n_eff 허수부');
 if(neffImagMagnitude<0||neffImagMagnitude>1)throw Error('Im(n_eff) 크기는 0~1 사이로 입력하세요. 솔버의 부호는 별도 확인하세요.');
 if(!Number.isInteger(turns)||turns<0||turns>100000)throw Error('권선 수는 0~100000 정수로 입력하세요.');
 if(neffReal!==null)real(neffReal,'n_eff 실수부');
 const betaImagPerM=(2*Math.PI/(wavelengthNm*1e-9))*neffImagMagnitude;
 const lossDbPerM=(20/Math.log(10))*betaImagPerM; // POWER dB attenuation; exp(i beta z) convention
 const lengthPerTurnM=Math.PI*bendDiameterMm/1000;
 const lossDbPerTurn=lossDbPerM*lengthPerTurnM;
 return {model:'complex-neff-postprocessor-no-FEM',
  wavelengthNm,bendDiameterMm,bendRadiusMm:bendDiameterMm/2,neffReal,neffImagMagnitude,
  betaImagPerM,lossDbPerM,lengthPerTurnM,lossDbPerTurn,turns,totalBendLossDb:turns*lossDbPerTurn,
  convention:'Im(neff) nonnegative attenuation magnitude, E(z)=E0 exp(i beta z)',
  caveat:'External complex mode/PML must be actually computed; numerical input alone is not predictive'};
}
export function imagNeffFromLoss({wavelengthNm,bendDiameterMm,lossDbPerTurn}){
 validateWavelength(wavelengthNm);validateDiameter(bendDiameterMm);
 real(lossDbPerTurn,'loss (dB/turn)');
 if(lossDbPerTurn<0)throw Error('손실은 0 이상이어야 합니다.');
 return lossDbPerTurn/((20/Math.log(10))*(2*Math.PI/(wavelengthNm*1e-9))*Math.PI*bendDiameterMm/1000);
}
export function assessPMLConvergence({wavelengthNm,bendDiameterMm,rows,tolerancePercent=2}){
 validateWavelength(wavelengthNm);validateDiameter(bendDiameterMm);
 if(!Array.isArray(rows)||rows.length<2)throw Error('PML 두께별 외부 FEM의 Im(neff) 값 2개 이상이 필요합니다.');
 if(!Number.isFinite(tolerancePercent)||tolerancePercent<=0||tolerancePercent>100)
  throw Error('수렴 기준은 0~100% 사이로 입력하세요.');
 const sorted=rows.map(x=>({...x})).sort((a,b)=>a.thicknessLambda-b.thicknessLambda);
 let before=0;
 for(const r of sorted){
   real(r.thicknessLambda,'PML 두께 (λ)');real(r.neffImagMagnitude,'Im(neff)');
   if(!(r.thicknessLambda>before)||r.neffImagMagnitude<=0)throw Error('PML 두께가 서로 다르고 오름차순이어야 하며 Im(neff) 크기는 양수여야 합니다.');
   before=r.thicknessLambda;
 }
 const converted=sorted.map(x=>({...x,lossDbPerTurn:bendLossFromComplexNeff({
  wavelengthNm,bendDiameterMm,neffImagMagnitude:x.neffImagMagnitude}).lossDbPerTurn}));
 const a=converted.at(-2).lossDbPerTurn,b=converted.at(-1).lossDbPerTurn;
 const relativeChangePercent=Math.abs(b-a)/Math.max(Math.abs(b),TOL)*100;
 return {passed:relativeChangePercent<=tolerancePercent,relativeChangePercent,tolerancePercent,
  rows:converted,referenceIsSevenLambda:converted.at(-1).thicknessLambda>=7,
  caveat:'Only tests stability between last two externally computed PML results; this is not mesh or physics validation'};
}
