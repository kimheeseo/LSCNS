/**
 * Statistical PMD / DGD model under strong random mode coupling.
 * DGD is Maxwell-distributed and scales as sqrt(length).
 * The coefficient convention (mean vs RMS DGD per sqrt(km)) is explicit.
 * Not derived from current scalar LP01 FEM; no fiber-level PMD certification.
 */
function erfApprox(x){
 const sign=x<0?-1:1,z=Math.abs(x),t=1/(1+0.3275911*z);
 const poly=(((((1.061405429*t-1.453152027)*t)+1.421413741)*t-0.284496736)*t+0.254829592)*t;
 const y=1-poly*Math.exp(-z*z);
 return sign*y;
}
export function maxwellCDF(x,scale){
 if(!(Number.isFinite(scale)&&scale>=0&&Number.isFinite(x)))throw Error('DGD 분포 입력값을 확인하세요.');
 if(x<=0)return 0;if(scale===0)return 1;
 const y=x/scale;
 return Math.max(0,Math.min(1,erfApprox(y/Math.SQRT2)-Math.sqrt(2/Math.PI)*y*Math.exp(-y*y/2)));
}
export function maxwellPDF(x,scale){
 if(!(Number.isFinite(x)&&Number.isFinite(scale)&&scale>=0))throw Error('DGD 밀도 입력 오류');
 if(scale===0||x<0)return 0;
 const z=x/scale;
 return Math.sqrt(2/Math.PI)*(z*z/scale)*Math.exp(-z*z/2);
}
export function maxwellQuantile(probability,scale){
 if(!Number.isFinite(probability)||probability<0||probability>=1||!Number.isFinite(scale)||scale<0)
  throw Error('분위수 범위: 확률 0≤p<1, 스케일 ≥0');
 if(scale===0||probability===0)return 0;
 let low=0,high=12*scale;
 for(let i=0;i<65;i++){
  const mid=(low+high)/2;
  if(maxwellCDF(mid,scale)<probability)low=mid;else high=mid;
 }
 return (low+high)/2;
}
export function estimatePMD(input){
 const {pmdCoefficientPsSqrtKm,lengthKm,coefficientDefinition}=input;
 if(!Number.isFinite(pmdCoefficientPsSqrtKm)||pmdCoefficientPsSqrtKm<0||
    pmdCoefficientPsSqrtKm>100)throw Error('PMD 계수는 0~100 ps/√km 범위여야 합니다.');
 if(!Number.isFinite(lengthKm)||lengthKm<0||lengthKm>50000)throw Error('길이는 0~50,000 km여야 합니다.');
 if(!['mean','rms'].includes(coefficientDefinition))throw Error('PMD 계수 정의를 평균 또는 RMS로 지정하세요.');
 const term=pmdCoefficientPsSqrtKm*Math.sqrt(lengthKm);
 const scale=coefficientDefinition==='mean'?term/(2*Math.sqrt(2/Math.PI)):term/Math.sqrt(3);
 const meanDgdPs=2*scale*Math.sqrt(2/Math.PI);
 const rmsDgdPs=Math.sqrt(3)*scale;
 const stdDgdPs=scale*Math.sqrt((3*Math.PI-8)/Math.PI);
 const medianDgdPs=maxwellQuantile(.5,scale);
 const p95DgdPs=maxwellQuantile(.95,scale),p99DgdPs=maxwellQuantile(.99,scale);
 return {
  model:'strong-mode-coupling-Maxwell-DGD-statistics',
  input:{...input},
  meanDgdPs,rmsDgdPs,stdDgdPs,medianDgdPs,p95DgdPs,p99DgdPs,
  maxwellScalePs:scale,
  limitations:[
   'Strong random-coupling, stationary Maxwell DGD approximation',
   'Distribution percentiles are realizations of one link; NOT a standardized cable PMD_Q calculation',
   'No stress/ellipticity/vector-mode eigenanalysis or actual frequency-dependent Jones matrix',
   'Coefficient must be calibrated to its mean-vs-RMS convention'
  ]
 };
}
