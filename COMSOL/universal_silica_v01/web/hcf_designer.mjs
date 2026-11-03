/**
 * NANF HCF analytical research pre-designer (geometry + ARROW + ideal capillary).
 * NOT FEM, NOT a confinement-loss predictor.
 * Source conventions: Poletti (Opt Express 2014), ARROW resonance approximation,
 * Marcatili/Schmeltzer ideal capillary fundamental mode proxy.
 */
export const SPEED=299792458;
export const U01=2.404825557695773;
export const DEFAULT={
  coreDiameterUm:32, tubeCount:6, nesting:2,
  outerTubeRatio:0.78, innerRatio:0.61, thirdRatio:0.55,
  gapCoreUm:0.25, minNeighborGapUm:0.35, nestedClearanceUm:0.20,
  wall1Um:0.58, wall2Um:0.45, wall3Um:0.36,
  wavelengthNm:1550, airIndex:1.00027,
  autoFit:true,
  scanMinNm:700, scanMaxNm:2300
};
const finite=x=>typeof x==='number'&&Number.isFinite(x);
function inRange(x,a,b,label){if(!finite(x)||x<a||x>b)throw Error(`${label}: ${a}–${b} 범위 값이 필요합니다.`);}
export function validate(p){
  inRange(p.coreDiameterUm,16,120,'공기 코어 직경 (µm)');
  inRange(p.tubeCount,4,12,'튜브 개수');
  if(!Number.isInteger(p.tubeCount))throw Error('튜브 개수는 정수입니다.');
  if(!Number.isInteger(p.nesting)||p.nesting<1||p.nesting>3)throw Error('중첩 단계는 1·2·3입니다.');
  inRange(p.outerTubeRatio,0.2,1.4,'바깥 튜브/코어 직경 비율');
  inRange(p.innerRatio,0.3,0.8,'2단계 직경 비율');
  inRange(p.thirdRatio,0.3,0.8,'3단계 직경 비율');
  inRange(p.gapCoreUm,0,6,'코어–튜브 간격');
  inRange(p.minNeighborGapUm,0,6,'튜브 간 최소 간격');
  inRange(p.nestedClearanceUm,0,1.5,'중첩 내부 여유');
  for(const k of ['wall1Um','wall2Um','wall3Um'])inRange(p[k],0.10,1.6,k);
  inRange(p.wavelengthNm,300,3000,'분석 파장');
  inRange(p.airIndex,1,1.01,'공기 굴절률');
  inRange(p.scanMinNm,300,3000,'스윕 시작');
  inRange(p.scanMaxNm,300,3500,'스윕 종료');
  if(p.scanMaxNm<=p.scanMinNm)throw Error('스윕 종료가 시작보다 커야 합니다.');
  return p;
}
export function silicaIndex(lambdaUm){
  inRange(lambdaUm,.22,3.65,'실리카 Sellmeier 파장 (µm)');
  const B=[0.6961663,0.4079426,0.8974794],C=[0.0684043**2,0.1162414**2,9.896161**2];
  const x=lambdaUm*lambdaUm;
  return Math.sqrt(1+B.reduce((s,b,i)=>s+b*x/(x-C[i]),0));
}
/** Ring arrangement; each nested ring is offset outward along the radial tube axis.
 * Adjacent outer tube overlap is blocked, or visibly autolimited when autoFit=true.
 */
export function buildGeometry(input){
  const p=validate({...input});
  const coreR=p.coreDiameterUm/2, s=Math.sin(Math.PI/p.tubeCount);
  const requestedD=p.coreDiameterUm*p.outerTubeRatio;
  const maxD=(2*s*(coreR+p.gapCoreUm)-p.minNeighborGapUm)/(1-s);
  if(maxD<=0)throw Error('설정한 튜브 간 최소 간격이 너무 큽니다.');
  const actualD=p.autoFit?Math.min(requestedD,maxD):requestedD;
  const r1=actualD/2,r2=r1*p.innerRatio,r3=r2*p.thirdRatio;
  const radii=[r1,r2,r3].slice(0,p.nesting);
  const walls=[p.wall1Um,p.wall2Um,p.wall3Um].slice(0,p.nesting);
  const offsets=[0];
  for(let j=1;j<p.nesting;j++){
    const next=radii[j-1]-walls[j-1]-radii[j]-p.nestedClearanceUm;
    if(next<0)throw Error(`${j+1}중 튜브가 이전 단계 내경에 들어가지 않습니다. 직경 비율을 줄이세요.`);
    offsets.push(offsets[j-1]+next);
  }
  for(let j=0;j<p.nesting;j++){
    if(walls[j]>=radii[j]*0.8)throw Error(`${j+1}번째 튜브 벽이 내부 반경에 비해 너무 두껍습니다.`);
  }
  const ringRadius=coreR+r1+p.gapCoreUm;
  const neighborGap=2*ringRadius*s-actualD;
  if(neighborGap < p.minNeighborGapUm-1e-7)throw Error(`바깥 튜브가 서로 겹칩니다: 실제 간격 ${neighborGap.toFixed(3)} µm. 자동 맞춤을 켜거나 D₁을 줄이세요.`);
  const rings=[];
  for(let i=0;i<p.tubeCount;i++){
    const angle=2*Math.PI*i/p.tubeCount-Math.PI/2;
    const ux=Math.cos(angle),uy=Math.sin(angle);
    const layers=radii.map((radius,j)=>({
      level:j+1,x:ux*(ringRadius+offsets[j]),y:uy*(ringRadius+offsets[j]),
      radius,wall:walls[j],outerDiameterUm:2*radius
    }));
    rings.push({index:i+1,angle,layers});
  }
  const coverage=Math.min(1,p.tubeCount*Math.asin(Math.min(1,r1/ringRadius))/Math.PI);
  const rMax=ringRadius+offsets[offsets.length-1]+radii.at(-1);
  return {input:p,coreRadiusUm:coreR,outerTubeDiameterUm:actualD,
    requestedOuterTubeDiameterUm:requestedD,autoFitted:actualD<requestedD-1e-7,
    maxAllowedOuterTubeDiameterUm:maxD,ringRadiusUm:ringRadius,
    minAdjacentGapUm:neighborGap,angularCoverageFraction:coverage,
    radiiUm:radii,wallsUm:walls,rings,extentUm:Math.max(ringRadius+r1,rMax)};
}
export function rootWallWavelengthUm(thicknessUm,order,airIndex=1.00027){
  inRange(thicknessUm,.05,2,'벽 두께');
  inRange(order,.5,10,'차수');
  const fn=wl=>2*thicknessUm*Math.sqrt(silicaIndex(wl)**2-airIndex**2)/order-wl;
  let lo=.221,hi=3.60;
  if(fn(lo)*fn(hi)>0)return null;
  for(let k=0;k<90;k++){const mid=(lo+hi)/2;if(fn(mid)>0)lo=mid;else hi=mid;}
  return (lo+hi)/2;
}
export function wallResonances(p){
  validate(p);
  const out=[];
  [p.wall1Um,p.wall2Um,p.wall3Um].slice(0,p.nesting).forEach((wall,i)=>{
    for(let m=1;m<=3;m++){
      for(const kind of ['resonance','antiresonance']){
        const ord=kind==='resonance'?m:m-0.5;
        const lam=rootWallWavelengthUm(wall,ord,p.airIndex);
        if(lam!==null)out.push({layer:i+1,kind,order:m,wallUm:wall,wavelengthNm:lam*1000,frequencyTHz:SPEED/(lam*1e-6)/1e12});
      }
    }
  });
  return out;
}
export function wallDetuning(p,lambdaNm=p.wavelengthNm){
  const lambdaUm=lambdaNm/1000;
  const q=Math.sqrt(silicaIndex(lambdaUm)**2-p.airIndex**2);
  return [p.wall1Um,p.wall2Um,p.wall3Um].slice(0,p.nesting).map((wallUm,i)=>{
    const phase=2*wallUm*q/lambdaUm;
    return {layer:i+1,phase,detuning:Math.abs(Math.sin(Math.PI*phase))};
  });
}
export function besselJ0(x){
  let t=1,ans=1;const y=-(x*x/4);
  for(let k=1;k<=22;k++){t*=y/(k*k);ans+=t;if(Math.abs(t)<1e-16)break;}
  return ans;
}
export function besselJ1(x){
  let t=x/2,ans=t;const y=-(x*x/4);
  for(let k=1;k<=22;k++){t*=y/(k*(k+1));ans+=t;if(Math.abs(t)<1e-16)break;}
  return ans;
}
export function capillaryMetrics(p,lambdaNm=p.wavelengthNm){
  const R=p.coreDiameterUm/2,lambdaUm=lambdaNm/1000;
  const k=(U01/(2*Math.PI*R))**2;
  const neff2=p.airIndex**2-k*lambdaUm**2;
  if(neff2<=0)throw Error('모세관 근사 전파 조건을 벗어났습니다.');
  const neff=Math.sqrt(neff2);
  const D=lambdaUm/SPEED*1e12*k*p.airIndex**2/(neff**3);
  const N=800;let i2=0,i4=0,imoment=0,igrad=0;
  // Simpson integration over radial dimension x = r/R.
  for(let j=0;j<=N;j++){
    const x=j/N,weight=j===0||j===N?1:j%2===0?2:4;
    const phi=besselJ0(U01*x), grad=besselJ1(U01*x);
    i2+=weight*phi*phi*x;
    i4+=weight*phi**4*x;
    imoment+=weight*phi*phi*x**3;
    igrad+=weight*grad*grad*x;
  }
  i2/=3*N;i4/=3*N;imoment/=3*N;igrad/=3*N;
  const Aeff=2*Math.PI*R*R*(i2*i2/i4);
  const mfd2=2*R*Math.sqrt(2*imoment/i2);
  const mfdPII=2*R*Math.sqrt(2*i2/(U01**2*igrad));
  return {model:'ideal-air-capillary-J0-Dirichlet—not-nested-fiber-FEM',
    wavelengthNm:lambdaNm,neff,Aeff_um2:Aeff,MFD_2nd_um:mfd2,MFD_PII_um:mfdPII,
    D_ps_per_nm_km:D,nAirAssumed:p.airIndex,
    confinementLossDbKm:null,modelIncludesTubeCount:false,modelIncludesNestedGeometry:false};
}
export function externalComplexNeffLoss(imAbs,wavelengthNm){
  inRange(imAbs,0,0.001,'외부 Im(neff) 절댓값');
  inRange(wavelengthNm,300,3000,'파장');
  const lamM=wavelengthNm*1e-9;
  return 40*Math.PI/(lamM*Math.log(10))*1000*imAbs;
}
export function spectrum(p,count=121){
  const a=p.scanMinNm,b=p.scanMaxNm;
  if(count<5||count>801)throw Error('파장 스윕 점수 5–801');
  return Array.from({length:count},(_,i)=>{
    const wl=a+(b-a)*i/(count-1);
    const mode=capillaryMetrics(p,wl);
    const detunings=wallDetuning(p,wl);
    return {lambdaNm:wl,neff:mode.neff,D:mode.D_ps_per_nm_km,
      minWallDetuning:Math.min(...detunings.map(x=>x.detuning)),
      layerDetunings:detunings.map(x=>x.detuning)};
  });
}
export function tubeCountSweep(p){
  return Array.from({length:9},(_,i)=>{
    const N=i+4;
    try{
      const g=buildGeometry({...p,tubeCount:N});
      return {tubeCount:N,valid:true,autoFitted:g.autoFitted,
        actualD1Um:g.outerTubeDiameterUm,coverageFraction:g.angularCoverageFraction,
        neighborGapUm:g.minAdjacentGapUm,
        confinementLossDbKm:null};
    }catch(e){return {tubeCount:N,valid:false,error:String(e.message),confinementLossDbKm:null};}
  });
}
export function designSummary(p){
  const geometry=buildGeometry(p);
  const mode=capillaryMetrics(p);
  const detuning=wallDetuning(p);
  const resonances=wallResonances(p);
  return {modelVersion:'hcf-analytical-designer-v0.1',source:'ARROW wall resonance + ideal J0 capillary; NO full-vector FEM',
    input:{...p},geometry:{...geometry,rings:undefined},mode,
    currentWallDetuning:detuning,resonances,
    warnings:['No quantitative nested-tube confinement-loss prediction. Loss in dB/km requires independently converged complex vector FEM/PML.',
      'Tube count/diameter/nesting do NOT modify capillary neff/MFD/Aeff/D in this proxy.',
      'Wall-resonance estimates are local thin-wall ARROW approximations; tube coupling shifts spectral features.',
      'Air index is constant with wavelength; material/gas dispersion, scattering and absorption are excluded.']};
}
