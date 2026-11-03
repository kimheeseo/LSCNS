/**
 * M8 two-identical-core *ideal coherent coupling* preview.
 * Four stored scalar-FEM pitch reference points, not a new FEM computation.
 * Results are neither measured cable XT nor a stochastic XT per kilometre.
 */
export const MCF_REFERENCE=Object.freeze([
  Object.freeze({pitch_um:12,delta_neff:0.0002952222507259794}),
  Object.freeze({pitch_um:16,delta_neff:0.00007427797622971966}),
  Object.freeze({pitch_um:20,delta_neff:0.000018904298877409076}),
  Object.freeze({pitch_um:24,delta_neff:0.00000496565386076675})
]);
export const MCF_ASSUMPTIONS='2 identical step-index cores; wavelength 1550 nm, core radius 4.1 um, delta n=0.005; M8 legacy FEM four-point log interpolation; lossless coherent power transfer; no bending, twist, random perturbations, polarization or attenuation; NOT measured cable XT.';
export function calculateMCF(pitch_um,length_mm){
  if(!Number.isFinite(pitch_um)||pitch_um<12||pitch_um>24)throw new RangeError('MCF 코어 간격은 12–24 µm 범위여야 합니다.');
  if(!Number.isFinite(length_mm)||length_mm<0||length_mm>100)throw new RangeError('MCF 전파 길이는 0–100 mm 범위여야 합니다.');
  const i=Math.min(2,Math.floor((pitch_um-12)/4)),a=MCF_REFERENCE[i],b=MCF_REFERENCE[i+1];
  const t=(pitch_um-a.pitch_um)/(b.pitch_um-a.pitch_um);
  const delta_neff=Math.exp(Math.log(a.delta_neff)*(1-t)+Math.log(b.delta_neff)*t);
  const kappa_per_m=Math.PI*delta_neff/1.55e-6;
  const coupling_length_mm=(Math.PI/(2*kappa_per_m))*1000;
  const phi=(Math.PI/2)*(length_mm/coupling_length_mm);
  const p2_raw=Math.sin(phi)**2, p1_raw=Math.cos(phi)**2;
  const core2_power=p2_raw<1e-24?0:p2_raw;
  const core1_power=p1_raw<1e-24?0:p1_raw;
  const xt_db=core2_power===0?-Infinity:core1_power===0?Infinity:10*Math.log10(core2_power/core1_power);
  const xt_label=xt_db===Infinity?'+∞ dB':xt_db===-Infinity?'−∞ dB':xt_db.toFixed(2)+' dB';
  return {pitch_um,length_mm,delta_neff,kappa_per_m,coupling_length_mm,
    core1_power,core2_power,xt_db,xt_label,xt_definition:'10log10(P2/P1)',
    model:'M8 legacy 2-core scalar FEM reference; log interpolation + lossless coherent coupled-mode equations',
    limitations:MCF_ASSUMPTIONS};
}
export function mcfLengthSeries(pitch_um,max_length_mm=100,count=201){
  if(!Number.isFinite(max_length_mm)||max_length_mm<=0||max_length_mm>100)throw new RangeError('XT 길이 스윕 범위 오류');
  if(!Number.isInteger(count)||count<2||count>2001)throw new RangeError('XT 스윕 점수 오류');
  return Array.from({length:count},(_,i)=>{
    const x=calculateMCF(pitch_um,max_length_mm*i/(count-1));
    return {length_mm:x.length_mm,core1_power:x.core1_power,core2_power:x.core2_power,
      xt_db:Number.isFinite(x.xt_db)?x.xt_db:null,
      xt_plot_db:Math.max(-60,Math.min(40,x.xt_db))};
  });
}
