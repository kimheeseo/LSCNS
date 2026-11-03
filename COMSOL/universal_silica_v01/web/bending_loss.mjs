/**
 * Macrobend measured-anchor estimator, NOT curved-fiber FEM or a physical
 * prediction from MFD/Aeff. Four positive dB/turn measurements:
 *   radii rSmall/rLarge at reference wavelengths 1550/1625 nm.
 * Log(loss_per_turn) is interpolated linearly in 1/R and in wavelength.
 * No radius/wavelength extrapolation. Loss/turn × turns assumes additive wraps.
 */
export function estimateBending(input){
 const {radiusSmallMm,radiusLargeMm,lossSmall1550DbTurn,lossLarge1550DbTurn,
  lossSmall1625DbTurn,lossLarge1625DbTurn,radiusMm,wavelengthNm,turns,
  straightLinkLossDb}=input;
 const values=Object.values(input);
 if(values.some(v=>!Number.isFinite(v)))throw Error('모든 굽힘 조건에 유효한 숫자를 입력하세요.');
 if(!(radiusSmallMm>0&&radiusLargeMm>radiusSmallMm&&radiusLargeMm<=1000))
  throw Error('반경 기준은 0 < 작은 반경 < 큰 반경 ≤1000 mm여야 합니다.');
 if(radiusMm<radiusSmallMm||radiusMm>radiusLargeMm)
  throw Error('측정 반경 범위 밖에서는 굽힘손실 외삽을 허용하지 않습니다.');
 if(wavelengthNm<1550||wavelengthNm>1625)
  throw Error('측정 파장 1550~1625 nm 범위 밖에서는 외삽을 허용하지 않습니다.');
 if(!Number.isInteger(turns)||turns<0||turns>100000)
  throw Error('권선 수는 0 이상의 정수여야 합니다.');
 if(straightLinkLossDb<0||straightLinkLossDb>100000)
  throw Error('직선부 링크 손실은 0 이상의 dB로 입력하세요.');
 const small=[lossSmall1550DbTurn,lossSmall1625DbTurn],large=[lossLarge1550DbTurn,lossLarge1625DbTurn];
 if([...small,...large].some(v=>v<=0||v>100))
  throw Error('보간용 측정 굽힘손실은 0보다 큰 100 dB/turn 이하의 값이어야 합니다. 검출한계 이하 측정값(0)은 그대로 로그 보간할 수 없습니다.');
 const invSmall=1/radiusSmallMm,invLarge=1/radiusLargeMm;
 const tRadius=(1/radiusMm-invSmall)/(invLarge-invSmall),tLambda=(wavelengthNm-1550)/75;
 const atWave=(j)=>Math.exp(Math.log(small[j])+(Math.log(large[j])-Math.log(small[j]))*tRadius);
 const at1550=atWave(0),at1625=atWave(1);
 const lossPerTurnDb=Math.exp(Math.log(at1550)+(Math.log(at1625)-Math.log(at1550))*tLambda);
 const bendLossDb=turns*lossPerTurnDb;
 return {model:'measured-anchor-log-curvature-and-wavelength-interpolation',
  lossPerTurnDb,bendLossDb,straightLinkLossDb,totalLinkLossDb:straightLinkLossDb+bendLossDb,
  bendLengthM:turns*2*Math.PI*radiusMm/1000,source:'user-entered or illustrative FOUR measured bend-loss anchors',
  input:{...input},interpolation:{radiusFraction:tRadius,wavelengthFraction:tLambda},
  limitations:['Not curved-mode FEM/PML radiation prediction','Do not extrapolate beyond measured radius or wavelength','No turn-to-turn coupling, tension/coating/thermal variability']};
}
