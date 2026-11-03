/**
 * Attenuation engineering estimator — measured/specification anchors only.
 * Interpolation is strictly bounded to the two entered wavelengths; not a
 * first-principles Rayleigh/IR absorption prediction.
 * Units: alpha dB/km, length km, launch and output power dBm.
 */
export function estimateAttenuation(input) {
 const {alpha1550,alpha1625,wavelengthNm,lengthKm,launchDbm,
        connectorCount,connectorDb,spliceCount,spliceDb}=input;
 const values=[alpha1550,alpha1625,wavelengthNm,lengthKm,launchDbm,connectorCount,connectorDb,spliceCount,spliceDb];
 if(values.some(v=>!Number.isFinite(v)))throw Error('모든 감쇠 항목에 유효한 숫자를 입력하세요.');
 if(alpha1550<0||alpha1625<0||alpha1550>20||alpha1625>20)throw Error('감쇠 계수는 0~20 dB/km 이내로 입력하세요.');
 if(wavelengthNm<1550||wavelengthNm>1625)throw Error('두 기준 파장 1550~1625 nm 사이에서만 보간할 수 있습니다. 외삽을 제공하지 않습니다.');
 if(lengthKm<0||lengthKm>20000)throw Error('전파 길이는 0~20,000 km로 입력하세요.');
 if(!Number.isInteger(connectorCount)||!Number.isInteger(spliceCount)||connectorCount<0||spliceCount<0||
   connectorCount>50000||spliceCount>50000)throw Error('접속부 개수는 0 이상의 정수여야 합니다.');
 if(connectorDb<0||spliceDb<0||connectorDb>30||spliceDb>30)throw Error('접속당 손실은 0~30 dB여야 합니다.');
 const fraction=(wavelengthNm-1550)/75;
 const alphaDbKm=alpha1550+(alpha1625-alpha1550)*fraction;
 const fiberLossDb=lengthKm*alphaDbKm;
 const connectorLossDb=connectorCount*connectorDb;
 const spliceLossDb=spliceCount*spliceDb;
 const totalLossDb=fiberLossDb+connectorLossDb+spliceLossDb;
 return {model:'measured-anchor-linear-interpolation',source:'user-entered example/spec/measurement, provenance must be verified',
  alphaDbKm,fiberLossDb,connectorLossDb,spliceLossDb,totalLossDb,
  outputDbm:launchDbm-totalLossDb,receiverFraction:10**(-totalLossDb/10),
  input:{...input},interpolationFraction:fraction,
  limitations:['No extrapolation outside 1550–1625 nm','No intrinsic material-loss prediction','No bend/PMD included']};
}
