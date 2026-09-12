# Code: GN/EGN NLI 계산 엔진
- Poggiolini GN 모델 관례를 따르는 범용 WDM 링크 NLI(비선형 잡음) 계산 엔진
`gn_integral_general.py`(핵심 적분 엔진)와 `gn_integral_general_modulation.py`(변조·시스템 성능 레이어) 두 모듈로 구성
### 1. `gn_integral_general.py` — GN 적분 엔진
- **Full WDM PSD 적분**: 불균일 채널/전력/간격 → SCI/XCI/MCI 자동 도출
- **스팬별 독립 설정**: 길이·손실·β2/β3·γ + Coherent/Incoherent 누적
- **증폭 모델**: Lumped EDFA / Ideal Distributed / Backward-Raman
- **QMC 적분**: Scrambled Sobol 기반 2D 주파수 적분
- **EGN SCI 보정**: μ4/μ6 기반, 옵트인 (XCI/MCI는 GN 유지)

**한계**: mid-span add/drop 미지원, EGN 보정은 SCI(자기채널)에만 적용되고 XCI/MCI는 항상 GN 유지되어 완전한 WDM EGN은 아님. 기본값은 항상 순수 GN과 동일하게 동작(self-test로 회귀 검증).

### 2. `gn_integral_general_modulation.py` — 성능 평가 레이어
- **변조 포맷**: BPSK~256QAM, 성상도 기반 정확한 Φ/Ψ 자동 계산
- **SNR/BER**: ASE+NLI 결합 GSNR, SNR_ASE, SNR_NLI, AWGN BER 근사
- **용량 추정**: Gross/Net rate, Shannon-gap capacity
- **고속 파워 스윕**: NLI P³ 법칙 이용, 1회 적분으로 GSNR 커브 산출

**한계**: `MODULATION_PHI`(EGN 초과첨도 Φ=μ4−2)와 `Channel.modulation_phi`(레거시 SCI 배율)는 이름은 비슷하나 다른 개념이라 혼용 주의. 8QAM/32QAM 등은 특정 성상 기하 가정에 기반한 근사치.


# 참고 논문
1. A Detailed Analytical Derivation of the GN Model of Non-Linear Interference in Coherent Optical Transmission Systems
- https://arxiv.org/abs/1209.0394
- 목적: 코히어런트 광전송 시스템에서 발생하는 비선형 간섭을 예측하는 GN-모델의 수학적 도출 과정과 상세한 이론적 근거 제공
- 핵심 가정: 광섬유 분산에 의해 NLI이 가우시안 잡음과 같은 통계적 특성을 띠며, ASE 잡음과 통계적으로 독립이라는 가정에 기반함
