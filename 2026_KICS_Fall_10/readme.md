# Code
- gn_integral.py = GN 적분형 범용 코어 모델
- 참고 논문: A Detailed Analytical Derivation of the GN Model of Non-Linear Interference in Coherent Optical Transmission Systems
- 특징
  1) 분산: β₂만 고려
  2) Span: 동일 스팬(손실을 매 Span 끝단에서 정확히 보상하는 lumped EDFA)
  3) Channel: 단일 채널(flat-top PSD 가정)
  4) 간섭 성분: SCI(채널이 하나 뿐이므로, XCI나 MCI 존재하지 않음)
  5) 계산 방식: 2층 적분을 수치적으로 직접 계산

# 참고 논문
1. A Detailed Analytical Derivation of the GN Model of Non-Linear Interference in Coherent Optical Transmission Systems
- https://arxiv.org/abs/1209.0394
- 목적: 코히어런트 광전송 시스템에서 발생하는 비선형 간섭을 예측하는 GN-모델의 수학적 도출 과정과 상세한 이론적 근거 제공
- 핵심 가정: 광섬유 분산에 의해 NLI이 가우시안 잡음과 같은 통계적 특성을 띠며, ASE 잡음과 통계적으로 독립이라는 가정에 기반함
