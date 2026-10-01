# 수치 적분 기반 GN 모델 — 2026_KICS_Fall_10

코히어런트 WDM 광전송 링크의 NLI를 수치 적분하고, ASE·GSNR·BER·전송률 및 용량을 계산하는 연구용 Python 모듈입니다.

## 파일과 기능

| 파일 | 주요 기능 |
|---|---|
| [gn_integral_general.py](./gn_integral_general.py) | Sobol QMC 기반 전체 WDM GN 적분, NLI PSD 및 수신 대역 내 NLI 전력 계산 |
| [gn_integral_general_modulation.py](./gn_integral_general_modulation.py) | 변조 모멘트, EDFA ASE, GSNR·근사 BER, 입사전력 탐색, DP 전송률·Shannon-gap 용량 |

GN 엔진은 비균일 채널 간격·전력·대역폭, 직사각형·RRC·상승 코사인·사용자 정의 PSD, 스팬별 광섬유 파라미터, beta2/beta3, 코히어런트·비코히어런트 누적을 지원합니다. 증폭 방식은 집중형 EDFA, 이상적인 분포형 증폭, 단순화한 역방향 Raman 증폭 및 사용자 정의 프로파일입니다.

지원 변조: **BPSK, QPSK, 8QAM, 16QAM, 32QAM, 64QAM, 256QAM**. 변조 모멘트는 코드에 정의된 신호점 배열에서 직접 계산합니다.

## 입력과 출력

- `WDMSystem` / `Channel`: 채널 수·간격, 심볼률, 채널 전력, PSD 형태.
- `Span`: 길이, 감쇠, 분산(`D` 또는 `beta2`), `gamma`, 파장, 증폭 이득·NF.
- `GNIntegralOptions`: `sobol_power`(표본 수 = 2의 해당 거듭제곱), 난수 시드, 누적 방식, 적분 차수.
- `evaluate_performance()`: 한 운용점의 ASE·NLI 전력, GSNR·BER, 총 전송률·순 전송률 및 용량.
- `launch_power_vs_gsnr()`: 입사전력별 결과 배열. 기본적으로 한 번의 NLI 적분과 `P_NLI ∝ P_ch³` 비례 관계를 사용합니다.

단위: 주파수 **THz**, 거리 **km**, 채널 전력 **W**(두 편광의 합계), 감쇠 **dB/km**(광전력 손실), 분산 **ps/(nm·km)**, `gamma` **1/(W·km)**. NLI PSD는 **W/THz**, 전송률·용량은 **Gb/s**입니다.

## 설치 및 사용

Python 3.10 이상에서 두 파일을 같은 폴더에 두고 사용합니다.

```bash
pip install numpy scipy
```

아래 코드를 같은 폴더에서 실행하면 채널당 최적 입사전력과 GSNR을 구합니다.

```python
import numpy as np
from gn_integral_general import WDMSystem, Span, GNIntegralOptions
from gn_integral_general_modulation import launch_power_vs_gsnr

system = WDMSystem.equispaced(
    n_channels=9, spacing_GHz=50.0, baud_GBd=32.0,
    power_dBm=0.0, pulse_shape="rect",
)
spans = [Span(
    length_km=80.0, alpha_db_per_km=0.20,
    gamma_W_inv_km=1.3, D_ps_nm_km=17.0,
    noise_figure_db=5.0,
) for _ in range(10)]

result = launch_power_vs_gsnr(
    base_system=system, spans=spans, cut_index=4,
    launch_power_dBm=np.arange(-6.0, 4.1, 0.5),
    modulation="QPSK", trx_snr_db=18.0,
    gn_options=GNIntegralOptions(
        sobol_power=18, seed=1, accumulation="coherent",
    ),
    receiver_points=7, nli_model="gn",
)
i = int(np.argmax(result["gsnr_db"]))
print(f"탐색 지점 중 최적 입사전력: {result['launch_power_dBm'][i]:.2f} dBm")
print(f"GSNR: {result['gsnr_db'][i]:.2f} dB")
```

## 모델 적용 범위

- 기본값은 `nli_model="gn"`입니다. `nli_model="egn_sci"`는 SCI에만 EGN 보정을 적용하며, XCI/MCI는 GN으로 계산합니다. 이 옵션은 직사각형 스펙트럼, 코히어런트 누적, CUT와 다른 채널의 스펙트럼이 겹치지 않는 조건이 필요합니다.
- `MODULATION_PHI = mu4 - 2`는 부호가 있는 EGN 계수입니다. 기존 SCI 배율인 `Channel.modulation_phi`(비음수)와 구분해야 합니다.
- BER은 오류 정정 부호를 적용하지 않은 AWGN 근사값입니다. 8QAM/32QAM은 코드에 정의된 신호점 배열을 사용하며, BER은 일반적인 QAM 근사식으로 계산합니다.
- 총 전송률과 순 전송률은 변조 차수와 부호율로 계산합니다. Shannon-gap 용량은 추정값으로, 실장비의 달성 용량이나 FEC 이후 BER을 나타내지 않습니다.
- 수치 수렴은 `sobol_power`, 난수 시드, `receiver_points`, EGN 적분 차수를 바꾸어 확인합니다. ASE 계산 함수는 집중형 EDFA용이며, 분포형 Raman 증폭의 ASE는 포함하지 않습니다.
- 링크 중간의 채널 추가·삭제, PMD/PDL, 레이저 위상 잡음, DSP 페널티는 모델에 포함하지 않습니다.

전체 WDM EGN의 별도 구현과 검증 자료는 [EGN_model](./EGN_model/), GN 검증 자료는 [GN_model](./GN_model/)을 참조하세요.

## 참고문헌

1. P. Poggiolini, “The GN Model of Non-Linear Propagation in Uncompensated Coherent Optical Systems,” JLT 30, 3857–3879 (2012).
2. P. Poggiolini et al., “A Detailed Analytical Derivation of the GN Model of Non-Linear Interference in Coherent Optical Transmission Systems,” arXiv:1209.0394.
3. A. Carena et al., “Modeling of the Impact of Nonlinear Propagation Effects in Uncompensated Optical Coherent Transmission Links,” JLT 30, 1524–1539 (2012).
4. A. Carena et al., “EGN model of non-linear fiber propagation,” Optics Express 22, 16335–16362 (2014), DOI: 10.1364/OE.22.016335.
