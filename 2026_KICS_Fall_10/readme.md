# Numerical GN Model — 2026_KICS_Fall_10

코히어런트 WDM 광전송 링크의 NLI를 수치 적분하고, ASE·GSNR·BER·전송률 및 용량을 계산하는 연구용 Python 모듈입니다.

## 파일과 기능

| 파일 | 주요 기능 |
|---|---|
| [gn_integral_general.py](./gn_integral_general.py) | Sobol QMC 기반 전체 WDM GN 적분, NLI PSD 및 수신 대역 내 NLI 전력 계산 |
| [gn_integral_general_modulation.py](./gn_integral_general_modulation.py) | 변조 moments, EDFA ASE, GSNR·근사 BER, 입사전력 sweep, DP 전송률·Shannon-gap 용량 |

GN 엔진은 비균일 채널 간격·전력·대역폭, rectangular/RRC/raised-cosine/custom PSD, span별 광섬유 파라미터, beta2/beta3, coherent/incoherent 누적을 지원합니다. 증폭 방식은 lumped EDFA, ideal distributed, simplified backward Raman 및 custom profile입니다.

지원 변조: **BPSK, QPSK, 8QAM, 16QAM, 32QAM, 64QAM, 256QAM**. 변조 moments는 코드에 정의된 constellation에서 직접 계산합니다.

## 입력과 출력

- `WDMSystem` / `Channel`: 채널 수·간격, baud rate, 채널 전력, PSD 형태.
- `Span`: 길이, 감쇠, 분산(`D` 또는 `beta2`), `gamma`, 파장, 증폭 이득·NF.
- `GNIntegralOptions`: `sobol_power`(표본 수 = 2의 해당 거듭제곱), seed, 누적 방식, 적분 차수.
- `evaluate_performance()`: 한 운용점의 ASE·NLI 전력, GSNR·BER, gross/net rate 및 용량.
- `launch_power_vs_gsnr()`: 입사전력별 결과 배열. 기본적으로 한 번의 NLI 적분과 `P_NLI ∝ P_ch³` scaling을 사용합니다.

단위: 주파수 **THz**, 거리 **km**, 채널 전력 **W**(총 dual-polarization), 감쇠 **dB/km**(power loss), 분산 **ps/(nm·km)**, `gamma` **1/(W·km)**. NLI PSD는 **W/THz**, 전송률·용량은 **Gb/s**입니다.

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
print(f"Best sampled launch power: {result['launch_power_dBm'][i]:.2f} dBm")
print(f"GSNR: {result['gsnr_db'][i]:.2f} dB")
```

## 모델 적용 범위

- `nli_model="gn"`が基本です。`nli_model="egn_sci"`はSCIのみをEGN補正し、XCI/MCIはGNのままです。矩形スペクトル・coherent累積・CUTと他チャネルの非重複が必要です。
- `MODULATION_PHI = mu4 - 2`は符号付きEGN係数です。非負の旧SCI倍率`Channel.modulation_phi`とは異なります。
- BERはuncoded AWGN近似です。8QAM/32QAMのconstellationとBER近似にはモデル選択が含まれます。
- gross/net rateは変調次数とcoding rateから計算します。Shannon-gap容量は推定値で、実機の達成容量やFEC後BERを表しません。
- 数値収束は`sobol_power`、seed、`receiver_points`、EGN積分次数を変えて確認してください。ASE helperはlumped EDFA用で、distributed Raman ASEは含みません。
- mid-link add/drop、PMD/PDL、レーザ位相雑音、DSPペナルティはモデルに含みません。

Full-WDM EGNの別実装と検証資料は[EGN_model](./EGN_model/)、GN検証資料は[GN_model](./GN_model/)を参照してください。

## 参考文献

1. P. Poggiolini, “The GN Model of Non-Linear Propagation in Uncompensated Coherent Optical Systems,” JLT 30, 3857–3879 (2012).
2. P. Poggiolini et al., “A Detailed Analytical Derivation of the GN Model of Non-Linear Interference in Coherent Optical Transmission Systems,” arXiv:1209.0394.
3. A. Carena et al., “Modeling of the Impact of Nonlinear Propagation Effects in Uncompensated Optical Coherent Transmission Links,” JLT 30, 1524–1539 (2012).
4. A. Carena et al., “EGN model of non-linear fiber propagation,” Optics Express 22, 16335–16362 (2014), DOI: 10.1364/OE.22.016335.
