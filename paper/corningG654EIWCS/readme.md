# G.654.E Paper Reproduction with `EGN_adaptive.py`

이 폴더는 **High Density Optical Cable with Ultra-Low-Loss, Large-Effective-Area ITU-T G.654.E Optical Fiber**의 G.654.E transmission condition을 기반으로 `EGN_adaptive.py` / `run_G654.py`의 launch-power–SNR 재현성을 검증합니다.

- 공식 출처: [IWCS Webinar 85](https://iwcs.org/webinar/85/)
- 참고 영상: [YouTube](https://www.youtube.com/watch?v=AVo-YAsVVFU&t=949s)
- 계산 엔진: [`EGN_adaptive.py`](./EGN_adaptive.py)
- 실행 스크립트: [`run_G654.py`](./run_G654.py)
- 실행 결과가 포함된 Colab 보고서: [`G654E_EGN_adaptive_validation_report.ipynb`](./G654E_EGN_adaptive_validation_report.ipynb)

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/paper/corningG654EIWCS/G654E_EGN_adaptive_validation_report.ipynb)

## 1. Validation purpose

핵심 검증 질문은 **논문/발표의 출력 SNR 값에 맞추기 위한 scale factor나 empirical fitting 없이**, 공개된 G.654.E fiber/system parameter를 `EGN_adaptive.py`에 입력했을 때 paper-reference launch-power/SNR 경향을 재현할 수 있는가입니다.

현재 `run_G654.py`는 G.654.E의 attenuation, effective area, dispersion, nonlinear refractive index와 90-channel / 95-GBd / 64QAM / 18-dB TRX SNR / 80-km span / 5-dB NF 조건을 직접 사용합니다. `γ`는 `n2`, `Aeff`, 1550-nm wavelength로부터 물리식으로 계산하며, paper 출력값을 맞추기 위한 correction factor는 넣지 않습니다.

> **Paper reference 정의**: 아래 비교 곡선은 기존 repository notebook [`practice_code/High_density_optical_cable_with_ultra_low_loss,_large_effective_area_ITU_T_G_654_E_Optical_fiber.ipynb`](./practice_code/High_density_optical_cable_with_ultra_low_loss,_large_effective_area_ITU_T_G_654_E_Optical_fiber.ipynb)에 저장된 figure-based fitted equation을 사용합니다. 원 논문에 인쇄된 폐형식 공식 자체라고 주장하지 않습니다.

## 2. Paper input vs `run_G654.py`

| Parameter | Paper / presentation | `run_G654.py` | Comparison |
|---|---:|---:|---|
| Fiber | ITU-T G.654.E | ITU-T G.654.E | Same |
| WDM channels | 90 | 90 | Same |
| Modulation | 64QAM | 64QAM | Same |
| Symbol rate | 95 GBd | 95 GBd | Same |
| Channel spacing | Nyquist spaced | 95 GHz | Same interpretation |
| TRX SNR | 18 dB | 18 dB | Same |
| Span length | 80 km | 80 km | Same |
| Number of spans | 1–50 spans stated | **30 spans fixed** | **Different** |
| EDFA noise figure | 5 dB | 5 dB | Same |
| Total gain bandwidth | 4.8 THz | **No explicit 4.8-THz filter** | **Different** |
| Shannon gap | 3 dB | 3 dB | Same; capacity calculation |
| Attenuation | 0.166 dB/km | 0.166 dB/km | Same |
| Effective area | 125 μm² | 125 μm² | Same |
| Dispersion | 21 ps/(nm·km) | 21 ps/(nm·km) | Same |
| Nonlinear refractive index `n2` | 2.2×10⁻²⁰ m²/W | 2.2×10⁻²⁰ m²/W | Same |
| Wavelength | Not shown in parameter box | 1550 nm | Code assumption |
| NLI path | Not specified in parameter box | `egn_sci` | Code assumption |
| Accumulation | Not specified in parameter box | coherent | Code assumption |

### Bandwidth note

90 channels × 95 GHz spacing corresponds to an occupied WDM span of approximately **8.55 THz**. The presentation box separately states **4.8 THz total gain bandwidth**. Since the current code does not impose a 4.8-THz amplifier/filter clipping function, bandwidth is intentionally treated as a remaining unmatched condition rather than silently fitted.

## 3. Paper reference vs Code

The existing paper-reference fit is

```text
SNR_paper(P) = 10 log10 { 1 /
  [0.02254·10^(-P/10) + 0.000817·10^(P/5) + 0.0158] }
```

This equation is used **only for comparison after the physics calculation**. These coefficients are not passed into `EGN_adaptive.py`.

![G.654.E Paper Reference vs EGN_adaptive.py](./g654e_paper_vs_code.svg)

## 4. Quantitative comparison

| Comparison range | MAE | RMSE | MAPE | Maximum absolute error |
|---|---:|---:|---:|---:|
| **−10 to +4 dBm** | **0.151 dB** | **0.186 dB** | **1.38%** | **0.500 dB** |
| −10 to +10 dBm | 0.749 dB | 1.328 dB | 6.35% | 3.761 dB |

At **+4 dBm**:

- Paper-reference: **15.239 dB**
- `EGN_adaptive.py`: **15.739 dB**
- Difference: **+0.500 dB**

Grid optimum values (1-dB launch-power step):

- Paper-reference: **+4 dBm, 15.239 dB**
- Code: **+6 dBm, 15.935 dB**

## 5. What this result demonstrates

The strongest result is not that every plotted point is identical. The important result is that the code uses the paper-reported fiber/transceiver parameters and **independently calculates ASE + NLI + TRX noise without using the paper output curve as a fitting target**. In the −10 to +4 dBm low-to-moderate launch-power region, the agreement is strong: **MAE ≈ 0.15 dB and MAPE ≈ 1.38%**.

This provides strong evidence that the implemented physical model reproduces the principal G.654.E launch-power/SNR behavior without output-target tuning.

At higher launch power, however, the code increasingly over-predicts SNR. Therefore this repository does **not** claim that the full −10 to +10 dBm curve is already perfectly reproduced.

Likely remaining contributors include:

1. fixed 30-span code condition versus the paper's broader 1–50-span statement,
2. the unresolved interpretation/application of the 4.8-THz amplifier gain bandwidth,
3. use of the `egn_sci` compatibility path for the 90-channel case rather than strict Full-EGN XCI/MCI correction,
4. ideal rectangular zero-rolloff spectrum and other numerical/model assumptions.

## 6. Conclusion

> **`EGN_adaptive.py` reproduces the paper-reference G.654.E launch-power/SNR trend without output-target fitting. With the paper-reported fiber and transceiver parameters, close agreement is obtained in the low-to-moderate launch-power region, supporting the physical validity and reproducibility of the implementation. High-power deviation remains and is explicitly retained rather than removed through empirical tuning.**

Accordingly, the current result can be used as evidence of **implementation validity and strong reproduction performance**, while strict Full-EGN validation should additionally reproduce XCI/MCI-sensitive reference cases and resolve the span/bandwidth condition differences.

## 7. Files

- [`EGN_adaptive.py`](./EGN_adaptive.py) — GN/EGN research engine
- [`run_G654.py`](./run_G654.py) — Python/IDLE G.654.E execution script
- [`G654E_EGN_adaptive_validation_report.ipynb`](./G654E_EGN_adaptive_validation_report.ipynb) — executed Colab-style validation report
- [`g654e_paper_vs_code.svg`](./g654e_paper_vs_code.svg) — Paper-reference vs Code comparison figure
- [`practice_code/`](./practice_code/) — previous reproduction notebooks
