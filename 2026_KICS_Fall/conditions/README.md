# KICS Fall 2026 — GN 계산 기준 조건

이 폴더는 2026 KICS 추계학회용 GN 계산의 **기준 입력 조건**을 보관합니다. [gn_input_conditions.png](./gn_input_conditions.png)는 아래 표와 같은 조건을 시각적으로 정리한 자료입니다.

| 항목 | 기준값 |
|---|---:|
| WDM 채널 수 | 50 |
| 심볼률 / 채널 간격 | 95 GBd / 95 GHz |
| 채널별 launch power | 0 dBm/ch |
| span 길이 / 수 | 80 km / 30 |
| 총 거리 | 2,400 km |
| 파장 | 1550 nm |
| 비선형 굴절률 (n_2) | 2.2 × 10⁻²⁰ m²/W |
| 감쇠 / 분산 | 0.146 dB/km / 22 ps/(nm·km) |
| NLI span 누적 | Incoherent |
| 송수신기 SNR | 18 dB |
| Sobol 표본 수 | 2¹⁸ |
| Sobol seed | 1, 2, 3 |
| CUT | 24 (0-based index) |
| 수신대역 NLI 적분점 | 7 |

## 재현 규칙

1. seed 1·2·3을 각각 계산합니다.
2. 각 결과의 **NLI 전력(W)을 선형 단위로 평균**합니다. dB 값의 산술 평균은 사용하지 않습니다.
3. TRX SNR 18 dB는 GSNR 계산에만 포함하며, GN 적분으로 구하는 (P_{NLI})에 직접 더하지 않습니다.
4. 코드가 (gamma)를 요구할 경우, 선택한 광섬유의 (n_2), 유효면적 (A_{eff}), 파장으로 일관되게 변환해 사용합니다.

이 조건은 [상위 GN 엔진](../readme.md)의 일반 실행 예제보다 우선합니다.
