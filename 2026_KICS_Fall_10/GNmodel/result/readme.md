# PSCF GN 모델 비교 결과

Carena 등 논문의 **그림 5(a), PSCF**에서 저자의 GN 모델 예측 실선과 두 TXT 코드의 실행 결과를 비교했습니다.

- **Paper:** 논문 실선을 이미지에서 판독한 근사값.
- **Code:** 두 TXT를 수정하지 않고 실행한 GN 계산값.
- **파일:** [PNG](./Carena_Fig5a_PSCF.png) · [PDF](./Carena_Fig5a_PSCF.pdf) · [비교값 CSV](./PSCF_solid_vs_TXT.csv).

계산 조건은 9채널, 32 GBd, 100 km 스팬, 감쇠 0.18 dB/km, 분산 20.1 ps/(nm·km), 비선형 계수 0.9 /(W·km), EDFA NF 5 dB, 비코히어런트 누적입니다. 그림 3의 요구 OSNR을 사용해 입사전력을 최적화하고 최대 전송거리를 계산했습니다.

21개 비교점의 **평균 절대 상대오차(MAPE)는 약 5.4%**입니다. 논문 실선 판독, 송신 PSD 및 직사각형 수신 대역의 근사를 포함하므로 절대 예측 정확도를 보증하는 결과는 아닙니다.

## 참고문헌

A. Carena et al., “Modeling of the Impact of Nonlinear Propagation Effects in Uncompensated Optical Coherent Transmission Links,” *Journal of Lightwave Technology*, 30(10), 1524–1539 (2012). [논문](https://ieeexplore.ieee.org/document/6158564) · DOI: 10.1109/JLT.2012.2189198.
