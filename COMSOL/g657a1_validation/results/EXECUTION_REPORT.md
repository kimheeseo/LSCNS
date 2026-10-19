# G.657.A1 Colab 실행 보고서

- MFD 보정 목표: 9.200 um @1310 nm; 코어 반경 4.1 um는 가정
- 보정 Delta n: 0.0046215295; 분석 파장 1310.0 nm
- MFD PII @1550 nm: 10.543027 um; Aeff1550: 84.459523 um2
- D1550: 16.641833; D1625: 20.687858 ps/(nm km)
- lambda0: 1303.2814 nm; S0: 0.0840416 ps/(nm2 km)
- gamma1550: 1.247880 1/(W km), assumed n2=2.6e-20 m2/W
- 수치 검사: 11/12 PASS
- 규격 수치 참고: 4/4 within reference ranges
- 독립 제품 검증: NONE. Bending loss, attenuation, actual profile and cable cutoff are NOT validated.

## 수치 검사
- PASS: Radial |neff error| <1e-6
- PASS: Radial |Aeff error| <1%
- PASS: Radial |second-moment MFD error| <1%
- PASS: Radial mesh error decreases at every refinement
- PASS: Radial residual <1e-8
- PASS: Fixed-spacing domain neff change <1e-7
- PASS: Fixed-spacing domain Aeff change <0.1%
- PASS: D1550 derivative-step spread <0.01
- FAIL: 2D |neff error| <1e-5
- PASS: 2D |Aeff error| <2%
- PASS: 2D |second-moment MFD error| <2%
- PASS: 2D residual <1e-8
