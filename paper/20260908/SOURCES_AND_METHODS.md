# 20260908 — 입력 출처, 구현식, 검증 범위

## 비교 기준

- Marco Petrovich et al., **Broadband optical fibre with an attenuation lower than 0.1 decibel per kilometre**, Nature Photonics **19**, 1203–1208 (2025). [DOI](https://doi.org/10.1038/s41566-025-01747-5).
- 비교 섬유는 **HCF2**. 본문 Fig. 2d의 cutback 평균은 1310 nm에서 0.128 dB/km, 1550 nm에서 0.091 dB/km. OTDR 값 0.123/0.095와 혼합하지 않았다. Fig. 1의 보정용 별도 섬유나 HCF1도 사용하지 않았다.
- [Fig. 3 Source Data 원본 XLSX](https://media.springernature.com/original/springer-static/esm/art%3A10.1038%2Fs41566-025-01747-5/MediaObjects/41566_2025_1747_MOESM3_ESM.xlsx)의 `HCF_cutback_loss` 시트를 사용했다. 원본은 `data/petrovich_fig3_source.xlsx`, 추출본은 `data/petrovich_hcf2_cutback_source.csv`에 보존했다. 같은 통합문서의 `gas_free_HCF_loss`는 가스 기여를 제거한 파생 곡선이므로 주 비교에 사용하지 않았다.
- XLSX SHA256: `e74c7cd115146dc93d254c0d26f3f991dd76e7fdb7c52cf7e8271e930f3bb31e`.
- 원자료 1799점: 1200–1919.2 nm, 간격 0.4 nm. 주 다파장 평가는 **1200–1650 nm, 10 nm 간격, 46점**이다. 모든 평가 파장은 원자료에 정확히 존재하며 보간하거나 그림을 digitize하지 않았다. 이 범위는 네 cutback 평균을 보고한 측정 범위에 해당한다. 원자료의 두 점은 0.128206 및 0.090698975 dB/km로, 본문의 반올림값과 별도 집계한다.
- [2025-10-17 정정문](https://doi.org/10.1038/s41566-025-01803-0)은 Fig. 2c의 x축 1200–1500 nm를 1200–1900 nm로 수정했다. 1310/1550 nm 손실값을 바꾸는 정정은 아니다.
- 이 CSV는 원자료 추출물이며 Petrovich et al.에 귀속된다. 원문에 명시된 [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/)에 따라 출처를 표시한다. 원문 전체나 논문 그림은 재배포하지 않았다.

## 계산에 사용한 기하와 가정

| 입력 | 값 (µm) | 근거/한계 |
|---|---:|---|
| 코어 반경 | 14.75 | 논문의 HCF2 nominal diameter 29.5 및 기존 코드 유지; 측정 지름 범위 29.1–29.6 |
| 큰 튜브 지름 | 31.05 | 공개 범위 30.4–31.7의 중간값 |
| 중간 튜브 지름 | 23.75 | 공개 범위 22.7–24.8의 중간값 |
| 작은 튜브 지름 | 7.70 | 공개 범위 7.0–8.4의 중간값 |
| 각 유리막 두께 | 0.50 | 본문 약 500 nm를 모든 막에 동일 적용; 개별 막 측정값 아님 |
| 환산 공기층 g1 | 6.30 | 큰 지름 − 2t − 중간 지름; 튜브가 같은 반경 방향에서 내부 접한다고 가정 |
| 환산 공기층 g2 | 15.05 | 중간 지름 − 2t − 작은 지름; 실제 측정 gap이 아님 |
| 공기 굴절률 | 1 | 진공 근사; 잔류 가스의 압력/조성 및 흡수는 미지 |
| 실리카 굴절률 | 파장별 실수 Sellmeier | 기존 코드 유지. Malitson, JOSA 55, 1205–1209 (1965), [DOI](https://doi.org/10.1364/JOSA.55.001205); 재료 분산 데이터에 대한 문헌식이며 HCF2 손실 fitting은 아님 |

27개 튜브 지름 조합은 각 공개 범위의 하한/중간/상한의 직교 조합이다. 실제 함께 관측된 단면 27개도, 확률분포나 신뢰구간도 아니다. 손실 오차가 가장 작은 조합을 고르지 않는다. 막 두께 분포, 타원도, 접촉 위치, 종방향 상관관계의 불확실성은 이 범위 계산으로 해결되지 않는다.

## 논문별 구현 위치와 적용 한계

1. **David Bird**, *Attenuation of model hollow-core, anti-resonant fibres*, Optics Express **25**, 23215–23237 (2017), [DOI](https://doi.org/10.1364/OE.25.023215). [출판사 제공 전체 원문 미러](https://www.researchgate.net/publication/319695985_Attenuation_of_model_hollow-core_anti-resonant_fibres).
   - Eqs. (4)–(5): 각 영역의 종방향 전기/자기장과 복소 횡파수.
   - Eqs. (6b), (6d), (7)–(8): 접선장 연속 조건. `hcf_vector_tmm._basis`에서 Ez, Z0Hz, Ephi, Z0Hphi 네 성분을 동시에 연결한다. exp(i*m*phi) 표기는 논문의 실수 sin/cos 각도 기저와 동등하다.
   - Eq. (25) 계열의 다층 전달: `reduced_matching`은 외부 outgoing Hankel 기저를 안쪽으로 전파한다. QR은 수치 크기 조절만 수행한다. 정확한 Bessel/Hankel 함수를 사용하고 Eq. (26)의 점근근사는 사용하지 않는다.
   - Eq. (3) 기하 및 **Table 1, HE11, rc/lambda=15, epsilon=2.25, N=1–4**의 정확 수치값을 `bird_benchmark`에서 직접 재계산한다. 논문의 N은 유한층 수로, 본 보고서의 전체 영역 수와 다르다. 표는 소수점 셋째 자리로 반올림되어 있으므로 ±0.0005 허용 폭도 보고한다. Table 2의 손실 최소화 치수는 사용하지 않는다.
   - 동심 구조의 일치는 실제 비원형 5-tube double-nested DNANF의 검증이 아니다.
2. **Morten Bache, Md. Selim Habib, Christos Markos, Jesper Lægsgaard**, *Poor-man’s model of hollow-core anti-resonant fibers*, JOSA B **36**, 69–80 (2019), [DOI](https://doi.org/10.1364/JOSAB.36.000069), [공개 원문](https://arxiv.org/abs/1806.10416).
   - 공개 원문 Eqs. (15), (16), (17)의 TE/TM 손실 및 hybrid 평균은 기존 `bache_bouncing_ray_loss_db_km`로 계산한다. 단일막의 편광 처리 점검에 사용한다.
   - 논문의 FEM 맞춤 스케일 `f_FEM`은 사용하지 않는다. FEM fitting을 포함한 poor-man 모델 전체를 무보정 예측이라고 부르지 않는다. 중첩막에 단일막 계수를 반복 곱하는 임의 손실식도 주 결과로 사용하지 않는다.
3. **Petrovich et al.**의 Methods: 실제 SEM 단면을 이용한 FEM leakage, 표면 산란 및 미세굽힘 등 총손실 항의 정의를 확인한다. 저자들이 15개 섬유에 대해 추정한 경험계수를 이 코드에 이식하지 않는다. 미공개 표면 PSD/굽힘 조건을 0으로 측정되었다고 간주하지 않는다.
4. **Leah R. Murphy & David Bird**, *Azimuthal confinement: the missing ingredient in understanding confinement loss in antiresonant, hollow-core fibers*, Optica **10**, 854–870 (2023), [DOI](https://doi.org/10.1364/OPTICA.492058), [저자 기관 원문 설명](https://researchportal.bath.ac.uk/en/publications/azimuthal-confinement-the-missing-ingredient-in-understanding-con/).
   - 다음 단계의 구조 개선 근거. 방위각 방향 장 구속의 중요성을 설명하므로 동심원 단순화의 한계를 해석하는 데 참고한다. 단일 반공진 유리층에 대해 검증된 식을 double-nested HCF2에 임의 적용하거나 감소 배율을 가져오지 않았다.

## 실제 수정과 검증의 의미

- 기존 `hcf_concentric_tmm_20260908.py`의 scalar 5영역 코드를 그대로 실행해 수정 전 기준으로 사용했다. 이전 scalar 7영역 진단도 다시 실행한다.
- 수정 주 결과는 **vector 7영역**이다. 공개된 세 중첩막을 기존 동심 근사 안에서 유지하면서 scalar m=0의 장/미분 연속을 **m=1 Maxwell 네 접선장 연속 조건**으로 바꿨다. 이는 물리 모델의 수정이며, 실측에 더 가까운 모델을 파장마다 선택한 결과가 아니다.
- 시간/전파 convention은 exp(i beta z − i omega t). Im(beta)>0인 수동 감쇠 해만 허용한다. 손실 변환은 `2*Im(beta)*10/ln(10)`에 길이 단위를 적용한다. 잘못된 부호에 절댓값을 씌우거나 beta를 켤레화해서 통과시키지 않는다.
- 안쪽/바깥쪽 전달의 최소 특이값과 40자리 별도 mpmath 구현으로 1310/1550 nm 근을 검산한다. 이는 같은 Maxwell 문제의 수치 검산이다. Bird 표는 별도 출판 수치 기준이다. 어느 것도 HCF2 FEM 데이터의 대체 검증이 아니다.
- **vector 9영역**은 별도 형상 민감도 진단이다. 동일 선상 내부 접촉을 가정한 세 튜브의 완전한 직경 경로를 유지하면 작은 튜브 내부 공기 6.70 µm와 반대쪽 접촉 유리벽 3t=1.50 µm가 추가된다. 실제 단면을 다시 동심원으로 바꾼 가정이므로 주 결과로 승격하지 않는다. 이 두꺼운 환산 벽의 공진은 인공적인 동심 구조 특성일 수 있다. 특정 파장에서 더 작은 오차가 나와도 선택하지 않는다.
- 알려진 1310/1550 nm 값을 본 뒤 세운 과거 보정 모델은 가져오지 않았다. 기존 개선 노트북에는 목표값을 기저 행렬로 풀어 정한 계수가 있었으므로 독립 계산으로 취급할 수 없다. 이번 예측 함수에는 논문 손실값 입력이 없다. 다만 논문값 자체는 이전 대화에서 이미 알려져 있으므로 완전한 blind 검증은 아니다.

## 10% 초과 시 개선 순서

| 우선순위 | 확인된 사실 / 남은 추정 | 바꿀 코드와 필요한 자료 |
|---|---|---|
| 1 | 동심 모델에는 실제 5개 튜브의 방위각 형상이 없음. Bird 검산 성공은 그 차이를 제거하지 않음 | `geometry_models` 및 Maxwell solver를 실제 2D 단면 기반으로 대체. HCF2의 여러 위치 SEM 윤곽, 개별 막 두께, gap/접촉/타원도, outer jacket 필요. Full-vector FEM/PML에서 mesh, 영역 크기, PML 수렴과 동일 모드 overlap 추적을 수행 |
| 2 | scalar 결과와 vector 결과의 차이가 큼. 모드와 편광 처리는 이미 수정했지만 실제 단면의 cladding-mode coupling은 미계산 | m=1 원통 가정을 해제하고 실제 두 편광과 core/cladding mode overlap, 공진 근처 모드 혼합을 검사. 손실이 가장 작은 고유치만 선택하지 않음 |
| 3 | 계산한 leakage가 실측 total보다 이미 큼. 양의 SSL/µBL/gas 항을 더해서 이 과대예측을 줄일 수 없음 | leakage 구조 검증 후에만 항별 합계 계산. 독립 표면 거칠기 PSD, 코팅/직경/굽힘 PSD 및 설치 조건, 가스 조성/압력/온도 필요. Petrovich 경험계수를 복사해 무보정이라고 표시하지 않음 |
| 4 | 지름 범위만 있고 각 위치의 공동 분포와 막 두께 오차는 없음 | 범위 표를 최적값 탐색으로 쓰지 말고 실제 longitudinal samples로 전파손실 평균 및 불확실성 평가 |

실제 HCF2 총손실 또는 FEM을 대체할 수 있다는 결론은 내릴 수 없다. 계산은 누설손실이므로 실측 총손실과의 백분율은 **총손실 예측 정확도가 아니라 물리량이 다른 두 값의 차이 진단**이다. 다음 버전이 10% 이내라고 보장할 근거는 없다.
