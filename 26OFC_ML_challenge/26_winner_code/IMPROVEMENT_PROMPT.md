# 구현 요청 프롬프트

기존 `26OFC_ML_challenge/PI_EDFA_Transformer`를 보존하고 별도 실험 폴더에 EDFA 95채널 예측 실험을 구현하라. 목표는 SJTU 논문의 COSMOS 사전학습 및 대회 데이터 미세조정과 FourierKAN/은닉축 spectral mixing의 효과를 공정하게 검증하는 것이다. 우승 성능을 보장하거나 점수를 임의 보정하지 말 것.

1. 데이터: 공식 대회 데이터 및 COSMOS JSON→동일 schema 변환, 00–94 정렬·단위·WSS mask·불량 측정 확인. 파일 경로 CLI 및 Colab Drive 지원. Train/val/test를 **먼저** 분할하고 scaler/encoder는 train에서만 fit. 원본 샘플/장치/시간·파일 기준 중복 제거 및 group-aware 분할. 적용 가능한 group이 없으면 명시하고 random split은 별도 보고.
2. 모델: 기존 Transformer 기준선을 그대로 실행 가능하게 유지. 동일 gain/tilt 기준선·mask·특성을 공유하고 FourierKAN + 은닉 특성 FFT mixing 실험 모델 추가. 각 모델의 출력은 기준선+잔차. 파장순서 및 tilt 부호 검증. 입력 총 출력파워의 추론 시 가용성 확인, 없는 경우 대조 실험.
3. 학습: COSMOS 사전학습→대회 train만 fine-tune, from-scratch 대조군, 체크포인트/재개·optimizer·scheduler·seed 및 환경 기록, CPU/T4 지원. 외부 및 대회 간 중복 방지. 활성 채널 없는 샘플은 명시적으로 제외/보고.
4. 평가: 동일 고정 holdout 및 공식 scorer로 score·MAE·RMSE·bias·p95·최대오차와 Aging/SHB/Unseen별 결과를 표·CSV·그래프·JSON으로 저장. 공식 TestGround는 최종 보고용으로 유지하고 반복 튜닝에 쓰지 말 것. 대회 비공개 점수와 직접 비교하지 말 것.
5. 실험표: Transformer vs FourierKAN, 사전학습 유무, spectral mixing 유무, masked MSE vs 점수 대리손실을 각각 통제해 3개 이상 seed로 평균·표준편차 보고. 같은 data split과 budget을 사용하고 불가능한 실험은 미실행으로 기록.
6. 테스트: 기준선·mask·채널 순서·split 겹침·scaler 누출·checkpoint 재개·제출 스키마를 검증. 실제 실행한 결과와 미검증 항목을 구별하여 README에 명시. 점수 개선이나 0.09258 재현을 보장하지 말 것.
