# PI-EDFA Transformer: 95채널 관계 학습 실험

기존 [PI-EDFA](../OFC_2026_Physics_Guided_EDFA_Colab.ipynb)를 보존하면서 별도 모델을 실험한 폴더입니다. [transformer_edfa.py](transformer_edfa.py)는 단독 실행 가능한 Python 스크립트입니다.

## 구조

```text
목표 gain + 학습된 tilt 계수 × 목표 tilt × 채널 위치 = 기본 곡선
95채널 입력 파워·사용 mask·파워 비중·채널 위치 + 전체 운전조건
  → 2층 Transformer Encoder(4 attention heads, hidden 64)
  → 채널별 보정 곡선
예측 gain = 기본 곡선 + Transformer 보정 곡선
```

입력의 전체 운전조건에는 목표 gain/tilt, EDFA 입력/출력 총파워, 채널 사용률과 스펙트럼 통계가 포함됩니다. 사용하지 않는 채널도 mask=0인 토큰으로 전달되며, 점수와 손실에는 사용 채널만 포함됩니다. 물리 방정식의 직접 해는 아닙니다.

## 공식 공개 데이터 사후 검증 (2026-11-02)

동일한 공식 데이터 저장소, seed 2026, 학습·검증 분할 코드, 점수식으로 기존 모델과 비교했습니다. 낮을수록 좋은 점수입니다.

| 모델 | Aging | SHB | Unseen | 전체 |
|---|---:|---:|---:|---:|
| 기존 PI-EDFA | 0.144064 | 0.147435 | 0.440072 | 0.382980 |
| 새 Transformer | 0.163582 | 0.193067 | 0.364862 | 0.333409 |
| 변화(새 모델 − 기존) | +0.019518 | +0.045632 | −0.075210 | −0.049570 |

Train 490행에서 학습, 검증 93행을 모델 선택·조기 종료에 사용했습니다. 새 모델의 해당 검증 점수는 0.185417(기존 0.225505), epoch 203에 조기 종료했습니다. 공개 TestGround의 Aging 4,400행·SHB 864행·Unseen 15,792행, 총 21,056행을 학습·스케일러 적합·모델 선택에 사용하지 않았습니다. **종합과 Unseen은 개선되었으나 Aging과 SHB는 악화되었습니다.** 종합은 세 항목 점수의 단순 평균이 아닙니다. 공개 TestGround 사후 평가이며 비공개 대회 리더보드 점수가 아닙니다.

## 재현

1. [공식 데이터 저장소](https://github.com/ofc-ml-challenge/ofc-ml-challenge-data-code)를 clone한 후 Git LFS CSV를 받습니다.
2. `pip install torch numpy pandas scikit-learn matplotlib`
3. `python transformer_edfa.py --data-root /path/to/ofc-ml-challenge-data-code --output transformer_results --epochs 250 --patience 40 --cpu --evaluate-testground`

기본 설정: batch 64, learning rate 0.002, AdamW, hidden 64, 2층 encoder, seed 2026. 학습 후 `best_physics_edfa.pt`, `submission.csv`, 지표와 진단 파일이 출력됩니다. 파일명 `best_physics_edfa.pt`는 공통 평가 코드에서 이어받은 이름으로 이 폴더에서는 **Transformer 가중치**를 담습니다. 실행 당시 모델 가중치와 대용량 예측 CSV는 이 GitHub 폴더에 포함하지 않았습니다.

- [공개 TestGround 지표](testground_metrics.json)
- [내부 검증 지표](validation_metrics.json)
- [학습 곡선 수치](training_history.csv)
