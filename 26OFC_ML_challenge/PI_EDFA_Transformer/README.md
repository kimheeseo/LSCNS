# Spectral Attention Transformer | 95-channel EDFA Gain

**목적:** 채널별 입력 파워와 WSS 활성 패턴의 관계를 Transformer Attention으로 학습해 EDFA Gain Profile을 예측합니다.

## 핵심 구조
- **기본 곡선:** 목표 Gain + 채널 위치에 따른 Tilt
- **입력:** 95채널 입력 파워·Mask·파워 비중·채널 위치 및 전체 운전 조건
- **모델:** Transformer Encoder **2층 / 4 Attention heads / hidden 64**
- **출력:** `예측 Gain = 기본 곡선 + 채널별 Attention 보정량`
- **평가:** WSS 활성 채널만 손실 및 공식 점수에 포함

> 물리 방정식을 직접 푸는 모델이 아니라, 운전 조건과 채널 관계를 학습하는 잔차형 Transformer입니다.

## 공식 공개 TestGround 사후 검증 (2026-11-02)
*동일 공개 데이터 및 재현 절차의 PI-EDFA 기준 모델과 비교. 점수는 낮을수록 좋습니다.*

| 모델 | Aging | SHB | Unseen | 종합 |
|---|---:|---:|---:|---:|
| PI-EDFA | **0.144064** | **0.147435** | 0.440072 | 0.382980 |
| **Spectral Transformer** | 0.163582 | 0.193067 | **0.364862** | **0.333409** |

**해석:** Transformer는 Unseen 및 종합 점수에서 개선되었으나 Aging·SHB에서는 악화되었습니다. 비공개 대회 순위가 아닌 공개 TestGround 사후 평가이며, Train 내부 검증 점수는 **0.185417**입니다.

## 주요 파일
- [transformer_edfa.py](transformer_edfa.py) — 모델, 학습, 평가, 제출 파일 생성
- [testground_metrics.json](testground_metrics.json) — 전체·카테고리별 사후 결과
- [validation_metrics.json](validation_metrics.json) — 내부 검증 결과
- [training_history.csv](training_history.csv) — 학습 이력

## 실행
공식 [데이터 저장소](https://github.com/ofc-ml-challenge/ofc-ml-challenge-data-code)의 CSV가 필요합니다.

```bash
pip install torch numpy pandas scikit-learn matplotlib
python transformer_edfa.py --data-root /path/to/ofc-ml-challenge-data-code \
  --output transformer_results --epochs 250 --patience 40 --evaluate-testground
```

**상태:** 공개 데이터 사후 검증 완료, SHB·Aging 편차 개선이 후속 과제.
