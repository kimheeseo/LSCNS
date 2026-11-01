# Physics-Guided PI-EDFA | Official Data Benchmark (2026-11-02)

**목적:** 물리 기반 특성·스펙트럼 보정형 PI-EDFA 모델을 OFC 2026 공식 데이터로 학습하고 공개 TestGround에서 **사후 검증한 결과 기록**입니다.

## 모델 및 검증
- 원본 코드: [OFC_2026_Physics_Guided_EDFA_Colab.ipynb](../OFC_2026_Physics_Guided_EDFA_Colab.ipynb)
- 학습: 공식 Train **490개**, 내부 검증 **93개** (seed 2026)
- 모델 구성: 물리적 Gain/Tilt 기준선 + 학습 가능한 스펙트럼 보정
- 평가: TestGround **21,056개** (Aging 4,400 / SHB 864 / Unseen 15,792)
- 데이터 누출 방지: TestGround 정답은 훈련·스케일링·체크포인트 선택에 사용하지 않음

## 결과
*점수는 낮을수록 좋습니다.*

| 구분 | Aging | SHB | Unseen | 종합 |
|---|---:|---:|---:|---:|
| **PI-EDFA 공개 TestGround** | 0.144064 | 0.147435 | 0.440072 | **0.382980** |

- Train 내부 검증: **0.225505**
- 상대적으로 **Unseen의 오차가 큼** → 미학습 조건 일반화 개선 필요
- [Spectral Transformer 비교](../PI_EDFA_Transformer/README.md): 공개 TestGround 종합 **0.333409**
- 본 실험의 숫자는 **비공개 대회 리더보드 순위가 아닙니다.**

## 결과 파일
- [testground_metrics.json](testground_metrics.json) — 종합·Aging·SHB·Unseen 점수
- [validation_metrics.json](validation_metrics.json) · [validation_by_category.json](validation_by_category.json) — 내부 검증
- [training_history.csv](training_history.csv) — Epoch별 학습 이력

## 재현
공식 [데이터 저장소](https://github.com/ofc-ml-challenge/ofc-ml-challenge-data-code)의 실제 CSV를 준비한 뒤, 원본 Colab에서 생성한 `physics_edfa.py`를 실행합니다.

```bash
python physics_edfa.py --data-root /path/to/ofc-ml-challenge-data-code \
  --output physics_results --epochs 250 --patience 40 --evaluate-testground
```

**상태:** 공개 TestGround 검증 기록 보존; 독립 Train 검증과 사후 TestGround 점수를 구분해 해석.
