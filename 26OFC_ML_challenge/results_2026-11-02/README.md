# PI-EDFA 공식 공개 데이터 사후 검증 (2026-11-02)

[PI-EDFA 원본 노트북](../OFC_2026_Physics_Guided_EDFA_Colab.ipynb)을 공식 데이터 저장소의 Train CSV로 학습하고 공개 TestGround 정답으로 평가한 실행 기록입니다. 점수는 낮을수록 좋습니다.

| 팀 | Aging | SHB | Unseen | 종합 |
|---|---:|---:|---:|---:|
| SJTU (제공된 비교표) | 0.1302 | 0.0798 | 0.0987 | 0.1080 |
| IPOC (제공된 비교표) | 0.1215 | 0.1130 | 0.1155 | 0.1173 |
| SIL (제공된 비교표) | 0.1212 | 0.0706 | 0.1242 | 0.1212 |
| NU (제공된 비교표) | 0.1164 | 0.0724 | 0.3130 | 0.2408 |
| **PI-EDFA (이번 실행)** | **0.144064** | **0.147435** | **0.440072** | **0.382980** |

PI-EDFA의 공개 TestGround 샘플 수는 Aging 4,400, SHB 864, Unseen 15,792, 총 21,056개입니다. Train에서 분리한 검증 점수는 0.225505입니다. 이 검증 세트는 조기 종료와 최고 체크포인트 선택에도 사용되었습니다.

**해석 범위:** PI-EDFA 수치는 공개 TestGround에 대한 사후 재현 결과입니다. 다른 팀의 네 행은 사용자가 제공한 비교표의 숫자이며, 이번 실행으로 재평가한 값이 아닙니다. 같은 비공개 Kaggle 리더보드 조건에서 직접 측정한 순위라고 주장할 수 없습니다. 종합 점수는 세 카테고리 점수의 산술평균이 아니라 전체 유효 채널을 합쳐 계산한 값입니다.

## 재현

1. 공식 저장소 `ofc-ml-challenge/ofc-ml-challenge-data-code`를 클론하고 Git LFS 파일을 받습니다.
2. 노트북의 `%%writefile /content/physics_edfa.py` 셀에서 생성되는 Python 스크립트를 실행합니다.
3. 실행 명령: `python physics_edfa.py --data-root /path/to/ofc-ml-challenge-data-code --output physics_results --epochs 250 --patience 40 --cpu --evaluate-testground`

설정: seed 2026, batch size 64, hidden 96, learning rate 0.002, validation fraction 0.2 (스크립트 기본값). CPU에서 epoch 124에 조기 종료되었습니다. 원본 Train 490행, 검증 대상 93행입니다. 학습에는 TestGround 정답을 사용하지 않았습니다.

- [TestGround 전체·카테고리별 지표](testground_metrics.json)
- [학습 중 검증 지표](validation_metrics.json)
- [검증 카테고리별 지표](validation_by_category.json)
- [학습 곡선 수치](training_history.csv)

공식 자료: [대회 설명](https://ofc-ml-challenge.github.io/ofc-ml-challenge/ml-challenge/) · [데이터/코드 저장소](https://github.com/ofc-ml-challenge/ofc-ml-challenge-data-code)
