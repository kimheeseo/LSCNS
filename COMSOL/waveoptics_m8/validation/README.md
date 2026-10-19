# M8 검증 프로토콜

기본 입력은 `run_validation.py`의 `DEFAULT_CONFIG`에 정의되어 있습니다. 단위와 모델 한계는 상위 [README](../README.md)를 확인하십시오.

## 사전 정의한 주요 허용 기준

- HCF 640 요소: neff 절대 오차 ≤1e-9, kt²·D·Aeff 상대 오차 각각 ≤0.0001%.
- HCF 9개 반경/파장 조합: 최대 neff 절대 오차 ≤1e-9, D 상대 오차 ≤0.0001%.
- HCF 5점 미분 분산: 상대 오차 ≤0.001%.
- MCF 독립 기준: 원통함수 차수 6→8 및 각도 샘플 256→512의 Δneff 차이 각각 ≤0.00001%.
- MCF 최종 FEM: neff 최대 절대 오차 ≤1e-7, Δneff·Lc 상대 오차 각각 ≤0.05%.
- MCF 영역 최종 두 단계의 Δneff 차이 ≤0.01%; 간격별 Lc 최대 오차 ≤0.05%.
- 이 허용치는 현재 수치 모델의 검증 기준입니다. ITU 규격이나 실측 제품 정확도 기준이 아닙니다.
- 기준을 통과시키기 위한 계수 보정이나 제품 데이터 피팅은 하지 않았습니다.

`checks.csv`에는 전체 22개 판정값과 임계치·통과 여부를 저장합니다. CLI는 하나라도 실패하면 실패 종료합니다. 노트북은 실패를 삭제하지 않고 표시합니다.

`tests.xml`의 pytest는 신규 M8 22개와 기존 연구 회귀 45개입니다. 이는 위 22개 수치 판정과 별개이며, 서로 다른 판정 수를 더해 오차율로 계산하면 안 됩니다.

## 재현

```bash
python COMSOL/waveoptics_m8/validation/run_validation.py
python -m pytest COMSOL/waveoptics_m8/tests COMSOL/research/tests/test_research.py -q
python COMSOL/waveoptics_m8/validation/make_report.py
python COMSOL/waveoptics_m8/validation/replay_notebook.py
python COMSOL/waveoptics_m8/validation/replay_notebook.py --edited
```

입력 변경 재실행은 MCF 간격을 16→18 µm, HCF 반경을 15→20 µm로 바꾸고 결합 길이와 분산이 실제로 달라지는지 검사합니다. 기본 결과 노트북은 보존합니다.
