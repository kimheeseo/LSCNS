# GN 모델 검증 자료

| 파일·폴더 | 설명 |
|---|---|
| [Poggiolini_assessment.py](./Poggiolini_assessment.py) | GN 수식 관계 7개와 Sobol 적분 수렴을 점검하고 결과 JSON을 저장하는 스크립트. |
| [execute_notebook.py](./execute_notebook.py) | 같은 폴더의 `GN_Model_Colab.ipynb`를 IPython에서 실행하고 출력을 저장하는 도구. 현재 해당 노트북은 이 폴더에 없습니다. |
| [reference_notes.md](./reference_notes.md) | 기존 Carena 그림 5 **시뮬레이션 마커 기준** 검증의 참조값·계산 조건·근사 설명. |
| [requirements.txt](./requirements.txt) | Python 실행에 필요한 패키지 목록. |
| [result/](./result/) | PSCF **저자 GN 실선 기준** 비교 그래프·CSV와 Poggiolini 평가 보고서. 자세한 설명은 [결과 안내](./result/readme.md) 참조. |

수식 검증은 상위 폴더의 [gn_integral_general.py](../gn_integral_general.py)를 사용합니다. 이 폴더에서 다음 명령으로 실행할 수 있습니다.

```bash
pip install -r requirements.txt
python Poggiolini_assessment.py
```

실행 결과는 `result/paper1_reproducibility.json`에 저장됩니다. 수식 검증의 PASS는 실제 링크의 절대 예측 정확도를 보증하지 않습니다.
