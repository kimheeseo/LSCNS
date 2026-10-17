# GNmodel — GN 엔진 검증 자료

상위 폴더의 `gn_integral_general.py`가 구현한 GN 적분의 **수식 일관성·Sobol QMC 수렴성·문헌 비교** 자료입니다. 실제 전송 링크의 절대 예측 정확도를 단독으로 보증하는 폴더는 아닙니다.

## 구성

| 파일·폴더 | 역할 |
|---|---|
| [Poggiolini_assessment.py](./Poggiolini_assessment.py) | GN 수식 관계 7개와 Sobol 적분 수렴을 점검하고 JSON 결과 저장 |
| [reference_notes.md](./reference_notes.md) | Carena 그림 5의 시뮬레이션 마커 기준값·계산 가정·근사 설명 |
| [result/](./result/) | PSCF 저자 GN 실선과의 비교 그래프·CSV·재현성 결과 |
| [requirements.txt](./requirements.txt) | 검증 실행용 Python 패키지 목록 |

## 실행

```bash
pip install -r requirements.txt
python Poggiolini_assessment.py
```

실행 결과는 `result/paper1_reproducibility.json`에 저장됩니다.

## 검증 범위와 한계

검증은 단위 일관성, 대칭성, 제한 조건, NLI의 (P^3) 의존성, QMC 수렴 및 공개 문헌의 GN 기준 곡선과의 비교를 대상으로 합니다.

PASS는 코드가 해당 수학·수치 검사를 통과했다는 의미입니다. 실제 장거리 링크에 적용할 때는 광섬유 파라미터, 증폭기 조건, 누적 방식, 송수신기 penalty를 별도로 설정하고 가능한 경우 SSFM·실험·상용 시뮬레이터와 교차 검증해야 합니다.
