main.py(gn model) 검증 논문 List
# 1. 폴더 - Carena
- Carena, Andrea, et al. "Modeling of the impact of nonlinear propagation effects in uncompensated optical coherent transmission links." Journal of Lightwave technology 30.10 (2012): 1524-1539.: Fig 5.
- https://ieeexplore.ieee.org/document/6158564
- GN 모델의 해석식과 수치 시뮬레이션 결과를 비교한 검증 논문 (실험 측정 논문 아님)
   - PSCF·SMF·NZDSF 조건에서 최대 전송거리 및 최적 launch power 검증
   - main.py 최대 전송거리 오차율: 전체: 7.41%
     · PSCF: 5.48%, SMF: 5.53%, NZDSF: 11.21%
## 오차 원인:
- 본 코드는 closed-form GN 근사식이고, 논문 기준값은 더 상세한 수치 시뮬레이션 결과이므로 특히 NZDSF처럼 분산이 낮고 비선형성이 큰 조건에서 차이가 커질 수 있음.
- 최대 전송거리를 100 km span 단위로 선택하므로, 짧은 거리에서는 한 span 차이만으로도 오차율이 크게 계산 됨.
- 비교값은 논문 raw 데이터가 아닌 Figure 5의 digitized·반올림 값이어서 추가적인 판독 오차가 포함됨.

Lightera Ocean Fiber
- https://www2.ofsoptics.com/ocean-fiber?srsltid=AfmBOopQ4B9HNl0dGYdbzixn-5WwkA7XuM9-45dupzbbbG2hMil808qZ
