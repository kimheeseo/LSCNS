# Universal Silica Fiber FEM Lab — v0.1 실행 검증 보고서

**실행일:** 2026-10-30. **방법:** Python 3 + SciPy 일반화 고유치 해석기, 축대칭 P1 radial FEM; Colab ipynb의 17개 셀을 로컬 Jupyter 커널에서 순서대로 실제 실행하고 출력을 노트북에 저장. **Colab 클라우드 GPU/CPU 자체에서의 재실행은 이 검증에 포함하지 않음.**

## 1. 설계 모델

- 기본 모델: G.654.E-like **가상 Step-index**, λ=1550 nm, core radius=5.5 µm (**가정**), Δn=0.003478579 (**기존 MFD 기준 보정값**), 계산영역 반경=40 µm.
- 클레딩: 순수 실리카 Sellmeier, core: `n_silica(λ) + constant Δn`.
- 범용 단면: Step-index, Graded core, Trench-assisted, 사용자 정의 동심원 다층.
- 방정식: LP01 약유도 스칼라 radial FEM. 모드의 실제 장을 Gauss 적분해 Aeff 및 두 종류 MFD 계산.
- D: FEM neff를 파장별 직접 재해석한 뒤 5점 유한차분으로 계산; λ 미분 간격 민감도 확인.

## 2. 기본 수치 출력 (직접 실행)

| 구분 | 결과 |
|---|---:|
| neff | 1.4457189604 |
| MFD Petermann-II | 12.500070 µm (**보정 기반**) |
| MFD second moment | 12.819972 µm |
| Aeff | 121.924352 µm² |
| gamma (n2=2.2e-20 m²/W 가정) | 0.731443 /(W·km) |
| generalized eigen residual | 6.75e-15 |
| outer 20% power fraction | 5.83e-8 |
| D (10 nm derivative step) | 19.82564 ps/(nm·km) |
| D (5 nm derivative step) | 19.82564 ps/(nm·km) |

## 3. 자동 검사

**기본 설정 7/7 통과:** guided proxy, 고유방정식 residual, 경계 tail, 메시 neff 안정성, 분산 미분간격 안정성, 외부 계산 영역 변경 영향, Step-index Bessel LP01 상대 독립검증(절대 neff 오차 <1e-6).

### 추가 구조 샘플 (illustrative, λ=1550 nm)

| 모델 | neff | MFD Petermann-II (µm) | Aeff (µm²) | D (ps/(nm·km)) |
|---|---:|---:|---:|---:|
| G.654.E-like Step-index | 1.445718940 | 12.50012 | 121.92564 | 19.82558 |
| G.652.D-like Step-index | 1.446104662 | 10.18291 | 79.29586 | 16.99274 |
| Graded core example | 1.445828769 | 10.19305 | 76.88761 | 14.91916 |
| Trench-assisted example | 1.446043440 | 9.83941 | 74.82661 | 17.80367 |
| Multilayer example | 1.447001712 | 9.58252 | 71.69077 | 19.16239 |

해당 모델명은 물리 구조에 대한 **개념 예시**입니다. 제조사 제품 데이터와 독립 검증한 값이 아닙니다.

## 4. 한계 및 다음 검사

1. Step-index의 Bessel 기준을 제외하면 Graded/Trench/Multilayer는 메시 및 미분간격 안정성만 확인했습니다. 공개 독립 reference를 추가해야 합니다.
2. Ge/F 도핑 실리카의 파장별 물성, 실제 RIP(Refractive Index Profile), C-band 외 광학 손실, MFD 측정 방식 차이는 별도 처리해야 합니다.
3. Full-vector, MCF, HCF, PCF, PMF, 비원형/고차 모드, macro/microbending, splice loss 등은 v0.1에서 분석하지 않습니다.
4. 환경의 달라진 Colab SciPy 버전에서 재실행하여 수치재현성과 경계조건을 점검하고, GitHub Actions 자동 회귀 테스트 추가가 권장됩니다.
5. 동일한 값으로 MFD를 보정한 뒤 MFD를 다시 맞춘 것을 독립 제품 검증이라고 해석하면 안 됩니다.

결론: **축대칭 실리카 광섬유의 범용 구조 정의 및 LP01 FEM 최소기능 구현 완료**. 모든 실리카 광섬유에 대한 COMSOL 동등성/제품 검증은 미완료.
