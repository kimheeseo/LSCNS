# EGN_model — Full-EGN 계산 및 검증

이 폴더는 GN 모델의 변조 의존 보정을 포함한 **Full EGN(SCI+XCI+MCI)** 계산과 문헌 재현 결과를 보관합니다.

## 사용할 파일

| 목적 | 파일·폴더 | 설명 |
|---|---|---|
| 새 Full-EGN 계산 | [EGN_adaptive.py](./EGN_adaptive.py) | 현재 권장 계산기. 적응형 적분과 수렴 점검을 사용 |
| 기존 구현·문헌 재현 | [EGN_MODEL_Validation/](./EGN_MODEL_Validation/) | Carena (2014) Fig. 1·3·6·8 비교 자료와 이전 구현 |
| GN 기본식·GN/EGN 비교 | [EGNvsGN/](./EGNvsGN/) | 상위 GN 엔진의 수학 검증·사용 가이드·논문 재현 노트북 |

## Full-EGN 결과

`EGN_adaptive.py`는 다음을 분리해 제공합니다.

```text
Total EGN NLI = SCI + XCI + MCI
```

- GN baseline과 EGN 보정항
- SCI/XCI/MCI별 NLI 전력
- GN 대비 EGN 비율 및 dB 차이
- 주파수 적분·수신대역 적분의 수렴 진단

`verify_convergence=True`와 `strict_convergence=True`를 사용하면 설정한 오차 기준을 충족하지 못한 계산을 오류로 처리합니다.

## 적용 범위

정밀 Full-EGN 경로는 다음 조건에서 검증되었습니다.

- dual-polarization, coherent accumulation
- 직사각형(roll-off 0) 스펙트럼, 동일 심볼률·등간격 WDM
- 동일한 lumped-EDFA span과 `beta2` 지배 분산
- MCI 계산: 홀수·대칭 채널 배열, 동일 채널 전력·변조, 중앙 CUT

이 범위를 벗어난 heterogeneous span, distributed Raman, `beta3`, ROADM filtering, PMD/PDL, HCF/MCF 효과는 별도 검증 또는 모델 확장이 필요합니다.

## 검증 해석

`converged=True`는 **수치 적분이 설정 tolerance 내에서 안정화되었다는 뜻**입니다. 이는 SSFM·실험·상용 시뮬레이터와의 절대 오차를 자동으로 보장하지 않습니다. 새 조건에서는 일부 기준점을 독립 결과와 교차 검증한 뒤 sweep에 사용해야 합니다.

## 참고문헌

A. Carena et al., *EGN model of non-linear fiber propagation*, Optics Express 22, 16335–16362 (2014), DOI: 10.1364/OE.22.016335.
