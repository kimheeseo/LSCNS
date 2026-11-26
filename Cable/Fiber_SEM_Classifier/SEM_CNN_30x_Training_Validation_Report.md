# 광섬유 SEM 이미지 CNN 학습·성능 검증 기술 보고서

> **대상:** ResNet18 기반 Bending / Cleaving / Fracture 분류 및 HCF 식별  
> **분석 기준:** Google Colab `SEM_CNN_30x_7_2_1.ipynb`의 저장된 실행 출력 (2026-11-08 실행 폴더)  
> **보고일:** 2026-11-09  
> **검증 수준:** 노트북 저장 출력·코드 확인. Drive 현존 파일 및 외부 Python import 실구동은 별도 검증 필요.

## 1. Executive Summary

본 프로젝트는 광섬유 단면의 SEM(Scanning Electron Microscopy) 이미지를 이용해 **Bending, Cleaving, Fracture** 3개 파단 형태를 분류하고, 추가 출력으로 **HCF(Hollow-Core Fiber) / 일반 광섬유**를 구분하는 다중 출력 CNN 모델을 개발하였다. 678장의 원본을 **원본별 분할 후 30배 증강**하여 총 20,340장을 구성했으며, 사전학습 ResNet18 기반 모델을 20 Epoch까지 학습하였다.

**핵심 검증 결과:** 원본만으로 구성한 Test 68장의 파단 유형 분류 정확도 **85%(58/68로 환산되는 반올림 표기)**, HCF 구분 정확도 **100%(68/68)**. 증강 이미지를 포함한 Test 2,040장에서는 파단 유형 정확도 **84%**, HCF 구분 정확도 **99%**로 기록되었다. **최적 체크포인트는 Epoch 17**에서 선택되었다. 작은 독립 시료 수와 특정 HCF/파단 유형 조합의 부재 때문에 범용 성능으로 확대 해석할 수 없다.

## 2. 데이터 및 전처리

| 항목 | 기록 |
| --- | ---: |
| 원본 이미지 | 678장 |
| Bending | 146장 |
| Cleaving | 406장 (HCF 105장 포함) |
| Fracture | 126장 (HCF 8장 포함) |
| 전체 HCF | 113장 |
| 증강 비율 | 원본 포함 30장/원본 1장 |
| 증강 후 전체 | 20,340장 |
| Train / Validation / Test | 14,220 / 4,080 / 2,040장 |
| 대응 원본 수 | 474 / 136 / 68장 |
| 난수 시드 | 42 |
| 입력 크기 | 224 × 224 |
| 배치 크기 | 32 |
| 학습 Epoch | 20 |

원본을 먼저 Train/Validation/Test로 **70:20:10 비율 분할**하고 각 분할 내부에서 별도로 증강하였다. 이에 따라 **동일 원본의 증강 이미지가 서로 다른 분할에 섞이는 누수는 방지**한다. 단, 서로 다른 파일이 동일한 실제 시편이나 촬영 세션에서 나왔는지까지 분리되었다고 단정할 수는 없다.

원본 Test 세트 구성은 Bending 15장, Cleaving 41장, Fracture 12장이다. **Test의 HCF 7장은 모두 Cleaving**이며, Bending-HCF 및 Fracture-HCF 조합은 Test에 없다.

## 3. 모델 구조 및 학습

- **Backbone:** ImageNet 사전학습 ResNet18. 공유 특징 추출부는 고정하고 최종 분류 헤드 및 HCF 판별 헤드를 학습.
- **파단 분류 헤드:** Bending / Cleaving / Fracture 3클래스.
- **HCF 판별 헤드:** standard / HCF 이진 분류.
- **선택 기준:** Validation loss 최소 체크포인트를 저장. 기록상 **Epoch 17** 선택.
- **데이터 처리:** 증강 이미지 Google Drive 저장 → 재로딩 → 학습 및 평가.
- **기본 실행 장치:** 해당 저장 출력에서는 CPU로 표시됨. 장치 선택과 추론 소요 시간은 환경에 따라 달라짐.

### 3.1 학습 기록 일부

| Epoch | Train 파단 정확도 | Validation 파단 정확도 | Validation HCF 정확도 | Validation loss |
| --- | ---: | ---: | ---: | ---: |
| 17 (선택) | 81.46% | 86.37% | 98.92% | 0.3963 |
| 20 (마지막) | 80.73% | 83.73% | 98.95% | 0.4448 |

Validation은 모델 선택에 사용되므로 최종 일반화 성능은 아래의 Test 결과로 판단한다.

## 4. 시험 성능

### 4.1 원본 Test만 사용 (총 68장 — 우선 해석)

| 분류 | Precision | Recall | F1 | 원본 수 |
| --- | ---: | ---: | ---: | ---: |
| Bending | 0.93 | 0.87 | 0.90 | 15 |
| Cleaving | 0.90 | 0.85 | 0.88 | 41 |
| Fracture | 0.67 | 0.83 | 0.74 | 12 |
| **전체 파단 정확도** | | **0.85** | | **68** |
| **파단 Macro F1** | | | **0.84** | **68** |
| **HCF 구분 정확도** | | **1.00** | | **68 (HCF 7)** |

원본 Test에서 Fracture의 Precision이 0.67로 상대적으로 낮아 다른 클래스의 파단을 Fracture로 과다 판정할 가능성을 점검해야 한다. 또한 HCF 정확도 100%는 **HCF 샘플 7장에 한정된 관찰**이다.

### 4.2 증강 포함 Test (총 2,040장 — 보조 지표)

| 분류 | Precision | Recall | F1 | 이미지 수 |
| --- | ---: | ---: | ---: | ---: |
| Bending | 0.84 | 0.85 | 0.84 | 450 |
| Cleaving | 0.90 | 0.84 | 0.87 | 1,230 |
| Fracture | 0.68 | 0.83 | 0.75 | 360 |
| **전체 파단 정확도** | | **0.84** | | **2,040** |
| **파단 Macro F1** | | | **0.82** | **2,040** |
| **HCF 구분 정확도** | | **0.99** | | **2,040** |

**파단 유형 혼동행렬 (행: 실제 / 열: 예측)**

| 실제 ＼ 예측 | Bending | Cleaving | Fracture |
| --- | ---: | ---: | ---: |
| Bending | 383 | 60 | 7 |
| Cleaving | 66 | 1,028 | 136 |
| Fracture | 8 | 52 | 300 |

증강본 30장은 원본 이미지 하나에서 유래한 **상관된 관측치**이므로, 2,040장을 서로 독립적인 실제 시편 2,040개로 해석해서는 안 된다.

## 5. 모델 산출물 및 Python 재사용

Colab 실행 출력에는 다음 파일들이 생성된 것으로 기록되었다.

```text
MyDrive/SEM_CNN_3epoch/20261108_081105/
├── best_model.pt
├── augmented_dataset/
└── python_package/
    ├── fiber_sem_model.py
    ├── best_model.pt
    └── README.txt
```

**중요:** 위 경로는 Colab 실행 로그에 기반하며, 현재 Drive에서의 파일 존재·다운로드 가능 여부를 별도로 확인한 것은 아니다. GitHub에는 보고서만 게시하며, 이 단계에서 학습된 `best_model.pt` 가중치 자체를 게시하지 않는다.

### 5.1 다른 Python 프로젝트에서 import

`python_package` 디렉터리의 세 파일을 같은 폴더에 내려받은 뒤:

```bash
pip install torch torchvision Pillow
```

```python
from fiber_sem_model import FiberSEMClassifier

classifier = FiberSEMClassifier("best_model.pt", device="cpu")
result = classifier.predict("sample.png")

print(result["fracture_type"])    # bending / cleaving / fracture
print(result["fracture_scores"])  # 각 파단 클래스 점수
print(result["fiber_type"])       # HCF / standard
print(result["hcf_scores"])       # HCF 식별 점수
```

`fiber_sem_model.py`는 Python 모듈이며 `best_model.pt`는 학습된 가중치이다. **이미지를 다시 학습할 필요는 없다.** 단, 이 예시는 노트북에 생성된 모듈 코드를 검토한 사용법이며 별도 컴퓨터에서의 실제 import 테스트는 수행하지 않았다. 또한 모델 점수는 **확률 보정(calibration)이 완료된 확률값으로 간주해서는 안 된다.**

### 5.2 ONNX / 웹 배포 상태

노트북에 HTML 연동을 위한 `sem_model.onnx` 내보내기 셀은 존재하지만 **해당 셀에 저장된 실행 결과는 없다**. 따라서 ONNX 모델 생성 완료 및 HTML 기반 추론 동작은 아직 확인되지 않았다.

## 6. 한계 및 개선 계획

1. **독립 표본 확대:** Test 원본은 68장이다. 광섬유 제품·시편·SEM 촬영 장비 및 획득 일자 기준으로 외부 독립 검증을 추가한다.
2. **HCF 클래스 균형:** Test HCF는 Cleaving 7장뿐이므로 HCF-Bending/HCF-Fracture 조합의 성능은 미검증이다. 특히 Train 데이터에도 HCF-Bending은 0장이다.
3. **오분류 정밀 분석:** Fracture Precision 0.67과 Cleaving ↔ Fracture 혼동을 집중 점검한다. SEM 전문가 판독 결과와 비교한다.
4. **물리적 해석 분리:** CNN의 예측 클래스·Grad-CAM 시각화는 SEM 형상에 대한 모델 판단이지 **파단 원인, 균열 기점 또는 파괴 메커니즘의 확정 증거가 아니다**. Mirror–Mist–Hackle 영역 및 표면 결함은 별도의 물리적 근거로 검증한다.
5. **배포 검증:** 실제 Drive `best_model.pt` 파일 존재 확인 → 새로운 Python 환경에서 import 및 단일 이미지 추론 → ONNX 내보내기 및 Python/웹 추론 결과 비교.
6. **재현성 강화:** 패키지 버전, 원본 SHA256/시편 ID, 데이터 분할 manifest, 체크포인트 SHA256, confusion matrix 및 전체 Test 예측 CSV를 함께 보존한다.

## 7. 종합 결론

저장된 Colab 출력 기준 **20 Epoch 학습·평가가 완료되었고, 최적 모델을 Python import용 패키지로 저장한 기록이 확인**되었다. 원본 Test 기준 **파단 분류 85%, HCF 식별 100%**를 얻었으나 원본 Test 68장 및 HCF Test 7장의 제약 때문에 추가 독립 검증이 필요하다. 현재는 **Python 재사용을 위한 모델 산출 단계까지 완료 기록이 있는 상태**이고, 실환경 import·ONNX·웹 배포는 별도 검증 단계이다.

---

### 원자료

- [Colab 실행 노트북](https://colab.research.google.com/drive/1XsMmF023B442l8kUkZfyp3ZoWe16XaXN)
- [Fiber_SEM_Classifier 저장소 폴더](https://github.com/kimheeseo/LSCNS/tree/main/Fiber_SEM_Classifier)

*이 보고서는 노트북의 저장된 코드 및 실행 출력에 근거한 기술 기록이며, 임의의 추가 학습이나 실험을 수행했다고 주장하지 않는다.*
