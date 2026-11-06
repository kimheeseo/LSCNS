# Fiber SEM Classifier

광섬유 SEM 단면 이미지를 **Bending / Cleaving / Fracture**로 분류하는 CNN 기반 분석 도구입니다. Google Colab에서 데이터 검사, 증강, 학습·검증, 신규 이미지 분석을 수행하고 결과를 Google Drive에 저장합니다.

## 주요 기능

- **폴더 기반 자동 라벨링**: `bending`, `cleaving`, `fracture` 폴더명을 정답으로 사용하며, HCF는 별도 클래스가 아닌 구조 태그로 관리합니다.
- **데이터 분할·증강**: 동일·유사 이미지와 시료 후보를 그룹화한 뒤 Train/Validation으로 분할합니다. Train은 원본 버전 포함 10배, Validation은 원본 버전만 저장합니다.
- **CNN 학습·비교**: 사전학습 ResNet18과 EfficientNet-B0를 사용해 원본 전용 학습과 증강 학습을 비교하고, Validation Macro F1 기준으로 모델을 선택합니다.
- **성능 확인**: 정확도·Macro F1, 혼동행렬, 학습 곡선, 오분류 목록, HCF/non-HCF 비교 결과를 제공합니다.
- **신규 이미지 분석**: 여러 이미지를 분류하고 Grad-CAM·활성 영역 박스, 통합 TXT 보고서, CSV·JSON·이미지를 ZIP으로 저장합니다.

## 파일 구성

| 파일 | 용도 |
| --- | --- |
| [SEM_CNN_classifier_epoch3.ipynb](./SEM_CNN_classifier_epoch3.ipynb) | SEM 이미지 검사 → 증강 → CNN 학습·검증 → 신규 이미지 분석 |
| [convert_tif_to_png_ipynb.ipynb](./convert_tif_to_png_ipynb.ipynb) | TIF/TIFF를 PNG로 변환하고 클래스·HCF 태그에 따라 파일명 정리, 대응표 저장 |
| [COSMOS_Full_Transformer_Evaluation.ipynb](./COSMOS_Full_Transformer_Evaluation.ipynb) | 별도 용도: OFC COSMOS 데이터 기반 EDFA Transformer 학습·평가 |

## 실행 방법

1. [SEM 분류 노트북을 Colab에서 열기](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/Fiber_SEM_Classifier/SEM_CNN_classifier_epoch3.ipynb).
2. 런타임을 **GPU**로 설정하고 Google 인증·Drive 마운트를 진행합니다.
3. `SOURCE_FOLDER_ID`를 입력 데이터 폴더에 맞게 설정합니다. 입력은 클래스별 폴더에 정리된 PNG이며, 필요하면 변환 노트북을 먼저 실행합니다.
4. 셀을 위에서부터 실행합니다: **원본 검사 → 10배 증강·Drive 저장 → 저장본 재검증·재로딩 → 학습·검증**.
5. 마지막 UI에서 신규 이미지를 업로드하고 분석 결과 ZIP을 다운로드합니다.

> 학습 셀의 기본 설정은 `EPOCHS = 3`, `IMAGE_SIZE = 384`입니다. Epoch은 1~50으로 변경할 수 있으며, UI의 학습 슬라이더 기본값은 50입니다.

## 저장 결과

결과는 입력 폴더 상위의 `SEM_CNN_Results/Fiber_Category_PNG_FolderLabels/`에 저장됩니다.

| 폴더 | 내용 |
| --- | --- |
| `01_Dataset_Manifest` | 원본 목록, 시료 그룹, 분할·전처리 기록 |
| `02_Augmented_Dataset` | Drive에 저장된 Train/Validation 이미지 |
| `03_Models` | 모델 가중치, 체크포인트, 선택 모델 정보 |
| `04_Training_Reports` | 성능 CSV, 혼동행렬·학습 곡선, 오분류 목록 |
| `05_Predictions` | 신규 이미지 분석 결과, 통합 보고서, ZIP |

## 결과 해석

- Validation은 모델 선택에 사용되므로 **독립 Test 성능이 아닙니다**.
- 자동 시료 그룹화와 SEM 정보바 제외에는 오류가 있을 수 있으므로 입력·전처리 기록을 확인하세요.
- Grad-CAM과 빨간 박스는 **모델 활성 영역**입니다. 실제 파괴 기점이나 물리적 파단 원인을 확정하는 결과는 아닙니다.
