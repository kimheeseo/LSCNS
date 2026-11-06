# Fiber SEM Classifier

광섬유 SEM 단면 이미지를 **Bending · Cleaving · Fracture**로 분류하는 CNN 분석 도구입니다. Google Colab에서 전처리부터 증강·학습·검증·신규 이미지 분석까지 수행하고, 데이터와 결과를 Google Drive에 저장합니다.

## 파일 구성

| 파일 | 용도 | 실행 |
| --- | --- | --- |
| [convert_tif_to_png_ipynb.ipynb](./convert_tif_to_png_ipynb.ipynb) | TIF/TIFF → PNG 변환, 파일명 정리, 원본 대응표 저장 | [Colab](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/Fiber_SEM_Classifier/convert_tif_to_png_ipynb.ipynb) |
| [SEM_CNN_classifier_epoch3.ipynb](./SEM_CNN_classifier_epoch3.ipynb) | 데이터 검사·증강, CNN 학습·검증, 다중 이미지 분류 | [Colab](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/Fiber_SEM_Classifier/SEM_CNN_classifier_epoch3.ipynb) |

## 주요 기능

- **자동 라벨링**: `bending/`, `cleaving/`, `fracture/` 폴더명을 정답으로 사용합니다. HCF는 하위 폴더 또는 파일명 태그로 구분하며, 3개 클래스에 함께 포함합니다.
- **분할·증강**: 동일·유사 이미지와 시료 후보를 그룹화한 후 Train/Validation으로 분할합니다. Train은 원본 버전 포함 10개, Validation은 원본 버전 1개를 저장합니다.
- **학습·비교**: 사전학습 **ResNet18 / EfficientNet-B0** 각각에 대해 원본 전용·증강 학습을 비교하고, Validation Macro F1 기준으로 모델을 선택합니다.
- **성능 보고**: 정확도·Macro F1, 혼동행렬, 손실·정확도 곡선, 오분류 목록, 증강 효과 및 HCF/non-HCF 비교를 제공합니다.
- **신규 이미지 분석**: 여러 이미지의 분류 결과, Grad-CAM·활성 영역 박스, 통합 TXT 보고서, CSV·JSON·이미지를 저장하고 ZIP으로 다운로드합니다.

## 사용 순서

1. 원본을 `bending/`, `cleaving/`, `fracture/` 폴더로 정리합니다. TIF/TIFF는 변환 노트북에서 `src_dir`·`dst_dir`를 설정하여 PNG로 준비합니다.
2. 분류 노트북을 Colab에서 열고 **GPU 런타임**을 선택합니다.
3. Google 인증·Drive 마운트를 진행하고 `SOURCE_FOLDER_ID`를 입력 폴더에 맞게 설정합니다.
4. 셀을 순서대로 실행합니다: **원본 검사 → Train 10배 증강 → Drive 저장·재로딩 검증 → CNN 학습·평가**.
5. 마지막 UI에서 신규 이미지를 업로드하고 분석 후 **전체 결과 ZIP 다운로드**를 누릅니다.

| 설정 | 현재 값 |
| --- | --- |
| 입력 이미지 | 클래스별 폴더의 PNG |
| 모델 입력 크기 | `384 × 384` |
| 학습 셀 기본 Epoch | `3` — `EPOCHS`에서 변경 |
| 허용 Epoch / UI 기본값 | `1~50` / `50` |
| 난수 Seed | `42` |

## 저장 결과

입력 폴더 상위의 `SEM_CNN_Results/Fiber_Category_PNG_FolderLabels/`에 저장합니다. 학습은 Drive에 저장된 이미지의 무결성을 확인한 뒤 재로딩하여 수행합니다.

| 폴더 | 내용 |
| --- | --- |
| `01_Dataset_Manifest` | 원본 목록, 시료 그룹, 분할·전처리 기록 |
| `02_Augmented_Dataset` | Train/Validation 이미지 |
| `03_Models` | 모델 가중치, 체크포인트, 선택 모델 정보 |
| `04_Training_Reports` | 성능 CSV, 혼동행렬·학습 곡선, 오분류 목록 |
| `05_Predictions` | 신규 분석 이미지·보고서·ZIP |

> **해석 시 유의점:** Validation은 모델 선택에 사용되므로 독립 Test 성능이 아닙니다. 자동 그룹화·정보바 제외 결과를 확인해야 하며, Grad-CAM과 빨간 박스는 모델 활성 영역을 나타냅니다. 실제 파괴 기점이나 물리적 원인을 확정하지 않습니다.
