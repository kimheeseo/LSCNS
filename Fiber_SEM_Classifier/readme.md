# Fiber SEM Classifier

광섬유 SEM 단면 이미지를 **Bending · Cleaving · Fracture**로 분류하는 CNN 분석 도구입니다. Google Colab에서 전처리부터 증강·학습·검증·신규 이미지 분석까지 수행하고, 데이터와 결과를 Google Drive에 저장합니다.

## 파일 구성

| 파일 | 용도 | 실행 |
| --- | --- | --- |
| [convert_tif_to_png_ipynb.ipynb](./convert_tif_to_png_ipynb.ipynb) | TIF/TIFF → PNG 변환, 파일명 정리, 원본 대응표 저장 | [Colab](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/Fiber_SEM_Classifier/convert_tif_to_png_ipynb.ipynb) |
| [SEM_CNN_classifier_epoch3.ipynb](./SEM_CNN_classifier_epoch3.ipynb) | 데이터 검사·증강, CNN 학습·검증, 다중 이미지 분류 | [Colab](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/Fiber_SEM_Classifier/SEM_CNN_classifier_epoch3.ipynb) |
| [SEM_CNN_classifier_regularized_v2.ipynb](./SEM_CNN_classifier_regularized_v2.ipynb) | 과적합 완화: 단계별 전이학습, Dropout·Label Smoothing·Train 전용 명암 증강 | [PR 브랜치 Colab](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/feature/sem-cnn-regularization-v2-20261108/Fiber_SEM_Classifier/SEM_CNN_classifier_regularized_v2.ipynb) |


## 과적합 개선 실험 노트북 v2

- **데이터셋 유지:** 기존 원본의 Train/Validation 그룹 분리, Train 10배 증강, Google Drive 이미지 저장·재검증·재로딩 규칙을 그대로 사용합니다. 이미 저장된 증강본을 재사용합니다.
- **학습 변경:** Head-only 3 Epoch 학습 후 ResNet18 `layer4` 또는 EfficientNet-B0 마지막 feature block만 미세조정합니다. 분류층 Dropout 0.35, Weight Decay 0.001, Label Smoothing 0.05 및 Train 전용 약한 밝기·대비 증강을 사용합니다.
- **학습률:** Head `3e-4`, 미세조정 block `3e-5`(초기 단계에서는 block 동결). 조기 종료 Patience 4. 기본 Epoch 10이며 GPU 런타임 사용을 권장합니다.
- **기존 결과 보존:** 모델은 `03_Models/Regularized_v2/`, 보고서는 `04_Training_Reports/Regularized_v2/`로 구분하여 과거 결과 파일을 덮어쓰지 않습니다. 다운로드/모델 Release 파일명 규칙은 유지합니다.
- **검증 한계:** 이 노트북은 코드를 개선한 실험판이며 **실제 Colab 학습 결과가 검증된 것은 아닙니다.** 물리적으로 동일한 시료가 자동 그룹화에 누락됐는지 점검해야 하고 Validation은 독립 Test가 아닙니다. 기존 모델 가중치와 v2의 분류층 구조는 호환되지 않습니다.

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

## 학습 완료 후 GitHub 자동 업로드

1. GitHub Fine-grained personal access token을 만들고 `LSCNS` 저장소의 **Contents: Read and write** 권한을 부여합니다.
2. Colab 왼쪽 **보안 비밀(열쇠 아이콘)**에 `GITHUB_TOKEN`으로 저장하고 노트북 접근을 허용합니다.
3. 최신 노트북을 **모두 실행**하면 마지막 **5. 학습 모델·성능 보고서 GitHub 자동 업로드** 셀이 학습·평가 이후 실행됩니다.

[sem_github_upload.py](./sem_github_upload.py)가 최적 모델 두 개·선택 정보·허용된 성능 CSV/차트를 ZIP으로 묶어 Drive의 `07_GitHub_Exports/`에 보관하고 [GitHub Releases](https://github.com/kimheeseo/LSCNS/releases)에 첨부합니다. 원본·증강·신규 SEM 사진과 optimizer 체크포인트는 제외합니다. 모델과 보고서는 저장소의 접근 범위에 따라 공개됩니다.

- 기존 개인 Drive 노트북은 자동 갱신되지 않습니다. 마지막 안내·업로드 셀을 복사하고 `requests`가 설치되어 있는지 확인하세요. 현재 런타임에서 학습을 마쳤다면 학습 셀을 다시 실행할 필요 없이 업로드 셀만 실행할 수 있습니다.
- UI 버튼으로 나중에 재학습한 경우 마지막 업로드 셀도 다시 실행해야 합니다.
- 매번 새 태그로 게시합니다. 네트워크 오류 시 노트북에 안내된 `upload_bundle(sem_export_zip, ...)`로 같은 ZIP을 재시도할 수 있습니다.
- 업로드 생략: `AUTO_UPLOAD_GITHUB = False`. 2 GiB 이상 ZIP은 모델별 분리가 필요합니다.
- ZIP은 가중치·보고서 패키지입니다. 다른 PC에서 추론하려면 동일한 Pipeline 코드·전처리가 필요합니다. 추출 후 `model_choice="resnet18"` 또는 `"efficientnet_b0"`로 모델을 지정하거나 선택 정보의 경로를 복원하세요.

업로드 코드의 ZIP 구성·허용 파일·재시도·HTTP 오류 처리는 모의 응답으로 점검했습니다. 실제 Colab 학습 및 GitHub 모델 업로드는 사용자의 런타임에서 실행됩니다.
