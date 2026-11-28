# Cable · 광섬유·해저케이블 자료

광섬유 제조·품질 분석, 해저케이블 계산 및 교육용 3D 도구를 모아 둔 폴더입니다.

## 폴더 구성

| 폴더 | 주요 내용 | 대표 파일 |
|---|---|---|
| [Fiber_SEM_Classifier](./Fiber_SEM_Classifier/) | 광섬유 SEM 이미지 전처리, CNN 분류 및 Grad-CAM 시각화 노트북과 학습·검증 보고서 | `convert_tif_to_png_ipynb.ipynb`, `SEM_CNN_gradcam.ipynb` |
| [drawing](./drawing/) | 광섬유 인출 공정, Draw Tower 체험, 인장·OTDR·광파워 측정을 설명하는 교육용 3D 웹 도구 | `index.html` |
| [simple_preform_data_analyzer](./simple_preform_data_analyzer/) | 프리폼 품질 Excel 데이터를 그룹별로 분류하고 보고서·월별 추세 그래프를 생성하는 Python 도구 | `app.py` |
| [submarine_tour](./submarine_tour/) | 한국–일본 데이터센터 사이의 광신호 경로와 육양국·해저케이블·포설 장비를 살펴보는 교육용 3D 투어 | `index.html` |
| [submarinecable](./submarinecable/) | 해저케이블의 충격·동관 내경 추정, OSNR 링크 예산, 광증폭 및 signal droop 분석 노트북 | `signal_droop.ipynb` 및 하위 폴더 |

## 주요 하위 폴더

| 경로 | 내용 |
|---|---|
| [drawing/tests](./drawing/tests/) | 광섬유 인출·측정 웹 도구의 기능 검사 코드 |
| [simple_preform_data_analyzer/practice](./simple_preform_data_analyzer/practice/) | 프리폼 분석 도구의 개발·실험용 Python 코드 |
| [submarine_tour/tests](./submarine_tour/tests/) | 해저케이블 투어의 기능 검사 코드 |
| [submarine_tour/tools](./submarine_tour/tools/) | 투어에 사용하는 해양 모델 제작 스크립트 |
| [submarinecable/OSNR_Link_Budget](./submarinecable/OSNR_Link_Budget/) | OSNR 링크 예산 계산 방법을 정리한 노트북 |
| [submarinecable/OptiCommpy](./submarinecable/OptiCommpy/) | EDFA·WDM 및 OOK 광전송 시뮬레이션 노트북 |
| [submarinecable/Tube_Diameter_Estimator](./submarinecable/Tube_Diameter_Estimator/) | 수중 낙하 물체의 속도·충격에너지와 동관 내경을 추정하는 노트북 |

## 웹 도구 실행

[광섬유 Drawing Studio](https://kimheeseo.github.io/LSCNS/Cable/drawing/) · [해저케이블 3D 투어](https://kimheeseo.github.io/LSCNS/Cable/submarine_tour/)

자세한 사용법과 변경 이력은 각 폴더의 README를 참고하세요.
