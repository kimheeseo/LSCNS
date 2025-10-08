# 1) 다른 사람에게 무엇을 전달하면 되나요?
## A안: “파이썬 설치 없이” 실행 (권장)
- 단 1개 파일: dist/FiberAnalyzerRunner.exe
사용자는 자기 PC에서 exe 더블클릭 → ab.xlsx, alls.xlsx만 선택해서 실행하면 됩니다.

(선택) 짧은 README.txt 한 장, 필요하면 샘플 ab.xlsx/alls.xlsx
- 주의: exe는 빌드한 OS/아키텍처(보통 64bit Windows) 와 동일한 환경에서 쓰세요. 사내 보안 정책에 따라 SmartScreen 경고가 뜰 수 있으니 안내해 주세요.

## B안: “파이썬 설치된 환경”에서 실행
- app.py, new_main.py 두 파일
- 사용자에게 pip install pandas openpyxl numpy만 안내하고 python app.py로 실행

# 2) 실행 내용을 바꾸려면 어디를 수정하나요?
- 간단히 무엇을 바꾸고 싶으냐에 따라 달라집니다.
## (1) 파이프라인 로직/출력(분석 동작)만 변경
예: 계산식, 경로, 생성 파일, 로그 문구, 스텝 구성 등 new_main.py만 수정하면 됩니다.

배포 형태별로:
스크립트 방식(파이썬으로 실행): new_main.py만 교체하면 끝. EXE 방식: exe 안에 new_main이 포함되어 있으므로, 다시 빌드 필요.

빌드 명령(권장):
pyinstaller --clean --onefile --noconsole ^
  --name FiberAnalyzerRunner ^
  --hidden-import new_main ^
  --collect-all pandas --collect-all numpy --collect-all openpyxl ^
  app.py

## (2) 실행 방법/입력 UI/로그 표시 방식 변경
예: 파일 선택 필드 추가, 버튼/진행률, 결과 자동 열기, 로그 인코딩/표시 방식, 작업 디렉터리 변경 등
- app.py를 수정해야 합니다.
EXE로 배포한다면 당연히 다시 빌드가 필요합니다. (위 명령 동일)

## (3) new_main.py의 CLI 인자가 바뀐 경우
예: --ab/--alls/run-all → 다른 옵션/서브커맨드로 변경
- new_main.py + app.py 둘 다 수정 필요 (앱에서 만드는 커맨드 라인을 바꿔줘야 함)
EXE 배포라면 다시 빌드.

# 빠른 판단 가이드
- “분석 내용/결과만 바꾼다” → new_main.py만 수정
   스크립트 배포면 끝, EXE 배포면 재빌드 필요
- “UI/실행/로그/인자 형식 바꾼다” → app.py 수정 (EXE는 재빌드)
- “인자 이름/순서 바뀐 new_main에 맞춘다” → new_main.py + app.py 수정 (EXE는 재빌드)

필요하시면, 변경 목적을 알려주세요. 해당 변경에 맞춰 딱 필요한 부분(예: app.py의 커맨드 생성부, run_cwd, 로그 처리, 새 입력 위젯)만 정확히 수정한 패치 버전으로 드릴게요.
