# 실행 기록 해석

- 최초 테스트 20개와 oracle을 solver 전에 작성·실행했다. 초기 수집 단계의 import 경로 문제를 해결하고, oracle 2개 통과와 solver module 부재 18개 실패를 기록했다.
- 초기 solver의 TM slab 복소 S21 오차가 0.0042162300859982915로 고정 기준 0.004를 넘었다. 추가 세분화 진단 결과는 slab_failure_diagnosis.json에 있다. 허용오차 및 최초 테스트 파일은 변경하지 않았다.
- 대규모 보간에서 13.1 GiB 배열 요구가 발생한 기록은 validation_first_run_failed.log이다. 점 탐색 및 P1 barycentric 평가를 수정했다.
- 병행 중간 validation/CLI 실행 일부는 도구가 exit 1을 반환했고 traceback 없이 로그가 끝났다. 정확한 종료 원인은 확정하지 못했다. 수치 validation은 이후 단독 실행에서 완료되어 validation_m3.json을 작성했고 CLI 두 예제도 정상 완료했다. 이전 validation_run.log는 중간 기록으로서 성공 증거로 사용하지 않는다.
- CLI NumPy int32 JSON 오류는 cli_first_serialization_failure.log에 있으며 Python int 반환으로 수정했다.
- 최종 전체 pytest의 stdout 파일이 일부만 기록된 문제 때문에 직접 tool stdout을 수집한 후 pytest_final.log에 저장했다. 이 최종 실행은 87 passed, 22 warnings, 56.70 s이다. XML은 정밀 실행시간과 모든 test case를 포함한다.
- 경고 22개 중 8개는 기존 M1의 유한 클래딩 절단 경고이고, 14개는 matplotlib와 Pyparsing의 폐기 예정 API 경고이다. 억제하거나 무시하는 설정으로 제거하지 않았다.
