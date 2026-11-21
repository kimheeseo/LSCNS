# 1단계 — V2/V3 원자료 조사 및 V3 정규화 (2026-11-21)

상태: 구현·조사 완료, **외부 원자료 확보는 부분 완료**. V2/V3 = NEEDS_EVIDENCE.
기준 커밋 b9105dde3d7155a5cca88a692620825b6eb5dcb6. 수치 원자료를 만들어 채우지 않았다.

## 확보·미확보

| 항목 | 확보 | 미확보 / 다음 입력 |
|---|---|---|
| V2 | COMSOL 6.4 공식 표의 10개 복소 neff, six-hole 형상 구성과 변수명, Application Library 경로 | 원본 microstructured_optical_fiber_parameters.txt, wl/d/Delta/rb/nbg/nair/wlr/neffMin/neffMax 값, 동일 버전 장 및 원시 모델 |
| V3 | 공식 파라미터와 축대칭 정의, power-averaged r0 적분식, geometry→power neff 변환 | full-precision bent complex CSV, 실제 ewfd2.rAverage 및 r0, 장·메시·PML export, conformal 비교 오차 정책 |

조사: 공식 6.4 문서·Application Gallery 88741의 Download Files 링크와 공개 웹 검색 및 GitHub code search를 확인했다. 검색 결과는 원본 txt가 아닌 문서의 파일명 인용이었다. 공개 렌더링된 Gallery 다운로드 목록에서는 원본을 얻지 못했다. 설치된 COMSOL 및 원시 실행 결과는 이 실행 환경에 없다. 따라서 원본 다운로드·실행 성공을 주장하지 않는다.

## V3에서 추가 확인한 정확한 관계

- r0 = intop1(r*ewfd2.Poavphi)/intop1(ewfd2.Poavphi). 해당 적분은 revolved geometry 옵션을 해제한다.
- n_eff_geometry = ewfd2.neff.
- n_eff_power = ewfd2.neff*ewfd2.rAverage/r0.
- 문서의 원시 고유값 정의는 lambda=-i*beta*r0이다. raw solver eigenvalue, angular propagation, dimensionless neff를 export 시 명시해야 한다.
- 모델 버전에 따라 UI 출력이 이미 정규화될 수 있으므로 raw eigenvalue와 ewfd2.neff를 혼동하지 않는다. 코드에서 representation을 필수로 받는다.
- 약 Rb+4 um는 설명용 근삿값이며 정밀 r0 입력으로 쓰지 않는다. rAverage도 실측 export 없이 Rb로 고정하지 않는다.
- alpha[dB/m]은 선택한 기준 반경의 호 길이로 변환하므로 geometry/power 두 기준을 함께 보존한다.

`bend_reference.py`는 위 관계·SI 단위·명시적 Im 부호 변환을 구현한다. 시험의 숫자는 합성 fixture이며 COMSOL 실제 export가 아니다.

## 외부 export 요구

V2: 원본 txt와 COMSOL 버전, 출처 URL, SHA256, 입력별 단위, PML 구성 및 10모드 ID/장 대응.
V3: wavelength_nm, mode_id, complex ewfd2.neff, ewfd2.rAverage[m], r0[m], raw eigenvalue 및 단위/정의, 시간 관례, geometry/mesh/PML manifest. 실제 비교 전 Re ppm / Im 상대·절대 / 손실 허용오차와 정식화 대응을 선언한다.

물리 gate는 null tolerance·입력 누락·정식화 불일치·모드 미확인을 PASS로 처리해서는 안 된다.

## 출처

- https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.microstructured_optical_fiber/microstructured_optical_fiber.html
- https://www.comsol.com/model/leaky-modes-in-a-microstructured-optical-fiber-88741
- https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.step_index_fiber_bend/step_index_fiber_bend.html (Model Definition; Variables 2; n_eff_geometry; n_eff_power)
