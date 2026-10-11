# M4 3D Maxwell 해석 검증 결과보고서

실제 pytest (핵심 12개 + 회귀·추가 92개 분리 실행): 104개, 실패 0개, 오류 0개. 두 실행의 경과시간 합 1031.932 s.

검증 실행 UTC 2026-10-11T02:50:18.336677+00:00; 경과시간 485.850815 s. 성능 benchmark가 아니다.

## 지배방정식과 약형식


Lengths are micrometres, time exp(-i omega t), propagation axis z. Isotropic passive dielectric n, mu_r=1. TE means E_z=0 and TM means H_z=0 in the uniform rectangular guide; unlike M3 these are propagation-axis labels.

curl curl E - k0^2 n^2 E = f, k0=2 pi/lambda0; H=curl E/(i k0 Z0).
PEC: n_hat cross E=0. Dielectric interfaces: tangential E,H continuous; weak form handles normal displacement through the gradient test space.

PML: s_j=1+i sigma_j, sigma_j=strength (distance/thickness)^power, positive imaginary outgoing-wave attenuation. J=diag(s_x,s_y,s_z), T=det(J) J^-1 J^-T. epsilon_PML=n^2 T, mu_PML=T. a(E,v)=int (T^-1 curl E).curl(conj(v))-k0^2 (n^2 T E).conj(v). Outer PML truncated with PEC. Physical H formula is used only where J=I.

Rectangular real uniform port width a,height b: kc^2=(m pi/a)^2+(n pi/b)^2, beta=sqrt((k0 np)^2-kc^2), branch Re beta>=0, Im beta>=0. Cutoff rejected. TE: e_t proportional to -z_hat cross grad[cos(m pi x/a) cos(n pi y/b)], m+n>0. TM: e_t proportional to grad[sin(m pi x/a) sin(n pi y/b)], m,n>0. Normalize int |e_t|^2=1. gamma_TE=beta, gamma_TM=k0^2 np^2/beta. TE modes (m,0),(0,n) have norm^2=kc^2 a b/2; both nonzero TE and all TM have norm^2=kc^2 a b/4.

For outward radiation n_hat cross curl E=-i sum gamma_m c_m e_m. Incident wave gives -i sum gamma_m (c_m-2 a_m)e_m. Weak form: a(E,v)-i sum gamma_m c_m conj(d_m)=int f.conj(v)-2i sum gamma_m a_m conj(d_m). S_mi=(c_m-a_m) sqrt(Re gamma_m/Re gamma_i) for propagating channels only. Port trace is restricted to a finite analytical TE/TM span via oriented edge line-integral interpolation; all modes up to caller's (m_max,n_max), including evanescent modes, are retained. Unretained trace modes are zero, NOT an exact complete DtN. Independent mesh and port-cutoff checks required. Final consistent port assembly uses C_mj=int e_m dot Pi_h e_j on the actual FE trace. E=R u, c=C u_port, A_port=-i C^T diag(gamma) C, rhs_port=-2i C^T diag(gamma) a. The original identity-C shortcut failed the immutable TM slope contract; C is now assembled with facet quadrature.

Discretization: conforming gmsh structured tetrahedral extrusion; coordinate spacing <= h/sqrt(3), actual maximum diameter measured <= nominal h. The HXT alternative did not satisfy the immutable coarse-grid TE slope criterion and was rejected. Use degree-7 positive-weight strictly interior tetrahedron quadrature, lowest-order Nedelec first kind edge element, one oriented circulation per edge; compatible grad(P1) subset N1. E coefficients are NOT nodal values. Sparse reduction E=R u imposes PEC and modal traces, A_red=R^T A R - i C^T diag(gamma) C on modal block. Real interpolation R means transpose equals conjugate transpose. General source PEC problems retain all non-PEC edges. Discrete weak Gauss residual is measured as ||G^T(M E+f/k0^2)|| / ||M E+f/k0^2||, with interior P1 gradient G (equivalent to G^T(A E-f)=0 through G^T K=0): for source-free guide this also gives G^T epsilon E=0. No eigenmode-spectrum completeness claim. Complex sparse LU with geometric nested dissection of the structural graph; vertex separators remove all cross-partition adjacency, default numerical pivot threshold 1 is retained. Fields and flux quadrature saved, PML transformed fields distinguished.

Independent tests (written before maxwell3d.py):
* Volume manufactured E=(sin(pi y/b)sin(pi z/c),0,0), div E=0, curlcurl E=pi^2(b^-2+c^-2) E. PEC holds on all faces. f=[pi^2(b^-2+c^-2)-k0^2 n^2]E. Field relative L2 three meshes expected O(h), fine <0.10, slopes .75..1.35.
* Uniform guide TE10 and TM11, exact forward vector fields, S11=0,S21=exp(i beta L). Three meshes, fine complex error <.025, field <.13, S slopes 1.4..2.6.
* Layered guide E_t and gamma(E_f-E_b) continuity, independent 2x2 transfer matrices, TE10/TM11 complex r,t errors <.025; power balance <1e-9 for real and positive absorption for complex n.
* One-port PML terminated guide analytic R=-exp(2 i beta [L+i strength*t/(r+1)]); complex R absolute error <.025; increased strength attenuation tested against independent formula, not zero-reflection assumption.
* Modal matrix symmetry and lossless unitarity <1e-9 are algebraic structural checks, not independent field accuracy.
* Mesh/PML validation, weak Gauss <1e-9, E/H units, NPZ, edge-gradient exact-sequence and three-axis tensor checks.

Sources: Monk and Zhang, Finite Element Methods for Maxwell's Equations (2019), https://arxiv.org/abs/1910.10069 ; official scikit-fem ElementTetN1 source; DOLFINx 0.9 Maxwell/PML demo https://docs.fenicsproject.org/dolfinx/v0.9.0/python/demos/demo_pml.html ; COMSOL Waveguide Adapter https://doc.comsol.com/6.4/doc/com.comsol.help.models.rf.waveguide_adapter/waveguide_adapter.html . The port reduction here is derived above, not a reproduction of COMSOL implementation.

Additional independent H validation: TE10 Ey=(mx/N) sin(mx*x) exp(i beta z), mx=pi/a, N=mx sqrt(ab/2). Hx=-beta Ey/(k0 Z0), Hy=0, Hz=mx^2 cos(mx*x) exp(i beta z)/(i k0 Z0 N). Three-grid physical H L2 finest <.13, observed slopes .75..1.35 (same first-order curl-space expectation), while retaining the existing 1e-13 curl/formula consistency assertion. This is an additional analytic test; no original frozen acceptance tolerance changed.


## 메시 3단계 실제 수렴

| 문제 | 입력 h μm | 실제 hmax μm | 사면체 수 | 자유 DOF | 필드 상대 L2 | 복소 S21 오차 | 필드 p | S21 p |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
|제조해|0.2|0.173205080757|1260|1167|2.926687551415e-01|—|0.980864606|—|
|제조해|0.1|0.094662741608|7722|7965|1.618138106546e-01|—|1.013559208|—|
|제조해|0.05|0.048702479084|56700|62155|8.250381076313e-02|—|—|—|
|TE|0.2|0.181827458020|1260|1191|2.330031744245e-01|0.011858091122561048|1.023718431|1.789716205|
|TE|0.1|0.096576045100|8316|8623|1.219141733967e-01|0.0038213654853750745|1.003008918|1.621435786|
|TE|0.05|0.049033233623|63504|69769|6.177170017435e-02|0.0012732161049577443|—|—|
|TM|0.2|0.181827458020|1260|1191|2.100259319297e-01|0.00933145903181832|1.022980936|2.302985678|
|TM|0.1|0.096576045100|8316|8623|1.099430895599e-01|0.0021732615799440512|1.017106027|1.446449262|
|TM|0.05|0.049033233623|63504|69769|5.517640520840e-02|0.000815282060692522|—|—|

## 독립 TE10 자기장 세 메시 검증

기준: Hx=-β Ey/(k0 Z0), Hy=0, Hz=mx² cos(mx x) exp(iβz)/(i k0 Z0 N), mx=π/a, N=mx√(ab/2). 기준 Z0는 SciPy CODATA 값이며 반환 curl을 기준값으로 사용하지 않았다.

| 입력 h | 실제 hmax | H 상대 L2 오차 | 다음 p |
|---|---:|---:|---:|
|0.2|0.181827458020|2.216047098208e-01|1.041733164|
|0.1|0.096576045100|1.146360141793e-01|1.032946805|
|0.05|0.049033233623|5.691717705512e-02|—|

## 유전체 층

| 모드 및 흡수 | S11 오차 | S21 오차 | FEM 출력전력비 | 해석 출력전력비 |
|---|---:|---:|---:|---:|
|TE_0|1.587656346669e-03|2.038344534634e-03|1.000000000000|1.000000000000|
|TM_0|6.449465743472e-04|7.964627237747e-04|1.000000000000|1.000000000000|
|TE_0.03|1.551522119343e-03|2.000667048516e-03|0.929256796200|0.929525640575|
|TM_0.03|6.557706928510e-04|8.056942582489e-04|0.906335424270|0.906392010032|

## PML 유한 종단

| 모드 | 강도 | FEM 반사 진폭 | 해석 반사 진폭 | 복소 오차 |
|---|---:|---:|---:|---:|
|TE|0|9.999230941786e-01|1.000000000000e+00|3.608186196260e-03|
|TE|1.2|2.142303263734e-01|2.133665241648e-01|1.378253586333e-03|
|TE|4|6.725574950884e-03|5.804359120310e-03|1.360670371203e-03|
|TM|0|9.999704717310e-01|1.000000000000e+00|1.284933389869e-03|
|TM|1.2|4.079091813553e-01|4.072265370428e-01|2.771199957967e-03|
|TM|4|5.337159685587e-02|5.005570363124e-02|5.980718032211e-03|

## 테스트 선작성 및 실패 수정


1. Before maxwell3d.py existed, independent oracle 1 passed and absent-solver 11 failed. Hash manifest records the immutable initial oracle and acceptance test files.
2. First implemented test run: 6 passed, 6 failed in 265.80 s. Field oracles received a scikit-fem DiscreteField ndarray subclass. np.zeros_like and item assignment did not construct the intended oracle vector; the same subclass affected in-place PML stretches. Fix: give public field/material/source callbacks plain numpy arrays; return plain arrays for coordinates.
3. First quadrature rule was degree 4 (11 points with one negative weight). Serialization test required positive quadrature weights. Changed to degree 5 (14 positive-weight points), increasing integration order. Original test/tolerance unchanged.
4. Corrected-array manufactured test still failed at input h=.05: field relative L2=.1482433938482849 >.10. Actual hmax=.08660254037860067. The direct three resolutions h=.2/.1/.05 yielded errors .46910574003557/.29266894774310315/.1482433938482849, indicating lowest-order edge field resolution error. Coordinate mesh spacing was changed from h to h/sqrt(3), so actual tetrahedral diameter <= nominal h. Original analytic tolerance unchanged.
5. A concurrent diagnostic run at h=.025 returned exit 137 before final result. No h=.025 value is reported. Memory pressure was observed while several 3D LU runs overlapped; exact termination source was not independently established. Final numerical checks run serially with symmetric MMD_AT_PLUS_A sparse ordering. A separate fixed manufactured test passed in 62.96 s; final suite is the authoritative completion record.
6. Some helper process listings returned 'fatal library error, lookup self'; process listings were not treated as numerical evidence.
7. The first refined suite used the old COLAMD ordering cached before the performance change. It passed the manufactured and TE10 cases, then was interrupted with exit 130 during the expensive TM11 factorization. It is not a completed suite and its partial log is retained. Final checks use MMD_AT_PLUS_A ordering and serial execution.
8. Rule audit found that degree-5's 14-point rule samples tetrahedral edges (minimum barycentric coordinate 0). A discontinuous callable index has ambiguous values exactly on a material interface even for aligned cells. Before accepting layer results, increased to degree 7, whose 24 points have positive weights and strictly interior barycentric coordinates (minimum .032986329573172934). No tolerance changed. The partial final suite was intentionally interrupted and retained separately; only the subsequent completed run is authoritative.
9. validate_m4.py reuses the same assembled 3D operator/LU for TE10 and TM11 incident right-hand sides at fixed geometry, material and PML. Assembly, solve and extraction use the tested maxwell3d kernel; independent continuous references are never used in assembly. This reduces repeated factorizations; reported elapsed time is not a performance comparison.
10. A completed degree-7 full suite stopped at TM11 slope: 55 passed, 1 failed in 216.57 s. Original TM errors [.007894160123914508,.0020842945326227354,.0008095748922152732] at actual hmax [.1818274580197363,.09657604509991888,.049033233622777] gave slopes [2.10468625,1.39514783]; second value failed immutable >1.4 threshold. Boundary audit found the identity-C shortcut was not a consistent projection of the interpolated Nedelec trace onto the analytic port modes. Fixed by facet-assembled C_mj=int e_m dot Pi_h e_j, port matrix -i C^T Gamma C, incident RHS -2i C^T Gamma a, and physical outgoing modal coefficients c=C u_port. Tests and tolerances remain unchanged. Corrected-run numbers are in the final validation JSON. A small LU pivot experiment is retained; default pivot threshold 1 was retained for numerical robustness, so it is not a solver change.
11. Consistent port-projection regression: TE10 and TM11 three-grid tests both passed (2 passed in 156.94 s). Default pivot threshold remained 1; original hashes/tolerances unchanged. Validation explicitly releases each previous prepared LU before creating the next operator to avoid overlapping large 3D factorizations.
12. The consistent-projection full suite completed 60 test cases, including all TE/TM layer cases, then terminated at PML with exit 137 and no complete XML. Cgroup memory.max=8589934592 bytes (8 GiB); memory.events recorded oom_kill=2 after the two exit-137 runs. This establishes active memory pressure/OOM events, although no per-process kernel attribution log was available. The terminated stdout is retained as pytest_PML_memory_terminated.log. No PML pass or numerical value is inferred from that incomplete run. Sparse factorization fill is being corrected; test tolerances remain immutable.

13. Default structured extrusion PML LU exceeded the 8 GiB runtime memory limit. Before changing meshing, recorded the alternative discretization in M4_FORMULATION.md. OCC-fragmented HXT tetrahedra preserve all specified interface planes. Initial target 0.8h probe gave manufactured fine field error .10291715408899126 > .10; it was not accepted. Target 0.7h is now checked against the immutable original tests. Actual hmax is measured, not claimed to be <= nominal h. Default LU pivot threshold remains 1; trial pivot reductions did not materially reduce cost.

14. HXT target 0.7h completed the manufactured test but failed the TE10 first-grid S slope: errors [.022471237758912826,.009825349885255701,.0023110168164291677], actual hmax [.2888235852586881,.15513930715664048,.08565543714496167], slopes [1.33110257,2.4365331]. Completed run: 1 failed, 2 passed in 387.91 s. Ctrl-C was requested only after the process had already completed; this is an actual analytic failure, not an interrupted pass. Nominal target-size refinement did not put the irregular coarse grid in the required asymptotic range. HXT meshing was rejected, restoring the previously accepted h/sqrt(3) structured geometric refinement.
15. Structural-graph geometric nested dissection recursively removes the smaller vertex separator crossing a median coordinate plane, then orders leaves by RCM; SuperLU uses NATURAL ordering with unchanged pivot threshold 1. Coarse independent ordering comparison: MMD 1.52158756 s, ND .168820241 s; relative solution difference 1.122694733951505e-14. Fine unstructured diagnostic ND factorization and solve completed in 6.216730904 s, residual 7.274067057638645e-14. It had no comparison solution, so its stored difference=0 is initialization, not an equivalence test. Full analytic tests now recheck the structured mesh with this ordering.

16. Regression/additional run after restoring structured mesh with nested-dissection LU: 92 passed, 22 warnings in 74.18 s. Includes all 87 M1–M3 tests plus 5 additional M4 checks. Original analytic core 12 are run separately; final summary is an explicitly labelled aggregate of disjoint test selections, not a claimed single process. M1–M3 common solver files remain byte-for-byte unchanged.

17. Large analytic-core factorizations are sequential. The 92-case regression/additional run used a separate process concurrently with the core run; those smaller operators fit the memory limit. validate_m4.py runs large operators sequentially and releases the previous LU before the next. CLI layer/full-S and one-port PML were rerun from /tmp using the installed entry point with the final structured mesh and ND factorization.

18. Structured-mesh ND analytic core completed: 12 passed in 933.87 s (15:33). With the disjoint 92-case regression/additional run: 104 passed, 0 failed/errors/skipped. Derived aggregate XML is labelled; both raw XML/logs and provenance hashes are retained. Cgroup oom_kill remained 2 with no additional OOM event. Original test hashes unchanged. The first notes-append command used the wrong working-directory-relative path and failed before touching any solver/data; corrected with an absolute path. Validation started independently and continues.

19. Report numerical validator completed successfully in 485.850815077 s (UTC start 2026-10-11T02:50:18.336677+00:00). Geometry: one connected volume, no exposed internal interfaces, .24 um^3 volume exactly. Tables/plots use its actual JSON/CSV/NPZ, never solver outputs as reference values. Independent references remain the analytic oracle. Extra H accuracy check was strengthened to independent TE10 H with three resolutions, preserving the original curl consistency threshold and all frozen M4 core tolerances; the 92-case regression/additional selection is rerun to reflect this test change.

20. Final regression/additional rerun with independent three-grid H test: 92 passed, 22 warnings in 98.07 s. All 104 distinct final tests passed across the disjoint 12+92 selections. validation_H.json is extracted verbatim from the passing test stdout in raw JUnit XML. Previous 92-case result is retained under before_independent_H filenames; it is not the final source.


## 검증된 범위

3D PEC 제조해, TE10/TM11 벡터 필드·복소 S의 독립 해석해와 세 메시 수렴, 유전체 층의 실수·흡수 전달행렬, z 방향 유한 PML 반사, 엣지 구배 호환성과 약한 Gauss 조건, 3축 tensor 대수, 독립 TE10 H의 세 메시 수렴, E/H·NPZ·CLI 및 M1–M3 회귀. 3D patch는 모드 변환과 reciprocity/unitarity 구조 시험이며 정량 정확도는 별도 검증하지 않았다.


## 검증되지 않은 범위

실제 3D 개방 광섬유의 수치 포트·누설손실, 6면 개방 Mie 구 산란, 층·PML 사례별 세 메시 차수와 두께 수렴, 임의 bend/taper/patch의 정량 오차, 상용 광섬유·COMSOL 비교, 이방성·금속·자성·이득·강공진·MPI scalability.


## 알려진 한계

최저차 Nédélec 필드 약 1차, 복소 S 차수는 benchmark별 실제 관측값이며 일반 형상에 보장되지 않는다. 포트는 finite analytical modal trace span으로 완전 DtN이 아니며 미포함 성분을 0으로 제한한다. 실수 균일 직사각형 포트만 지원하고 cutoff/횡방향 PML 포트는 지원하지 않는다. 포트 인접 구조의 z 균일성을 caller가 확보해야 한다. 필드 진폭은 1 W가 아니다. PEC benchmark는 실제 광섬유가 아니다. Sparse LU와 gmsh 단일 프로세스, 미세 3D 메시의 메모리 비용이 크다.


## 원시 실행 데이터

아래 JSON은 실제 실행 파일과 동일하며 전력·차수·복소 계수의 정확한 원시 수치이다.

```json
{
  "executed_utc": "2026-10-11T02:50:18.336677+00:00",
  "versions": {
    "python": "3.11.17",
    "numpy": "2.2.6",
    "scipy": "1.15.3",
    "skfem": "12.0.2",
    "gmsh": "4.14.1",
    "matplotlib": "3.10.3"
  },
  "manufactured": {
    "rows": [
      {
        "h_nominal_um": 0.2,
        "field_l2_error": 0.29266875514149365,
        "tetrahedra": 1260,
        "vertices": 336,
        "edge_dofs": 1809,
        "free_dofs": 1167,
        "h_max_um": 0.173205080757195,
        "linear_residual": 1.1293744485260982e-14,
        "weak_gauss_residual": 8.861560865352496e-15,
        "gradient_curl_residual": 2.2894220599590097e-18
      },
      {
        "h_nominal_um": 0.1,
        "field_l2_error": 0.16181381065457476,
        "tetrahedra": 7722,
        "vertices": 1680,
        "edge_dofs": 10119,
        "free_dofs": 7965,
        "h_max_um": 0.09466274160784169,
        "linear_residual": 4.624483297886957e-14,
        "weak_gauss_residual": 4.68066714566482e-14,
        "gradient_curl_residual": 8.728049197476412e-19
      },
      {
        "h_nominal_um": 0.05,
        "field_l2_error": 0.08250381076312768,
        "tetrahedra": 56700,
        "vertices": 10868,
        "edge_dofs": 70273,
        "free_dofs": 62155,
        "h_max_um": 0.04870247908365109,
        "linear_residual": 1.9936616794801315e-13,
        "weak_gauss_residual": 1.790691681430755e-13,
        "gradient_curl_residual": 3.412654277747247e-19
      }
    ],
    "field_slopes": [
      0.9808646062832422,
      1.013559207532603
    ]
  },
  "straight": {
    "TE": {
      "rows": [
        {
          "h_nominal_um": 0.2,
          "S21": [
            -0.9942628741461166,
            -0.0880707363083197
          ],
          "S11": [
            0.0003495300529332557,
            -3.493056581285367e-05
          ],
          "S21_error": 0.011858091122561048,
          "field_l2_error": 0.23300317442450166,
          "tetrahedra": 1260,
          "vertices": 336,
          "edge_dofs": 1809,
          "free_dofs": 1191,
          "h_max_um": 0.1818274580197363,
          "linear_residual": 7.632060192397654e-15,
          "weak_gauss_residual": 9.710445197811011e-15,
          "gradient_curl_residual": 2.2179275968102036e-18,
          "port_trace_modes": 24,
          "propagating_channels": 8,
          "output_power": 1.0000000000000007,
          "max_orders": [
            2,
            2
          ],
          "port_projection_deviation_norm": 0.37249272377014625
        },
        {
          "h_nominal_um": 0.1,
          "S21": [
            -0.9966200808303697,
            -0.08034396153705473
          ],
          "S11": [
            3.481711652963959e-05,
            -0.00015811438127280277
          ],
          "S21_error": 0.0038213654853750745,
          "field_l2_error": 0.12191417339672282,
          "tetrahedra": 8316,
          "vertices": 1800,
          "edge_dofs": 10873,
          "free_dofs": 8623,
          "h_max_um": 0.09657604509991888,
          "linear_residual": 2.0569723144436737e-14,
          "weak_gauss_residual": 3.941705239777012e-14,
          "gradient_curl_residual": 8.699907673609664e-19,
          "port_trace_modes": 24,
          "propagating_channels": 8,
          "output_power": 0.9999999999999998,
          "max_orders": [
            2,
            2
          ],
          "port_projection_deviation_norm": 0.10753889952176511
        },
        {
          "h_nominal_um": 0.05,
          "S21": [
            -0.9969583904493792,
            -0.07781735802947427
          ],
          "S11": [
            5.15025167135974e-06,
            -4.990623983560173e-05
          ],
          "S21_error": 0.0012732161049577443,
          "field_l2_error": 0.06177170017434623,
          "tetrahedra": 63504,
          "vertices": 12122,
          "edge_dofs": 78565,
          "free_dofs": 69769,
          "h_max_um": 0.049033233622777,
          "linear_residual": 6.95575023960071e-14,
          "weak_gauss_residual": 1.7783433942380818e-13,
          "gradient_curl_residual": 3.3225900989621166e-19,
          "port_trace_modes": 24,
          "propagating_channels": 8,
          "output_power": 1.0000000000000004,
          "max_orders": [
            2,
            2
          ],
          "port_projection_deviation_norm": 0.02886592049978908
        }
      ],
      "exact_S21": [
        -0.9970658448258897,
        -0.07654868438866175
      ],
      "exact_beta_per_um": 6.43643258906458,
      "gamma_per_um": 6.43643258906458,
      "transmission_slopes": [
        1.7897162046408035,
        1.6214357856321793
      ],
      "field_slopes": [
        1.0237184307421194,
        1.0030089178619352
      ]
    },
    "TM": {
      "rows": [
        {
          "h_nominal_um": 0.2,
          "S21": [
            -0.28700178149010014,
            0.9546507720099371
          ],
          "S11": [
            -0.062067903141154956,
            -0.01828548514472748
          ],
          "S21_error": 0.00933145903181832,
          "field_l2_error": 0.21002593192974642,
          "tetrahedra": 1260,
          "vertices": 336,
          "edge_dofs": 1809,
          "free_dofs": 1191,
          "h_max_um": 0.1818274580197363,
          "linear_residual": 6.614251116962105e-15,
          "weak_gauss_residual": 1.1531538220035169e-14,
          "gradient_curl_residual": 2.2179275968102036e-18,
          "port_trace_modes": 24,
          "propagating_channels": 8,
          "output_power": 1.0000000000000029,
          "max_orders": [
            2,
            2
          ],
          "port_projection_deviation_norm": 0.37249272377014625
        },
        {
          "h_nominal_um": 0.1,
          "S21": [
            -0.2941897908615086,
            0.9555030897322964
          ],
          "S11": [
            -0.0168292652772547,
            -0.005154543419081987
          ],
          "S21_error": 0.0021732615799440512,
          "field_l2_error": 0.10994308955988258,
          "tetrahedra": 8316,
          "vertices": 1800,
          "edge_dofs": 10873,
          "free_dofs": 8623,
          "h_max_um": 0.09657604509991888,
          "linear_residual": 1.9611370926988907e-14,
          "weak_gauss_residual": 4.8856857512503755e-14,
          "gradient_curl_residual": 8.699907673609664e-19,
          "port_trace_modes": 24,
          "propagating_channels": 8,
          "output_power": 1.000000000000001,
          "max_orders": [
            2,
            2
          ],
          "port_projection_deviation_norm": 0.10753889952176511
        },
        {
          "h_nominal_um": 0.05,
          "S21": [
            -0.29553966405759297,
            0.9553137658005052
          ],
          "S11": [
            -0.004543248252513177,
            -0.0014041710225051476
          ],
          "S21_error": 0.000815282060692522,
          "field_l2_error": 0.05517640520840283,
          "tetrahedra": 63504,
          "vertices": 12122,
          "edge_dofs": 78565,
          "free_dofs": 69769,
          "h_max_um": 0.049033233622777,
          "linear_residual": 5.367356782524976e-14,
          "weak_gauss_residual": 2.023503341425071e-13,
          "gradient_curl_residual": 3.3225900989621166e-19,
          "port_trace_modes": 24,
          "propagating_channels": 8,
          "output_power": 1.0000000000000009,
          "max_orders": [
            2,
            2
          ],
          "port_projection_deviation_norm": 0.02886592049978908
        }
      ],
      "exact_S21": [
        -0.29632300172025633,
        0.955087785835154
      ],
      "exact_beta_per_um": 3.7432735263093986,
      "gamma_per_um": 15.186953598425305,
      "transmission_slopes": [
        2.302985678474179,
        1.4464492616139557
      ],
      "field_slopes": [
        1.022980936296609,
        1.0171060266540504
      ]
    }
  },
  "layer": {
    "TE_0": {
      "S11": [
        0.19590333778608437,
        -0.2711741561619801
      ],
      "S21": [
        -0.7638784397027322,
        -0.5518610367252101
      ],
      "exact_S11": [
        0.19616946696802337,
        -0.27273934878068223
      ],
      "exact_S21": [
        -0.764634480535985,
        -0.5499680891030277
      ],
      "reflection_error": 0.0015876563466692582,
      "transmission_error": 0.0020383445346338473,
      "exact_output_power": 1.0000000000000004,
      "tetrahedra": 89280,
      "vertices": 16800,
      "edge_dofs": 109767,
      "free_dofs": 98727,
      "h_max_um": 0.04377182821012478,
      "linear_residual": 1.4583478288585392e-13,
      "weak_gauss_residual": 2.7735262948319667e-13,
      "gradient_curl_residual": 2.9652865583460526e-19,
      "port_trace_modes": 24,
      "propagating_channels": 8,
      "output_power": 0.9999999999999998,
      "max_orders": [
        2,
        2
      ],
      "port_projection_deviation_norm": 0.022613784128827162
    },
    "TM_0": {
      "S11": [
        0.06398913555866881,
        0.11914513575237429
      ],
      "S21": [
        -0.8728918973742488,
        0.4687917284925217
      ],
      "exact_S11": [
        0.06419294255145384,
        0.11975703345089433
      ],
      "exact_S21": [
        -0.8731914045833853,
        0.4680537255778676
      ],
      "reflection_error": 0.0006449465743472171,
      "transmission_error": 0.0007964627237747402,
      "exact_output_power": 0.9999999999999999,
      "tetrahedra": 89280,
      "vertices": 16800,
      "edge_dofs": 109767,
      "free_dofs": 98727,
      "h_max_um": 0.04377182821012478,
      "linear_residual": 9.555492014849268e-14,
      "weak_gauss_residual": 3.029703688055431e-13,
      "gradient_curl_residual": 2.9652865583460526e-19,
      "port_trace_modes": 24,
      "propagating_channels": 8,
      "output_power": 0.9999999999999994,
      "max_orders": [
        2,
        2
      ],
      "port_projection_deviation_norm": 0.022613784128827162
    },
    "TE_0.03": {
      "S11": [
        0.20127216731265873,
        -0.2532381884538885
      ],
      "S21": [
        -0.7389871937565916,
        -0.5277222279173975
      ],
      "exact_S11": [
        0.20166966194964017,
        -0.25473792806474826
      ],
      "exact_S21": [
        -0.7398483267370809,
        -0.5259163711598083
      ],
      "reflection_error": 0.0015515221193430548,
      "transmission_error": 0.0020006670485158626,
      "exact_output_power": 0.9295256405751625,
      "tetrahedra": 89280,
      "vertices": 16800,
      "edge_dofs": 109767,
      "free_dofs": 98727,
      "h_max_um": 0.04377182821012478,
      "linear_residual": 1.4179765747309394e-13,
      "weak_gauss_residual": 3.134048552505957e-13,
      "gradient_curl_residual": 2.9652865583460526e-19,
      "port_trace_modes": 24,
      "propagating_channels": 8,
      "output_power": 0.9292567962002443,
      "max_orders": [
        2,
        2
      ],
      "port_projection_deviation_norm": 0.022613784128827162
    },
    "TM_0.03": {
      "S11": [
        0.0646987230491276,
        0.11201359658387407
      ],
      "S21": [
        -0.8313336642980348,
        0.44551455677811075
      ],
      "exact_S11": [
        0.06518959276876925,
        0.11244843232290802
      ],
      "exact_S21": [
        -0.8316657878113958,
        0.4447805014602716
      ],
      "reflection_error": 0.0006557706928509916,
      "transmission_error": 0.0008056942582488631,
      "exact_output_power": 0.906392010032438,
      "tetrahedra": 89280,
      "vertices": 16800,
      "edge_dofs": 109767,
      "free_dofs": 98727,
      "h_max_um": 0.04377182821012478,
      "linear_residual": 9.679086563231394e-14,
      "weak_gauss_residual": 3.233417347677093e-13,
      "gradient_curl_residual": 2.9652865583460526e-19,
      "port_trace_modes": 24,
      "propagating_channels": 8,
      "output_power": 0.9063354242696452,
      "max_orders": [
        2,
        2
      ],
      "port_projection_deviation_norm": 0.022613784128827162
    }
  },
  "pml": {
    "TE": [
      {
        "strength": 0,
        "S11": [
          0.6393177570586754,
          0.7688426365526244
        ],
        "exact_S11": [
          0.6421365803491996,
          0.7665902505102944
        ],
        "reflection_amplitude": 0.9999230941786106,
        "exact_reflection_amplitude": 1.0,
        "reflection_error": 0.003608186196260485,
        "tetrahedra": 102312,
        "vertices": 19140,
        "edge_dofs": 125469,
        "free_dofs": 113427,
        "h_max_um": 0.049033233622777,
        "linear_residual": 1.292014888499914e-13,
        "weak_gauss_residual": 2.031279843795218e-13,
        "gradient_curl_residual": 2.622426663768183e-19,
        "port_trace_modes": 12,
        "propagating_channels": 4,
        "output_power": 1.000000000000005,
        "max_orders": [
          2,
          2
        ],
        "port_projection_deviation_norm": 0.02041128813059261
      },
      {
        "strength": 1.2,
        "S11": [
          0.13673843023178422,
          0.16491584046355742
        ],
        "exact_S11": [
          0.13701045018820115,
          0.1635646972100306
        ],
        "reflection_amplitude": 0.2142303263734013,
        "exact_reflection_amplitude": 0.21336652416483365,
        "reflection_error": 0.0013782535863331106,
        "tetrahedra": 102312,
        "vertices": 19140,
        "edge_dofs": 125469,
        "free_dofs": 113427,
        "h_max_um": 0.049033233622777,
        "linear_residual": 8.349964651139501e-14,
        "weak_gauss_residual": 1.945763050089086e-13,
        "gradient_curl_residual": 2.572255028017168e-19,
        "port_trace_modes": 12,
        "propagating_channels": 4,
        "output_power": 0.045902389438232306,
        "max_orders": [
          2,
          2
        ],
        "port_projection_deviation_norm": 0.02041128813059261
      },
      {
        "strength": 4,
        "S11": [
          0.0050869435105518335,
          0.004399586814738311
        ],
        "exact_S11": [
          0.0037271913166346004,
          0.004449565112090213
        ],
        "reflection_amplitude": 0.006725574950884467,
        "exact_reflection_amplitude": 0.0058043591203100755,
        "reflection_error": 0.0013606703712027112,
        "tetrahedra": 102312,
        "vertices": 19140,
        "edge_dofs": 125469,
        "free_dofs": 113427,
        "h_max_um": 0.049033233622777,
        "linear_residual": 9.341693413393872e-14,
        "weak_gauss_residual": 2.3582977718129833e-13,
        "gradient_curl_residual": 2.5635041183677767e-19,
        "port_trace_modes": 12,
        "propagating_channels": 4,
        "output_power": 4.684778086752739e-05,
        "max_orders": [
          2,
          2
        ],
        "port_projection_deviation_norm": 0.02041128813059261
      }
    ],
    "TM": [
      {
        "strength": 0,
        "S11": [
          -0.9567063332655323,
          0.2909534949498756
        ],
        "exact_S11": [
          -0.9571075681746475,
          0.2897328130240214
        ],
        "reflection_amplitude": 0.9999704717310042,
        "exact_reflection_amplitude": 1.0,
        "reflection_error": 0.0012849333898688386,
        "tetrahedra": 102312,
        "vertices": 19140,
        "edge_dofs": 125469,
        "free_dofs": 113427,
        "h_max_um": 0.049033233622777,
        "linear_residual": 1.177691059548274e-13,
        "weak_gauss_residual": 2.495651756774784e-13,
        "gradient_curl_residual": 2.622426663768183e-19,
        "port_trace_modes": 12,
        "propagating_channels": 4,
        "output_power": 1.0000000000000084,
        "max_orders": [
          2,
          2
        ],
        "port_projection_deviation_norm": 0.02041128813059261
      },
      {
        "strength": 1.2,
        "S11": [
          -0.38962567414833593,
          0.12075485198690851
        ],
        "exact_S11": [
          -0.38975960056521797,
          0.11798689011544149
        ],
        "reflection_amplitude": 0.40790918135526866,
        "exact_reflection_amplitude": 0.4072265370428007,
        "reflection_error": 0.0027711999579666,
        "tetrahedra": 102312,
        "vertices": 19140,
        "edge_dofs": 125469,
        "free_dofs": 113427,
        "h_max_um": 0.049033233622777,
        "linear_residual": 9.412573135596742e-14,
        "weak_gauss_residual": 2.6194719321330024e-13,
        "gradient_curl_residual": 2.572255028017168e-19,
        "port_trace_modes": 12,
        "propagating_channels": 4,
        "output_power": 0.16639634041533385,
        "max_orders": [
          2,
          2
        ],
        "port_projection_deviation_norm": 0.02041128813059261
      },
      {
        "strength": 4,
        "S11": [
          -0.04935814204561062,
          0.02030520043611426
        ],
        "exact_S11": [
          -0.04790869277576942,
          0.014502779820976622
        ],
        "reflection_amplitude": 0.05337159685586941,
        "exact_reflection_amplitude": 0.05005570363124254,
        "reflection_error": 0.005980718032211305,
        "tetrahedra": 102312,
        "vertices": 19140,
        "edge_dofs": 125469,
        "free_dofs": 113427,
        "h_max_um": 0.049033233622777,
        "linear_residual": 9.953278875397213e-14,
        "weak_gauss_residual": 2.8494242207415694e-13,
        "gradient_curl_residual": 2.5635041183677767e-19,
        "port_trace_modes": 12,
        "propagating_channels": 4,
        "output_power": 0.0028594090804407034,
        "max_orders": [
          2,
          2
        ],
        "port_projection_deviation_norm": 0.02041128813059261
      }
    ]
  },
  "geometry_checks": {
    "connected_components": 1,
    "internal_exposed_faces": 0,
    "volume_um3": 0.24,
    "exact_volume_um3": 0.24
  },
  "mode_conversion": {
    "channels": [
      [
        "left",
        "TE",
        0,
        1
      ],
      [
        "left",
        "TE",
        1,
        0
      ],
      [
        "left",
        "TE",
        1,
        1
      ],
      [
        "left",
        "TM",
        1,
        1
      ],
      [
        "right",
        "TE",
        0,
        1
      ],
      [
        "right",
        "TE",
        1,
        0
      ],
      [
        "right",
        "TE",
        1,
        1
      ],
      [
        "right",
        "TM",
        1,
        1
      ]
    ],
    "reciprocity_residual": 5.3420440911275784e-15,
    "unitarity_residual": 5.544107596633996e-15,
    "TE10_to_TM11_amplitude": 0.1161242776592288
  },
  "port_cutoff": [
    {
      "max_order": 1,
      "S21": [
        -0.9966228352347082,
        -0.08031452298103259
      ],
      "error": 0.0037918066672324394,
      "tetrahedra": 8316,
      "vertices": 1800,
      "edge_dofs": 10873,
      "free_dofs": 8607,
      "h_max_um": 0.09657604509991888,
      "linear_residual": 2.056407120243299e-14,
      "weak_gauss_residual": 3.97318088261504e-14,
      "gradient_curl_residual": 8.699907673609664e-19,
      "port_trace_modes": 8,
      "propagating_channels": 8,
      "output_power": 0.9999999999999993,
      "max_orders": [
        1,
        1
      ],
      "port_projection_deviation_norm": 0.02029819388224977
    },
    {
      "max_order": 2,
      "S21": [
        -0.9966200808303697,
        -0.08034396153705473
      ],
      "error": 0.0038213654853750745,
      "tetrahedra": 8316,
      "vertices": 1800,
      "edge_dofs": 10873,
      "free_dofs": 8623,
      "h_max_um": 0.09657604509991888,
      "linear_residual": 2.0569723144436737e-14,
      "weak_gauss_residual": 3.941705239777012e-14,
      "gradient_curl_residual": 8.699907673609664e-19,
      "port_trace_modes": 24,
      "propagating_channels": 8,
      "output_power": 0.9999999999999998,
      "max_orders": [
        2,
        2
      ],
      "port_projection_deviation_norm": 0.10753889952176511
    },
    {
      "max_order": 3,
      "S21": [
        -0.9966199387733604,
        -0.08034578188416155
      ],
      "error": 0.0038231899767096496,
      "tetrahedra": 8316,
      "vertices": 1800,
      "edge_dofs": 10873,
      "free_dofs": 8647,
      "h_max_um": 0.09657604509991888,
      "linear_residual": 1.9266131559933738e-14,
      "weak_gauss_residual": 3.616286193269877e-14,
      "gradient_curl_residual": 8.699907673609664e-19,
      "port_trace_modes": 48,
      "propagating_channels": 8,
      "output_power": 1.0000000000000024,
      "max_orders": [
        3,
        3
      ],
      "port_projection_deviation_norm": 0.3036597539847975
    }
  ],
  "elapsed_seconds": 485.8508150769994
}
```


## 독립 H 원시 수치

```json
{
  "rows": [
    {
      "h_nominal_um": 0.2,
      "H_l2_error": 0.22160470982084415,
      "tetrahedra": 1260,
      "vertices": 336,
      "edge_dofs": 1809,
      "free_dofs": 1191,
      "h_max_um": 0.1818274580197363,
      "linear_residual": 7.632060192397654e-15,
      "weak_gauss_residual": 9.710445197811011e-15,
      "gradient_curl_residual": 2.2179275968102036e-18,
      "port_trace_modes": 24,
      "propagating_channels": 8,
      "output_power": 1.0000000000000007,
      "max_orders": [
        2,
        2
      ],
      "port_projection_deviation_norm": 0.37249272377014625
    },
    {
      "h_nominal_um": 0.1,
      "H_l2_error": 0.1146360141793297,
      "tetrahedra": 8316,
      "vertices": 1800,
      "edge_dofs": 10873,
      "free_dofs": 8623,
      "h_max_um": 0.09657604509991888,
      "linear_residual": 2.0569723144436737e-14,
      "weak_gauss_residual": 3.941705239777012e-14,
      "gradient_curl_residual": 8.699907673609664e-19,
      "port_trace_modes": 24,
      "propagating_channels": 8,
      "output_power": 0.9999999999999998,
      "max_orders": [
        2,
        2
      ],
      "port_projection_deviation_norm": 0.10753889952176511
    },
    {
      "h_nominal_um": 0.05,
      "H_l2_error": 0.056917177055115994,
      "tetrahedra": 63504,
      "vertices": 12122,
      "edge_dofs": 78565,
      "free_dofs": 69769,
      "h_max_um": 0.049033233622777,
      "linear_residual": 6.95575023960071e-14,
      "weak_gauss_residual": 1.7783433942380818e-13,
      "gradient_curl_residual": 3.3225900989621166e-19,
      "port_trace_modes": 24,
      "propagating_channels": 8,
      "output_power": 1.0000000000000004,
      "max_orders": [
        2,
        2
      ],
      "port_projection_deviation_norm": 0.02886592049978908
    }
  ],
  "H_slopes": [
    1.0417331635444307,
    1.0329468054233675
  ],
  "reference_Z0_ohm": 376.730313412
}
```
