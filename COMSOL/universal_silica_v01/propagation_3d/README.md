## v0.2 — Native WebGL2 3D Animation (2026-11-04)

[기존 툴 내 실제 3D 모드 실행](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#fiber-3d) · [3D WebGL 모듈](../web/fiber_3d_webgl.mjs) · [UI 사용법](../web/fiber_3d_webgl/README.md) · [구현·검증 한계](../web/fiber_3d_webgl/VALIDATION_REPORT.md).

기존의 2D Canvas 투영 대신 GPU 기반 **WebGL2 xyz 광섬유 단면 압출 및 BPM 강도 3D 광장 시각화**를 추가했습니다. 기존 2D 화면도 전환할 수 있고, HCF 1/2/3중 튜브는 기하학 구조만 표시합니다. **물리적 전파 계산은 기존 2+1D Scalar BPM 그대로**입니다. 광학 주파수의 실제 시간영역 전자기파나 Full-vector 3D Maxwell 해석으로 해석하지 마십시오.

# 3D Fiber Geometry & Propagation — Research v0.1 (2026-11-04)

**[실행: 기존 Universal Silica Fiber Lab의 3D 광섬유·전파](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#fiber-3d)** · [Colab 연구용 전파 노트북](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/universal_silica_v01/propagation_3d/Universal_Silica_Fiber_3D_Propagation.ipynb) · [검증 보고서](VALIDATION_REPORT.md)

## 지원 범위

- **Solid silica Step / Graded / Trench / Multilayer**: 브라우저에서 실제 **2+1D scalar paraxial Split-step FFT Beam Propagation Method (BPM)** 계산. 매 z 스텝에서 복소 2D 전기장 엔벌로프를 FFT 전파시키고, 계산된 강도 단면을 3D 입체 공간에 샘플링해 시각화합니다.
- **3D UI**: 마우스 드래그로 입체 광섬유 회전, 전파 위치 슬라이더/애니메이션, 실제 BPM transverse intensity heatmap, z-MFD 그래프, 파워 오차·가장자리 에너지 확인, PNG/CSV/JSON 저장.
- **HCF/Nested Tube**: 기존 `web/hcf_designer.mjs`의 코어 직경·튜브 수(4–12)·중첩 단계(1–3)·튜브 직경 비율을 사용해 3D 압출 도식을 표시합니다. **HCF 실제 전파, 반공진 누설손실(dB/km)은 계산하지 않으며 Run 버튼이 비활성화**됩니다.
- **별도 Python/Colab**: `propagation_3d.py`, `Universal_Silica_Fiber_3D_Propagation.ipynb`가 NumPy FFT 전파 및 3D 강도 슬라이스, CSV/JSON/NPZ를 생성합니다.

## 사용 순서

1. [기존 웹페이지](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#fiber-3d)에서 **3D 광섬유 구조 및 광 전파**로 이동합니다.
2. Solid silica 프로파일과 파장, 코어 반경, 굴절률 차, 전파 거리, Gaussian waist/오프셋, FFT Grid 및 Z Steps를 설정합니다.
3. **BPM 전파 계산**을 클릭합니다. 3D 회전과 z 슬라이더로 광장 변화를 살펴봅니다. 결과는 입사 Gaussian의 **전파 중 변화**입니다. **정지상태 LP01 고유모드 MFD/Aeff 사양이 아닙니다.**
4. **JSON/CSV/3D PNG**로 결과를 저장합니다. 웹에서 바꾼 파라미터나 결과는 **자동으로 GitHub에 올라가지 않습니다**.
5. HCF를 고르면 3D 튜브 구조만 바뀝니다. 정밀 HCF 공진/손실은 Full-vector complex FEM/PML 연구가 별도로 필요합니다.

## 수학 모델과 정확도

Scalar envelope `A(x,y,z)`, `E=A exp(i k0 nref z)`:

```text
∂A/∂z = i/(2 k0 nref) ∇²_transverse A
       + i k0 (n(x,y)² - nref²)/(2 nref) A
```

연산자 분할: 반 스텝 위상 → 횡방향 FFT 회절 → 반 스텝 위상.

```text
Aeff(z) = ( ∫∫ |A|² dxdy )² / ∫∫ |A|⁴ dxdy
MFD_2nd(z) = 2 sqrt( 2 * ∫∫ |r-rcentroid|² |A|² dxdy / ∫∫ |A|² dxdy )
```

순수 실리카 Sellmeier + 상수 Δn 가정. x,y,z 단위는 µm, λ는 nm. 주기적 FFT 횡방향 경계이며 PML이 아닙니다. 경계에 닿으면 wrap-around 오류가 발생할 수 있으므로 계산 영역·횡방향 Grid·Z Steps를 바꿔 수렴성을 확인해야 합니다. **실수 굴절률 계산에서 광 파워 보존은 알고리즘 특성이지 물리적인 감쇠가 0이라는 실험 결과가 아닙니다.**

> **범위 주의:** 광장이 x,y,z 좌표를 따라 실제 계산되지만 **Full-vector 3D Maxwell FEM/FDTD는 아닙니다**. 비선형, 산란, 반사/역방향 파, 물질 손실, 굽힘, 편광 결합, PML 방사손실, HCF/NANF 구조 손실은 미지원입니다.

## 재현

```bash
python -m pip install numpy matplotlib
python COMSOL/universal_silica_v01/propagation_3d/propagation_3d.py
node COMSOL/universal_silica_v01/web/tests/test_fiber_3d.mjs
```

[기존 Full-vector 2D 단면 FDFD](../full_vector/)와 본 z-전파 BPM은 **서로 다른 해석기**이며, 두 결과를 동일한 Full-vector 3D FEM으로 표시하지 않습니다.
