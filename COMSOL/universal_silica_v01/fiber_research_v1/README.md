# LSCNS Fiber Research v1 — STEP 1–3

**Prototype source code; not numerically certified.** The existing web scalar radial FEM, browser Yee-FDFD, web geometry and project schema 3 remain unchanged. New Python solver and design JSON are intentionally separate.

## Python 3.11 quick start

    cd COMSOL/universal_silica_v01/fiber_research_v1
    python -m pip install -r requirements.txt
    python -m unittest discover -s tests -v
    python cli.py presets --out presets
    python cli.py validate presets/index_guiding__step.json
    python cli.py mesh presets/index_guiding__step.json --out out_step
    python cli.py solve presets/index_guiding__step.json --out out_step --modes 2
    python cli.py study presets/anti_resonant__nanf.json --out out_nanf_study --modes 1

With package import from LSCNS root, use python -m COMSOL.universal_silica_v01.fiber_research_v1.cli. Source-file execution also works.

## Model definition and coverage

25 versioned design presets, eight fiber families A-H: index-guided 5, solid PCF 5, PBGF 2, Kagome 1, anti-resonant HCF 5, Bragg 1, MCF 3, polarization-maintaining 3. All parameters are engineering starting values, not qualified product designs. Save/load via CLI and schema.py; validation rejects invalid nesting and unsupported dopant Sellmeier claims.

Geometry uses gmsh OCC fragments, material physical groups, circular PML annulus, local thin-wall sizing, mesh diagnostics, and outputs geometry.msh, mesh.npz and mesh_report.json. Some variants are **illustrative approximations**: Kagome hypocycloids, conjoined and negative-curvature designs do not yet have exact multi-arc tube interfaces; graded/multilayer presets use piecewise regions; bow-tie stress rods are not thermoelastic calculations. Endlessly-single-mode diagrams and PBGF Bloch bandgap maps are **not implemented**.

Full vector is a 2D HCurl(Nedelec first-kind order 1) transverse E and H1(P1) axial Ez mixed quadratic-beta eigenproblem. It solves with SciPy sparse shift-invert of a companion linearization. Polynomial radial PML uses coordinate-transformed anisotropic eps/mu; SBC is a first-order Robin approximation at a fixed target neff, not exact outgoing DtN. Mode outputs: Re/Im(neff), total dB/km sign-preserved, six complex E/H components sampled at triangle centroids, relative QEP residual, weak Gauss residual, polarization fraction and PML confinement ratio. This **does not** prove all spurious modes have been removed.

Pure silica Sellmeier: Malitson 1965 https://doi.org/10.1364/AO.4.000901. Ambient air n=1.00027 is an uncalibrated assumption, constants are user-specified. Ge/F doped Sellmeier requires a composition-dependent authoritative coefficient table: TODO and explicit error. PM stress-optic anisotropy and bandgap are not solved.

## Acceptance and safety

exp(i beta z - i omega t), ntilde=n+i*kappa. Im(neff)>0 means absorption. Total loss dB/km = (20/ln 10)*2pi/lambda_m * Im(neff)*1000. Negative loss is NOT silently made positive.

A converged linear algebra residual is necessary, not sufficient. For any published HCF/PCF confinement loss, run at least 3 meshes, multiple physical windows, 3 PML thicknesses, 3 strengths and 3 polynomial orders, match modes using field overlap, compare HE11/independent reference, examine boundary energy and Gauss residual. The study command runs numerical sweeps but uses nearest-index matching only; even a sensitivity PASS remains independent-benchmark-unverified. Do not claim COMSOL parity or measured product specifications.

Environment note: gmsh/scikit-fem runtime was not available during source authoring in this conversation. New FEM code is NOT claimed to have executed or passed its physical validation. Run in Colab/local with dependencies and inspect output.

API docs: https://scikit-fem.readthedocs.io/en/stable/api.html and https://gmsh.info/doc/texinfo/gmsh.html
