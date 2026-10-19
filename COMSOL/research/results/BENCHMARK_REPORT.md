# Research qualification benchmarks

Created: 2026-10-19T02:44:22.843393+00:00

Raw inputs, meshes, numerical values, environment and source hashes: [benchmarks.json](benchmarks.json).
These are scalar-model verification results. No direct COMSOL run or manufacturer measurement comparison.

| wavelength um | core radius um | delta n | V | finest neff error | Aeff rel error | MFD rel error | gate |
|---:|---:|---:|---:|---:|---:|---:|---|
|1.31|3.5|0.004|1.8073|1.75e-08|1.68e-05|7.54e-06|PASS|
|1.55|4.1|0.005|1.9989|1.85e-08|1.13e-05|5.54e-06|PASS|
|1.625|5.0|0.003|1.7999|6.22e-09|7.66e-06|2.34e-06|PASS|
|1.55|5.5|0.002|1.6950|3.27e-09|7.45e-07|3.48e-05|PASS|

Radial acceptance: |delta neff| <2e-7, relative Aeff and second-moment MFD errors <5e-4; residual <1e-9. Thresholds fixed in runner.

## Two-core MCF numerical screening

| Pitch um | last mesh change | fixed-spacing box change | screening |
|---:|---:|---:|---|
|12.0|1.679%|0.002%|PASS|
|16.0|0.670%|0.008%|PASS|
|20.0|2.057%|0.029%|PASS|
|24.0|0.193%|0.110%|PASS|

This screening uses a 5% threshold for BOTH mesh and box changes. It is not an error bound. Failure excludes that configuration from quantitative coupling claims.
Lc=lambda/(2 delta neff) is the ideal identical-core complete power-transfer length. It is not deployed-cable crosstalk.

## Dispersion

| core elements | wavelength step um | D ps/(nm km) | difference from Bessel |
|---:|---:|---:|---:|
|80|0.01|16.99280062|-0.000204176|
|80|0.005|16.99280024|-0.000204557|
|80|0.0025|16.99280061|-0.000204186|
|160|0.01|16.99295524|-4.95581e-05|
|160|0.005|16.99295491|-4.98872e-05|
|160|0.0025|16.99295481|-4.99829e-05|
|320|0.01|16.99299371|-1.10878e-05|
|320|0.005|16.99299338|-1.14092e-05|
|320|0.0025|16.99299228|-1.25113e-05|

## Required qualification gates

| Gate | Result |
|---|---|
|radial_four_cases|PASS|
|mcf_mesh_and_domain_screening|PASS|
|radial_domain_neff|PASS|
|radial_domain_Aeff|PASS|
|radial_domain_MFD|PASS|
|radial_outer_tail|PASS|
|dispersion_agreement|PASS|
|dispersion_step|PASS|

## Cartesian mesh diagnostic
The centroid-sampled circular interface does not converge monotonically across grids 41/61/81/121/161. A low residual or a favorable single mesh is insufficient evidence of accuracy.
Use the interface-conforming radial solver for circular-fiber precision studies. Cartesian results remain available as a separate coarse/discretization comparison; they are not promoted to high-precision qualification by these radial gates.
