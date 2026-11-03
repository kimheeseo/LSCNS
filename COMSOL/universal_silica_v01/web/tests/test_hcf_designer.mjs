import test from 'node:test';
import assert from 'node:assert/strict';
import {
 DEFAULT, buildGeometry, wallResonances, wallDetuning, capillaryMetrics,
 silicaIndex,rootWallWavelengthUm,spectrum,tubeCountSweep,
 externalComplexNeffLoss,designSummary,besselJ0,besselJ1
} from '../hcf_designer.mjs';
const eq=(a,b,tol)=>assert.ok(Math.abs(a-b)<=tol,`${a} differs from ${b} beyond ${tol}`);
test('Silica Sellmeier n(1550) independent known',()=>eq(silicaIndex(1.55),1.4440236217,2e-10));
test('Geometry N and nesting rings',()=>{
 for(const n of [4,6,8,12])for(const nested of [1,2,3]){
  const g=buildGeometry({...DEFAULT,tubeCount:n,nesting:nested});
  assert.equal(g.rings.length,n);
  assert.ok(g.rings.every(c=>c.layers.length===nested));
  assert.ok(g.minAdjacentGapUm>=DEFAULT.minNeighborGapUm-1e-8);
  assert.ok(g.rings.every(c=>c.layers.every(x=>x.radius>x.wall)));
 }
});
test('N sweep enforces actual outer tube geometric clearance',()=>{
 const r=tubeCountSweep(DEFAULT);assert.equal(r.length,9);
 assert.ok(r.every(x=>x.valid));assert.ok(r[4].autoFitted);
 assert.ok(r[4].actualD1Um<DEFAULT.coreDiameterUm*DEFAULT.outerTubeRatio);
 assert.equal(r[0].confinementLossDbKm,null);
});
test('Invalid geometry without autofit rejects intersection',()=>assert.throws(()=>buildGeometry({...DEFAULT,tubeCount:12,autoFit:false}),/겹칩니다/));
test('Physical nested tube containment',()=>{
 const g=buildGeometry({...DEFAULT,nesting:3});
 for(const c of g.rings){
  for(let j=1;j<3;j++){
   const a=c.layers[j-1],b=c.layers[j];
   const displacement=Math.hypot(b.x-a.x,b.y-a.y);
   assert.ok(displacement+b.radius+DEFAULT.nestedClearanceUm <= a.radius-a.wall+1e-9);
  }
 }
});
test('ARROW t1 analytical root self-consistency (dispersive silica)',()=>{
 const lam=rootWallWavelengthUm(.58,1,DEFAULT.airIndex);
 eq(lam,1.214314508876,1e-8);
 const rhs=2*.58*Math.sqrt(silicaIndex(lam)**2-DEFAULT.airIndex**2);
 eq(lam,rhs,1e-10);
});
test('ARROW thicker walls -> longer λ resonance, smaller frequency',()=>{
 const a=rootWallWavelengthUm(.4,1),b=rootWallWavelengthUm(.7,1);
 assert.ok(b>a);const pts=wallResonances({...DEFAULT,nesting:3});assert.equal(pts.length,18);
});
test('ARROW score = abs(sin(pi phase)), phase near integer zero',()=>{
 const arr=wallDetuning(DEFAULT);
 assert.equal(arr.length,2);assert.ok(arr.every(x=>x.detuning>=0&&x.detuning<=1));
});
test('Bessel J0/J1 fundamental values',()=>{
 eq(besselJ0(0),1,1e-14);eq(besselJ1(0),0,1e-14);
 eq(besselJ0(2.404825557695773),0,1e-12);
 eq(besselJ1(2.404825557695773),0.519147497289467,1e-11);
});
test('Ideal capillary metrics match independent expected values, explicitly no nesting effect',()=>{
 const a=capillaryMetrics(DEFAULT);
 eq(a.neff,0.9995825630483459,1e-12);
 eq(a.Aeff_um2,383.27525143178116,1e-7);
 eq(a.D_ps_per_nm_km,2.9638547986567185,1e-10);
 assert.equal(a.confinementLossDbKm,null);
 const b=capillaryMetrics({...DEFAULT,tubeCount:12,nesting:3,innerRatio:.5});
 eq(a.neff,b.neff,1e-15);eq(a.Aeff_um2,b.Aeff_um2,1e-9);
});
test('Ideal capillary dispersion finite difference second derivative independent',()=>{
 const m=DEFAULT,h=.001,wl=m.wavelengthNm/1000;
 const n=(x)=>capillaryMetrics(m,x*1000).neff;
 const d2=(n(wl+h)-2*n(wl)+n(wl-h))/(h*h);
 const D=-wl/299792458*1e12*d2;
 eq(capillaryMetrics(m).D_ps_per_nm_km,D,2e-4);
});
test('Spectrum wavelength points and finite diagnostics',()=>{
 const pts=spectrum(DEFAULT,25);assert.equal(pts.length,25);
 assert.equal(pts[0].lambdaNm,DEFAULT.scanMinNm);assert.equal(pts.at(-1).lambdaNm,DEFAULT.scanMaxNm);
 assert.ok(pts.every(x=>Number.isFinite(x.D)&&x.layerDetunings.length===2));
});
test('External complex mode postprocessor',()=>{
 eq(externalComplexNeffLoss(1e-10,1550),3.5209712357,2e-8);
 assert.throws(()=>externalComplexNeffLoss(-1e-8,1550));
});
test('Design payload explicitly excludes predicted loss',()=>{
 const s=designSummary({...DEFAULT,nesting:3});
 assert.equal(s.mode.confinementLossDbKm,null);
 assert.ok(s.warnings.some(x=>x.includes('Loss in dB/km')));
});
