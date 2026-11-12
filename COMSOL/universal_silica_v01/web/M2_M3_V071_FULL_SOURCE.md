# LS Wave Optics v0.7.1 — All new M2/M3 source files

These are the complete contents of the **three new JS files**. Existing M2 n/k chart modules and true Yee eigensolver are reused, not rewritten. See [M2_M3_V071_ACCEPTANCE.md](M2_M3_V071_ACCEPTANCE.md) for the original-file minimal diffs, inputs, reference calculations, timing/RSS, regression tolerances, and gaps.

## 1. `m3_yee_quickstart_v071.mjs`

```javascript
/**
 * M3 v0.7.1 strict Yee entrypoint, additive only.
 * Existing collocated Hx/Hy study and experimental M3–M6 workspace are not
 * modified. This shortcut opens the TRUE staggered Yee FDFD (yee_solver.mjs).
 * PML/HCF/bending are deliberately OFF for its explicit M3-only preset.
 * No calculation runs without a separate user click of the existing vb-run.
 *
 * Flag: window.LSWaveFeatures.m3YeeQuickstart=false before module load.
 */
const ENABLED=globalThis.LSWaveFeatures?.m3YeeQuickstart!==false;
const $=id=>document.getElementById(id);
function open(){
 const api=globalThis.LSWaveVectorWorkbench;
 if(!api||!$("vb-tab-m3"))return false;
 api.show("m3");
 return true;
}
function set(id,value){
 const el=$(id);if(!el)throw Error("Yee M3 control missing: "+id);
 if(el.type==="checkbox")el.checked=Boolean(value);
 else el.value=String(value);
 el.dispatchEvent(new Event("change",{bubbles:true}));
}
/** Explicitly requested by pressing the setup button. Does not start solve. */
function applyM3StepIndex(){
 if(!open())throw Error("Yee workbench is not mounted.");
 // Reference: n_clad = Malitson 1550 nm; Δn = 0.005 (user-selected
 // step-index scenario, NOT the official COMSOL a=8 um case).
 set("vb-kind","radial");
 set("vb-material-source","constant");
 set("vb-wavelength",1550);
 set("vb-radius",4.1);
 set("vb-clad",1.444023621703261);
 set("vb-core",1.449023621703261);
 set("vb-grid",48);
 set("vb-modes",2);
 set("vb-half",18);
 set("vb-subpixel",true);
 set("vb-target","");
 set("vb-kcore",0);
 set("vb-kclad",0);
 set("vb-pml",false);
 set("vb-bend",0);
 const banner=$("vb-m3-quickstart-state");
 if(banner)banner.textContent="M3 단독 설정 완료: radial step-index, Malitson nclad(1550), Δn=0.005, Yee 48×48, PML/κ/굽힘 OFF. 이제 기존 '현재 단면 계산'을 누르세요.";
}
function init(){
 if(!ENABLED||$("vb-m3-quickstart")||!$("vwb"))return;
 const m3Pane=document.querySelector('#vwb [data-vb-pane="m3"]');
 if(!m3Pane)return;
 const article=document.createElement("div");
 article.id="vb-m3-quickstart";
 article.style.cssText="padding:12px;margin:12px 0;border:1px solid #356c81;border-radius:8px;background:#112e44";
 article.innerHTML=[
  '<strong style="color:#baf2eb">M3만 실행 · Yee staggered 2D FDFD</strong>',
  '<p style="font-size:12px;color:#c7dce6">이 경로는 결합된 Hx/Hy collocated 기존 v0.7 연구 탭과 다릅니다.',
  ' <b>yee_solver.mjs</b>의 Maxwell staggered curl 이산화를 사용합니다.',
  ' 스칼라 LP₀₁, 기존 FEM, Colab 비교표와 Ex/Ey/Ez, MFD, Aeff는 결과 탭에서 확인할 수 있습니다.</p>',
  '<button id="vb-m3-preset" type="button" style="padding:8px 12px;cursor:pointer">M3 단독 step-index 조건 설정</button>',
  '<p id="vb-m3-quickstart-state" role="status" style="color:#8de2da;font-size:11px">',
  '설정 버튼은 사용자 입력을 변경하지만 실행하지 않습니다. PML/HCF/굴힘·복소 κ는 이 M3 기준 실험에서 사용하지 않습니다.</p>'
 ].join("");
 m3Pane.querySelector(".vb-grid")?.before(article);
 const nav=document.querySelector("nav.eng-shortcuts");
 if(nav){
  const link=document.createElement("a");
  link.id="m3-yee-quicklink";link.href="#vwb";
  link.textContent="M3 · Yee Full-vector";
  nav.append(link);
  link.addEventListener("click",e=>{
   e.preventDefault();
   if(!open()){link.title="Yee workbench is disabled or unavailable";}
  });
 }
 $("vb-m3-preset").addEventListener("click",()=>{
  try{applyM3StepIndex();}
  catch(error){$("vb-m3-quickstart-state").textContent=String(error?.message||error);}
 });
 globalThis.LSWaveM3YeeQuickstart=Object.freeze({enabled:true,open,applyM3StepIndex});
}
if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",init,{once:true});else init();
```

## 2. `tests/test_m2_m3_acceptance_v071.mjs`

```javascript
/**
 * Acceptance test for existing M2 preview and STRICT M3 Yee solver.
 * Run from web/: node tests/test_m2_m3_acceptance_v071.mjs
 * Only verifies M2/M3; intentionally does not exercise M4 PML/HCF/M5/M6.
 */
import assert from "node:assert/strict";
import {performance} from "node:perf_hooks";
import {PRESETS,solve,silicaIndex} from "../solver.mjs";
import {runM1Quantitative} from "../m1q_engine.mjs";
import {M2_CATALOG,M2_SOURCES,evaluateM2Material,interpolateMeasured,
 kFromAlpha,alphaFromK} from "../m2_preview_engine_v06.mjs";
import {parseIndexTable} from "../material_models.mjs";
import {solveYee,portableYee} from "../yee_solver.mjs";
const m1=runM1Quantitative();
assert.equal(m1.passCount,4,"M1 4 analytic cases must remain PASS");
assert.equal(m1.regressionPassCount,5,"M1 original five FEM presets must remain unchanged");
const pure=evaluateM2Material({kind:"silica"},1550);
assert.equal(pure.n,silicaIndex(1550),"SiO2 must exactly reuse original Sellmeier");
assert.ok(M2_SOURCES.silica.url.includes("JOSA.55.001205"),"Malitson primary reference required");
const values={};
for(const entry of M2_CATALOG){
 const n=evaluateM2Material(entry.model,1550).n;
 assert.ok(n>1&&n<3);
 values[entry.id]=n;
}
assert.ok(values.geO2>values.silica);
assert.ok(values.f_wt<values.silica);
assert.ok(values.p2o5>values.silica);
assert.throws(()=>evaluateM2Material({kind:"f_mol",molPercent:1},1550),/requires user Sellmeier coefficients/);
const rows=parseIndexTable("1400,1.45,0\n1500,1.448,0.00000001\n1600,1.446,0");
assert.equal(interpolateMeasured(rows,1550,"linear").n,1.447);
assert.ok(Math.abs(interpolateMeasured(rows,1550,"pchip").n-1.447)<1e-12);
assert.throws(()=>interpolateMeasured(rows,1650,"linear"),/Out of table range/);
const k=kFromAlpha(0.2,1550);
assert.ok(Math.abs(k-5.6802508913155675e-12)<1e-20);
assert.ok(Math.abs(alphaFromK(k,1550)-.2)<1e-12);

const ncl=silicaIndex(1550),cfg={
 wavelength_nm:1550,core_radius_um:4.1,n_core:ncl+.005,n_clad:ncl,
 window_halfwidth_um:18,grid_n:24,modes:2,geometry:{kind:"radial"},
 pml:{enabled:false,thickness_um:4,strength:3,order:3},
 bend_radius_um:0,k_core:0,k_clad:0,subpixel:true
};
const scalar=solve({...PRESETS.g652,clad_index_override:ncl});
const before=process.memoryUsage();
const t0=performance.now();
const result=solveYee(cfg);
const wallMs=performance.now()-t0,after=process.memoryUsage();
const portable=portableYee(result);
assert.equal(portable.schema,"LSCNS_vector_mode_v1");
assert.equal(result.modes.length,2);
assert.ok(result.modes.every(m=>Number.isFinite(m.neff)&&m.residual<1e-8));
assert.ok(Math.abs(result.modes[0].neff-scalar.n_eff)<5e-4,
 "At coarse N=24, the genuine Yee vector neff should be near scalar LP01, not equal.");
assert.ok(portable.modes.every(m=>m.field_map?.intensity_normalized_to_max_total?.Ez));
assert.ok(result.modes.every(m=>m.Aeff_um2>0&&m.MFD_2nd_um>0));
assert.equal(portable.input.pml.enabled,false);
console.log(JSON.stringify({
 type:"M2M3 v0.7.1 acceptance",
 m1:"4/4 cases + 5/5 original presets PASS",
 m2_1550:values,alpha_0p2_kappa:k,
 yee_n:24,yee_neff:result.modes.map(m=>m.neff),
 scalar_neff:scalar.n_eff,yee_residuals:result.modes.map(m=>m.residual),
 measured_wall_ms:wallMs,process_rss_before_mib:before.rss/1048576,
 process_rss_after_mib:after.rss/1048576,
 note:"Process RSS is measured local Node total, not a browser Worker memory measurement.",
 grid_convergence:"NOT certified by a single N24 run; see existing N24/48/72/96 table"
},null,2));
```

## 3. `tests/browser_m2_m3_v071.mjs`

```javascript
/**
 * Optional Playwright smoke test: strict Yee M3 shortcut, M2 provenance,
 * preserved legacy Colab import and real Worker mode. Not a cloud dependency.
 *
 * Serve LSCNS repository root: python3 -m http.server 8765
 * Run: node tests/browser_m2_m3_v071.mjs
 */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try {
 const page=await browser.newPage({viewport:{width:1560,height:960},acceptDownloads:true});
 const errors=[];page.on("pageerror",e=>errors.push(e.message));
 await page.goto(process.env.TEST_BASE_URL||
  "http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/?v=v0.7.1",
  {waitUntil:"load",timeout:90000});
 if(await page.locator('[data-wizard="cancel"]').isVisible())
  await page.locator('[data-wizard="cancel"]').click();
 await page.waitForFunction(()=>window.LSWaveVectorWorkbench?.show&&window.LSWaveM3YeeQuickstart?.open);
 assert.equal(await page.locator("#m3-yee-quicklink").count(),1);
 assert.equal(await page.locator("#m2v6-nav").count(),1);
 assert.equal(await page.locator("#fv-json-file").count(),1);
 assert.ok((await page.locator("body").textContent()).includes("v0.7.1"));
 await page.locator("#m3-yee-quicklink").click();
 assert.ok(await page.locator('#vwb [data-vb-pane="m3"]').isVisible());
 await page.locator("#vb-m3-preset").click();
 assert.equal(await page.locator("#vb-kind").inputValue(),"radial");
 assert.equal(await page.locator("#vb-pml").isChecked(),false);
 assert.equal(await page.locator("#vb-bend").inputValue(),"0");
 assert.equal(await page.locator("#vb-kcore").inputValue(),"0");
 assert.equal(await page.locator("#vb-kclad").inputValue(),"0");
 assert.equal(await page.locator("#vb-material-source").inputValue(),"constant");
 assert.equal(await page.locator("#vb-grid").inputValue(),"48");
 assert.equal(await page.locator("#vb-modes").inputValue(),"2");
 // Small grid to keep UI smoke inexpensive. Default baseline remains 48.
 await page.locator("#vb-grid").fill("24");
 await page.locator("#vb-run").click();
 await page.waitForFunction(()=>window.LSWaveVectorWorkbench?.latest?.()?.result?.modes?.length===2,
  undefined,{timeout:180000});
 const data=await page.evaluate(()=>window.LSWaveVectorWorkbench.latest().result);
 assert.equal(data.schema,"LSCNS_vector_mode_v1");
 assert.equal(data.modes.length,2);
 assert.ok(data.modes.every(m=>Number.isFinite(m.neff)));
 assert.deepEqual(errors,[]);
 console.log("PASS v0.7.1 M2 source and strict Yee M3 entry, no-PML preset, 2-mode Worker, legacy Colab uploader preserved");
} finally {
 await browser.close();
}
```

