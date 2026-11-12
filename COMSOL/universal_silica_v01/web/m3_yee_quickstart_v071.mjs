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
