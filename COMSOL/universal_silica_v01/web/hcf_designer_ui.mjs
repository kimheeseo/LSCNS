/** Integrated HCF / NANF nested-tube conceptual designer for Universal Silica Lab.
 * Imported analytical functions do not predict loss or solve full-vector FEM. */
import {
  DEFAULT, buildGeometry, wallResonances, wallDetuning, capillaryMetrics,
  spectrum, tubeCountSweep, designSummary, externalComplexNeffLoss
} from './hcf_designer.mjs';

const STYLE=`
#hcf-designer{margin:30px 0;border:1px solid #345470;background:linear-gradient(150deg,#14253c,#102035 56%,#11243a);border-radius:20px;overflow:hidden;box-shadow:0 18px 55px #030a192f;color:#e7f1fc}
#hcf-designer *{box-sizing:border-box}
#hcf-designer .hcf-head{padding:27px 27px 22px;background:radial-gradient(ellipse 60% 120% at 95% 0%,#164a66bb,transparent);border-bottom:1px solid #29415d}
#hcf-designer .hcf-kicker{font-size:11px;color:#60ded5;letter-spacing:1.9px;font-weight:800;margin:0 0 10px}
#hcf-designer h2{font-size:23px;color:#eef9ff;letter-spacing:-.7px;margin:0 0 8px}
#hcf-designer .hcf-desc{font-size:12px;line-height:1.75;color:#acc0d7;margin:0;max-width:1000px}
#hcf-designer .hcf-badges{display:flex;gap:6px;flex-wrap:wrap;margin-top:13px}
#hcf-designer .hcf-tag{font-size:10px;padding:5px 9px;border:1px solid #2a6572;border-radius:20px;color:#9ddedc;background:#17435255;font-weight:700}
#hcf-designer .hcf-body{padding:24px}
#hcf-designer .hcf-top{display:grid;grid-template-columns:minmax(270px,370px) minmax(0,1fr);gap:20px;align-items:start}
#hcf-designer .hcf-card{background:#0c1a2e;border:1px solid #2a435e;border-radius:14px;overflow:hidden;min-width:0}
#hcf-designer .hcf-card-head{padding:15px 18px 11px;border-bottom:1px solid #243c57;display:flex;align-items:center;justify-content:space-between;gap:8px}
#hcf-designer h3{margin:0;color:#e9f7ff;font-size:14px;letter-spacing:-.3px}
#hcf-designer .hcf-sub{font-size:10px;color:#81a3ba}
#hcf-designer .hcf-fields{padding:16px;display:grid;grid-template-columns:1fr 1fr;gap:11px}
#hcf-designer .hcf-field{min-width:0}
#hcf-designer .hcf-field label{display:block;color:#bcd0e3;font-size:11px;margin-bottom:5px;line-height:1.35}
#hcf-designer .hcf-field input,#hcf-designer .hcf-field select{display:block;width:100%;padding:9px;border:1px solid #3c536e;border-radius:7px;background:#172b45;color:#f4fbff;font-size:12px;outline:none}
#hcf-designer .hcf-field input:focus,#hcf-designer .hcf-field select:focus{border-color:#5be2d1;box-shadow:0 0 0 2px #42c7b633}
#hcf-designer .hcf-field-wide{grid-column:1/-1}
#hcf-designer .hcf-check{display:flex;align-items:center;gap:8px;font-size:11px;color:#c3d9e9}
#hcf-designer .hcf-check input{width:auto;accent-color:#40d0c3}
#hcf-designer .hcf-presets{display:flex;gap:6px;flex-wrap:wrap;padding:0 16px 16px}
#hcf-designer button{border:1px solid #3c5d77;background:#1b3851;color:#daf7ff;padding:8px 10px;border-radius:8px;font-size:11px;font-weight:750}
#hcf-designer button:hover{background:#235472;border-color:#49b8d6}
#hcf-designer button.hcf-primary{background:linear-gradient(90deg,#0b899e,#13669c);border-color:#26acb6;color:white}
#hcf-designer .hcf-out{display:grid;gap:15px}
#hcf-designer .hcf-geometry-wrap{padding:12px 13px 2px}
#hcf-designer .hcf-geometry svg{width:100%;height:auto;display:block;max-height:455px}
#hcf-designer .hcf-foot{font-size:10px;color:#93aec8;line-height:1.6;padding:6px 18px 14px;margin:0}
#hcf-designer .hcf-stats{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:9px}
#hcf-designer .hcf-stat{background:#132842;border:1px solid #2b4865;padding:12px;border-radius:10px;min-width:0}
#hcf-designer .hcf-stat span{display:block;color:#8ba7c0;font-size:10px;margin-bottom:7px}
#hcf-designer .hcf-stat strong{display:block;color:#ebffff;font-size:17px;font-variant-numeric:tabular-nums;overflow-wrap:anywhere}
#hcf-designer .hcf-stat em{font-style:normal;color:#87a5bd;font-size:10px}
#hcf-designer .hcf-tables{display:grid;grid-template-columns:1fr 1fr;gap:15px;margin-top:15px}
#hcf-designer .hcf-chart-wrap{padding:12px 12px 7px;min-height:195px}
#hcf-designer .hcf-chart-wrap svg{width:100%;display:block;height:auto;max-height:265px}
#hcf-designer .hcf-tablebox{overflow:auto;max-height:290px;padding:0 12px 12px}
#hcf-designer table{width:100%;border-collapse:collapse;font-size:11px;min-width:350px}
#hcf-designer th{text-align:left;color:#85abc1;font-weight:750;padding:9px 6px;border-bottom:1px solid #325471;white-space:nowrap}
#hcf-designer td{padding:9px 6px;border-bottom:1px solid #24384e;color:#d7e9f6;font-variant-numeric:tabular-nums}
#hcf-designer tr.hcf-selected td{background:#1e4755;color:#dff}
#hcf-designer .hcf-message{margin:14px 0 0;border:1px solid #3c685f;background:#16373d;padding:11px 13px;border-radius:9px;color:#bff4e1;font-size:11px;line-height:1.65}
#hcf-designer .hcf-message.hcf-error{border-color:#995959;background:#472530;color:#ffd4ce}
#hcf-designer .hcf-inline{display:flex;gap:7px;flex-wrap:wrap;padding:10px 15px 15px;align-items:center}
#hcf-designer .hcf-bottom{margin-top:15px;display:grid;gap:15px}
#hcf-designer .hcf-compare{margin:14px 16px 15px;overflow:auto}
#hcf-designer .hcf-external{padding:15px 16px;font-size:12px;display:flex;gap:12px;align-items:end;flex-wrap:wrap}
#hcf-designer .hcf-external .hcf-field{width:240px;max-width:100%}
#hcf-designer .hcf-warning{font-size:11px;color:#c4d5e5;line-height:1.75;padding:14px 16px;border-left:3px solid #c6a15d;background:#262f3a;margin:15px 0 0}
#hcf-designer .hcf-muted{color:#8ea8be;font-size:11px}
@media(max-width:1200px){#hcf-designer .hcf-top{grid-template-columns:1fr}#hcf-designer .hcf-stats{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media(max-width:750px){#hcf-designer .hcf-tables{grid-template-columns:1fr}#hcf-designer .hcf-body{padding:12px}#hcf-designer .hcf-head{padding:19px}#hcf-designer h2{font-size:20px}}
`;
const root=document.createElement('section');root.id='hcf-designer';
root.innerHTML=`
<header class="hcf-head"><p class="hcf-kicker">07 / HCF ADVANCED DESIGN WORKSPACE</p>
<h2>HCF / Nested Tube Designer <span style="color:#55d7c8">v0.1</span></h2>
<p class="hcf-desc">공기 코어 직경, 바깥 튜브 수, 2중·3중 중첩 비율, 실리카 벽 두께를 바꾸면 <strong>실제 입력값대로 단면도</strong>가 바뀝니다. ARROW 공진 파장과 이상적 공기 모세관의 n<sub>eff</sub>·MFD·A<sub>eff</sub>·D를 계산합니다. <strong>Full-vector FEM 및 누설손실 dB/km 예측은 아직 아닙니다.</strong></p>
<div class="hcf-badges"><span class="hcf-tag">LIVE 2D GEOMETRY</span><span class="hcf-tag">ARROW RESONANCE</span><span class="hcf-tag">IDEAL CAPILLARY PROXY</span><span class="hcf-tag">NO LOSS CLAIMS</span></div></header>
<div class="hcf-body">
<div class="hcf-top">
 <div class="hcf-card"><div class="hcf-card-head"><h3>설계 파라미터</h3><span class="hcf-sub">µm · nm</span></div>
  <div class="hcf-fields">
   <div class="hcf-field"><label for="hcf-nesting">중첩 튜브 단계</label><select id="hcf-nesting"><option value="1">1중 (기준)</option><option value="2" selected>2중 (NANF)</option><option value="3">3중 (NANF)</option></select></div>
   <div class="hcf-field"><label for="hcf-tubeCount">클래딩 튜브 수 N</label><input id="hcf-tubeCount" type="number" min="4" max="12" step="1" value="6"/></div>
   <div class="hcf-field"><label for="hcf-coreDiameterUm">Air Core 직경 D<sub>c</sub></label><input id="hcf-coreDiameterUm" type="number" min="16" max="120" step="1" value="32"/></div>
   <div class="hcf-field"><label for="hcf-outerTubeRatio">D₁/D<sub>c</sub> (바깥 튜브)</label><input id="hcf-outerTubeRatio" type="number" min="0.2" max="1.4" step="0.01" value="0.78"/></div>
   <div class="hcf-field"><label for="hcf-innerRatio">D₂/D₁ (중첩 비율)</label><input id="hcf-innerRatio" type="number" min="0.3" max="0.8" step="0.01" value="0.61"/></div>
   <div class="hcf-field"><label for="hcf-thirdRatio">D₃/D₂ (3중)</label><input id="hcf-thirdRatio" type="number" min="0.3" max="0.8" step="0.01" value="0.55"/></div>
   <div class="hcf-field"><label for="hcf-wall1Um">벽 두께 t₁</label><input id="hcf-wall1Um" type="number" min="0.1" max="1.6" step="0.01" value="0.58"/></div>
   <div class="hcf-field"><label for="hcf-wall2Um">벽 두께 t₂</label><input id="hcf-wall2Um" type="number" min="0.1" max="1.6" step="0.01" value="0.45"/></div>
   <div class="hcf-field"><label for="hcf-wall3Um">벽 두께 t₃</label><input id="hcf-wall3Um" type="number" min="0.1" max="1.6" step="0.01" value="0.36"/></div>
   <div class="hcf-field"><label for="hcf-wavelengthNm">분석 파장 λ (nm)</label><input id="hcf-wavelengthNm" type="number" min="300" max="3000" step="10" value="1550"/></div>
   <div class="hcf-field"><label for="hcf-gapCoreUm">Air core–튜브 외벽 간격</label><input id="hcf-gapCoreUm" type="number" min="0" max="6" step="0.05" value="0.25"/></div>
   <div class="hcf-field"><label for="hcf-minNeighborGapUm">인접 튜브 최소 간격</label><input id="hcf-minNeighborGapUm" type="number" min="0" max="6" step="0.05" value="0.35"/></div>
   <div class="hcf-field"><label for="hcf-nestedClearanceUm">Nested 내부 여유</label><input id="hcf-nestedClearanceUm" type="number" min="0" max="1.5" step="0.05" value="0.2"/></div>
   <div class="hcf-field"><label for="hcf-airIndex">공기 굴절률 n<sub>air</sub></label><input id="hcf-airIndex" type="number" min="1" max="1.01" step="0.00001" value="1.00027"/></div>
   <div class="hcf-field-wide hcf-check"><input type="checkbox" id="hcf-autoFit" checked/><label for="hcf-autoFit">튜브 수 증가 시 바깥 직경 자동 맞춤 (충돌 방지)</label></div>
  </div>
  <div class="hcf-presets"><button type="button" data-preset="2">6개 튜브 · 2중</button><button type="button" data-preset="3">8개 튜브 · 3중</button><button type="button" data-preset="1">6개 튜브 · 1중</button></div>
 </div>
 <div class="hcf-out">
  <div class="hcf-card"><div class="hcf-card-head"><h3>실시간 HCF 단면 설계</h3><span class="hcf-sub" id="hcf-geometry-badge">입력 대기</span></div><div class="hcf-geometry hcf-geometry-wrap" id="hcf-svg"></div><p class="hcf-foot" id="hcf-geometry-notes">N개 클래딩 튜브를 공기 코어 주위에 배치합니다.</p></div>
  <div class="hcf-stats">
    <div class="hcf-stat"><span>공기 코어 n<sub>eff</sub> (모세관 근사)</span><strong id="hcf-neff">—</strong><em>튜브 효과 미반영</em></div>
    <div class="hcf-stat"><span>Petermann-II MFD</span><strong id="hcf-mfd">—</strong><em>µm · 이상적 J₀</em></div>
    <div class="hcf-stat"><span>유효면적 A<sub>eff</sub></span><strong id="hcf-aeff">—</strong><em>µm² · 이상적 J₀</em></div>
    <div class="hcf-stat"><span>Waveguide D</span><strong id="hcf-d">—</strong><em>ps/(nm·km) · n_air 일정</em></div>
    <div class="hcf-stat"><span>유리 튜브 각도 커버율</span><strong id="hcf-coverage">—</strong><em>기하학적 지표 · 손실 아님</em></div>
    <div class="hcf-stat"><span>튜브 사이 최소 간격</span><strong id="hcf-gap">—</strong><em>µm · 0 이상 필요</em></div>
    <div class="hcf-stat"><span>첫 튜브 1차 공진 λ</span><strong id="hcf-res">—</strong><em>nm · ARROW 근사</em></div>
    <div class="hcf-stat"><span>Confinement Loss</span><strong id="hcf-loss">미계산</strong><em>복소 벡터 FEM/PML 필요</em></div>
  </div>
 </div>
</div>
<div id="hcf-message" class="hcf-message" role="status" aria-live="polite">설계값을 변경하면 단면과 근사 계산이 즉시 갱신됩니다.</div>
<div class="hcf-tables">
 <div class="hcf-card"><div class="hcf-card-head"><h3>튜브 벽별 공진 · 반공진 파장</h3><span class="hcf-sub">ARROW · Sellmeier</span></div><div class="hcf-tablebox" id="hcf-res-table"></div></div>
 <div class="hcf-card"><div class="hcf-card-head"><h3>튜브 개수 N 변화</h3><span class="hcf-sub">기하학적 커버율 (%)</span></div><div class="hcf-chart-wrap" id="hcf-count-chart"></div><p class="hcf-foot">N=4–12 스윕. 자동 맞춤 시 실제 D₁이 달라집니다. 커버율은 누설손실/효율이 아닙니다.</p></div>
 <div class="hcf-card"><div class="hcf-card-head"><h3>공진 회피 지표 vs 파장</h3><span class="hcf-sub">0=공진 · 1=반공진 (벽별)</span></div><div class="hcf-chart-wrap" id="hcf-phase-chart"></div><p class="hcf-foot">얇은 벽 ARROW 위상 근사값. 그래프는 전송률이나 dB/km 손실이 아닙니다.</p></div>
 <div class="hcf-card"><div class="hcf-card-head"><h3>이상적 모세관 분산 D(λ)</h3><span class="hcf-sub">튜브 수·중첩 효과 제외</span></div><div class="hcf-chart-wrap" id="hcf-d-chart"></div><p class="hcf-foot">공기 굴절률을 파장에 무관한 상수로 가정했습니다. 실제 HCF 분산은 벽 공진 근처에서 크게 달라집니다.</p></div>
</div>
<div class="hcf-bottom">
 <div class="hcf-card"><div class="hcf-card-head"><h3>설계 A/B 비교 및 결과 저장</h3><span class="hcf-sub">표와 데이터 내려받기</span></div><div class="hcf-inline"><button type="button" class="hcf-primary" id="hcf-save-a">현재 설계 A 저장</button><button type="button" id="hcf-compare">A와 현재 B 비교</button><button type="button" id="hcf-export-json">설계+수치 JSON</button><button type="button" id="hcf-export-csv">스펙트럼 CSV</button><button type="button" id="hcf-export-svg">단면 SVG</button></div><div id="hcf-compare-table" class="hcf-compare hcf-muted">설계 A를 저장한 뒤 코어·튜브 수 등을 바꾸어 B와 비교할 수 있습니다.</div></div>
 <div class="hcf-card"><div class="hcf-card-head"><h3>외부 복소 FEM 결과 → Confinement Loss 후처리</h3><span class="hcf-sub">사용자 입력값만 환산</span></div><div class="hcf-external"><div class="hcf-field"><label for="hcf-external-im">외부 계산 |Im(n<sub>eff</sub>)|</label><input id="hcf-external-im" type="number" min="0" max="0.001" step="1e-9" placeholder="예: 1e-10"/></div><button type="button" id="hcf-external-calc">외부값으로 dB/km 계산</button><span id="hcf-external-answer" class="hcf-muted">외부 FEM 결과를 입력하면 환산값을 표시합니다.</span></div><p class="hcf-foot">PML과 메시 수렴을 거친 복소 고유모드의 부호를 확인하고 절댓값을 입력하십시오. 계산된 dB/km는 **현재 웹 툴이 직접 예측한 값이 아닙니다**. PML 미수렴 결과는 제품 성능값으로 사용할 수 없습니다.</p></div>
</div>
<div class="hcf-warning"><strong>물리 모델의 명시적 한계:</strong> 실리카 벽 공진식은 각 벽의 국소 간섭 조건만 근사합니다. N·D₂/D₁·D₃/D₂는 기하학적 공극/배치에는 영향을 주지만, 현재 모세관 모드 해석식의 n<sub>eff</sub>·MFD·A<sub>eff</sub>·D에는 반영되지 않습니다. 실제 중첩 튜브 수 변화에 따른 손실 dB/km, 표면 산란, 반공진 모드 결합, 굽힘손실은 복소 **2D Full-vector Maxwell FEM + PML** 및 독립 검증이 필요합니다. 따라서 이 도구는 HCF 사전 설계기이며 완성된 손실 시뮬레이터가 아닙니다.</div>
</div>`;
const style=document.createElement('style');style.id='hcf-style-v01';style.textContent=STYLE;document.head.append(style);
const anchor=document.getElementById('pmd-tool');
if(anchor)anchor.before(root);else document.querySelector('main')?.append(root);
const $=id=>root.querySelector('#'+id);
const keys=['coreDiameterUm','tubeCount','outerTubeRatio','innerRatio','thirdRatio','gapCoreUm','minNeighborGapUm','nestedClearanceUm','wall1Um','wall2Um','wall3Um','wavelengthNm','airIndex'];
let snapshotA=null,current=null;
function getInputs(){
 const p={...DEFAULT};
 p.nesting=Number($('hcf-nesting').value);
 for(const key of keys)p[key]=Number($('hcf-'+key).value);
 p.autoFit=$('hcf-autoFit').checked;
 return p;
}
function setInputs(p){
 $('hcf-nesting').value=String(p.nesting);
 for(const k of keys)$('hcf-'+k).value=String(p[k]);
 $('hcf-autoFit').checked=!!p.autoFit;
 refresh();
}
function f(num,d=2){return Number.isFinite(num)?Number(num).toFixed(d):'—';}
const esc=s=>String(s).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;');
function ringSvg(l,x,y,r,wall,scale){
 const palette=['#72e9ea','#78a6f7','#d39aed'];const color=palette[l-1];
 const cx=x.toFixed(3),cy=y.toFixed(3),rr=(r*scale).toFixed(3),ri=(Math.max(0,r-wall)*scale).toFixed(3);
 return `<g><circle cx="${cx}" cy="${cy}" r="${rr}" fill="${color}" fill-opacity="0.88" stroke="${color}" stroke-width="1.6"/><circle cx="${cx}" cy="${cy}" r="${ri}" fill="#14273a" stroke="#1d4964" stroke-width="1"/></g>`;
}
function geometrySvg(g){
 const W=800,H=510,cx=W/2,cy=H/2;const sc=215/(g.extentUm+2.0);
 const px=x=>cx+x*sc,py=y=>cy+y*sc;
 const field=g.input;
 const circles=[];
 for(const tube of g.rings){
  for(const layer of tube.layers){circles.push(ringSvg(layer.level,px(layer.x),py(layer.y),layer.radius,layer.wall,sc));}
 }
 const grid=[];for(let u=-4;u<=4;u++){grid.push(`<line x1="${cx+u*56}" y1="0" x2="${cx+u*56}" y2="${H}" stroke="#1b3650" opacity="0.38"/>`);grid.push(`<line y1="${cy+u*56}" x1="0" y2="${cy+u*56}" x2="${W}" stroke="#1b3650" opacity="0.38"/>`);}
 return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" role="img" aria-label="실시간 생성된 ${field.tubeCount}개 튜브 ${field.nesting}중 중첩 공기 코어 광섬유 단면"><defs><radialGradient id="hcf-core-grad"><stop stop-color="#113f61"/><stop offset="1" stop-color="#0d2840"/></radialGradient></defs><rect width="${W}" height="${H}" rx="14" fill="#0b1b30"/>${grid.join('')}<circle cx="${cx}" cy="${cy}" r="${(g.coreRadiusUm*sc).toFixed(2)}" fill="url(#hcf-core-grad)" stroke="#53c9e1" stroke-dasharray="6 5" stroke-width="2"/>${circles.join('')}<circle cx="${cx}" cy="${cy}" r="${Math.max(5,g.coreRadiusUm*sc*0.28).toFixed(2)}" fill="#0a2b41" opacity="0.6"/><text x="${cx}" y="${cy-5}" text-anchor="middle" fill="#ddfbfe" font-size="16" font-weight="750">AIR CORE</text><text x="${cx}" y="${cy+17}" text-anchor="middle" fill="#98d9e9" font-size="13">Ø${f(field.coreDiameterUm,1)} µm</text><text x="22" y="31" fill="#ddf7ff" font-size="13" font-weight="700">${field.tubeCount}-TUBE · ${field.nesting}-LAYER</text><text x="22" y="52" fill="#9ebbd1" font-size="11">D₁ = ${f(g.outerTubeDiameterUm)} µm ${g.autoFitted?'[AUTO-FIT]':''}</text><g transform="translate(20 ${H-25})"><circle r="5" fill="#72e9ea"/><text x="13" y="4" fill="#d3e9fc" font-size="11">Outer glass</text><circle cx="132" r="5" fill="#78a6f7"/><text x="145" y="4" fill="#d3e9fc" font-size="11">Nested 2</text><circle cx="264" r="5" fill="#d39aed"/><text x="278" y="4" fill="#d3e9fc" font-size="11">Nested 3</text></g><line x1="550" y1="${H-32}" x2="700" y2="${H-32}" stroke="#e4f5fb" stroke-width="2"/><line x1="550" y1="${H-38}" x2="550" y2="${H-26}" stroke="#e4f5fb"/><line x1="700" y1="${H-38}" x2="700" y2="${H-26}" stroke="#e4f5fb"/><text x="625" y="${H-43}" text-anchor="middle" fill="#bfd7e9" font-size="11">${f(150/sc,1)} µm</text></svg>`;
}
function plotSvg(series,opts){
 const W=660,H=250,L=60,R=15,T=15,B=42;
 const flat=series.flatMap(s=>s.data.map(q=>q.y)).filter(Number.isFinite);
 let ymin=opts.ymin??Math.min(...flat),ymax=opts.ymax??Math.max(...flat);
 if(!Number.isFinite(ymin)||!Number.isFinite(ymax)){ymin=0;ymax=1;}
 if(ymax<=ymin){ymax=ymin+1;}
 if(opts.pad!==false){const extra=(ymax-ymin)*0.09;ymin-=extra;ymax+=extra;}
 const xmin=opts.xmin,xmax=opts.xmax;const X=x=>L+(x-xmin)/(xmax-xmin)*(W-L-R);const Y=y=>T+(ymax-y)/(ymax-ymin)*(H-T-B);
 const ax=[];for(let i=0;i<=4;i++){const value=ymin+(ymax-ymin)*i/4,yp=Y(value);ax.push(`<line x1="${L}" x2="${W-R}" y1="${yp.toFixed(2)}" y2="${yp.toFixed(2)}" stroke="#23445e"/><text x="${L-9}" y="${(yp+4).toFixed(2)}" text-anchor="end" font-size="11" fill="#9bb8cd">${opts.tick?opts.tick(value):f(value,2)}</text>`);}
 for(let i=0;i<=4;i++){const v=xmin+(xmax-xmin)*i/4,xp=X(v);ax.push(`<text x="${xp.toFixed(2)}" y="${H-22}" text-anchor="middle" font-size="11" fill="#9bb8cd">${opts.xTick?opts.xTick(v):f(v,0)}</text>`);}
 const paths=series.map(s=>`<path d="${s.data.filter(v=>Number.isFinite(v.y)).map((v,i)=>(i?'L':'M')+f(X(v.x),2)+' '+f(Y(v.y),2)).join(' ')}" fill="none" stroke="${s.color}" stroke-width="2.5" stroke-linejoin="round"/>`).join('');
 let cursor='';if(Number.isFinite(opts.cursor)&&opts.cursor>=xmin&&opts.cursor<=xmax){const x=X(opts.cursor);cursor=`<line x1="${x}" x2="${x}" y1="${T}" y2="${H-B}" stroke="#f4d28d" stroke-width="2" stroke-dasharray="5 5"/>`}
 const legend=series.map((s,i)=>`<g transform="translate(${L+10+i*155},${H-6})"><rect width="12" height="3" y="-7" fill="${s.color}"/><text x="18" y="-2" fill="#b9d2e3" font-size="11">${esc(s.name)}</text></g>`).join('');
 return `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(opts.title)}"><rect width="${W}" height="${H}" fill="#0c1a2e"/>${ax.join('')}${paths}${cursor}${legend}</svg>`;
}
function coverageSvg(rows,selected){
 const W=660,H=250,L=50,R=12,T=12,B=41;
 const bw=(W-L-R)/rows.length;
 const bars=rows.map((x,i)=>{
  const v=x.valid?x.coverageFraction:0;
  const h=v*(H-T-B),xx=L+i*bw+3,y=H-B-h;
  return `<g><rect x="${f(xx,2)}" y="${f(y,2)}" width="${f(bw-7,2)}" height="${f(h,2)}" rx="4" fill="${x.tubeCount===selected?'#58e6da':x.valid?'#326a98':'#784655'}"/><text x="${f(xx+(bw-7)/2,1)}" y="${H-19}" text-anchor="middle" fill="${x.tubeCount===selected?'#baffee':'#a5c1d4'}" font-size="12">${x.tubeCount}</text><text x="${f(xx+(bw-7)/2,1)}" y="${f(Math.max(T+12,y-6),1)}" text-anchor="middle" fill="#b0d6e5" font-size="10">${x.valid?f(v*100,0)+'%':'×'}</text></g>`;
 }).join('');
 const ticks=[0,.25,.5,.75,1].map(v=>{const y=H-B-v*(H-T-B);return `<line x1="${L}" y1="${y}" x2="${W-R}" y2="${y}" stroke="#26465f"/><text x="${L-8}" y="${y+4}" text-anchor="end" fill="#97b7ce" font-size="11">${Math.round(v*100)}%</text>`}).join('');
 return `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="N=4부터 12까지 튜브 각도 커버율 비교"><rect width="${W}" height="${H}" fill="#0c1a2e"/>${ticks}${bars}<text x="${W-22}" y="${H-6}" text-anchor="end" fill="#91aaca" font-size="10">튜브 수 N</text></svg>`;
}
function resultsTable(res,p){
 const useful=res.filter(x=>x.order===1 || (x.kind==='resonance' && x.order===2));
 return `<table><thead><tr><th>벽</th><th>상태</th><th>차수</th><th>λ (nm)</th><th>f (THz)</th></tr></thead><tbody>${useful.map(x=>`<tr><td>t${x.layer} (${f(x.wallUm,2)} µm)</td><td>${x.kind==='resonance'?'공진':'반공진'}</td><td>${x.order}</td><td>${f(x.wavelengthNm,1)}</td><td>${f(x.frequencyTHz,1)}</td></tr>`).join('')}</tbody></table>`;
}
function updateMessage(message,error=false){const x=$('hcf-message');x.textContent=message;x.classList.toggle('hcf-error',error);}
function emptyOnError(){
 $('hcf-geometry-badge').textContent='입력 오류';$('hcf-svg').innerHTML='';
 for(const key of ['neff','mfd','aeff','d','coverage','gap','res'])$('hcf-'+key).textContent='—';
 $('hcf-loss').textContent='미계산';
 for(const key of ['res-table','count-chart','phase-chart','d-chart'])$('hcf-'+key).innerHTML='';
}
function refresh(){
 try{
  const p=getInputs();
  const s=designSummary(p),g=buildGeometry(p),m=s.mode;
  const spec=spectrum(p,97),counts=tubeCountSweep(p);
  current={p,s,g,spec,counts};
  $('hcf-svg').innerHTML=geometrySvg(g);
  $('hcf-geometry-badge').textContent=`N=${p.tubeCount} · ${p.nesting}중 구조`;
  $('hcf-geometry-notes').textContent=`실제 D₁=${f(g.outerTubeDiameterUm,2)} µm · 인접 튜브 간격=${f(g.minAdjacentGapUm,2)} µm · 내부 Nested 튜브는 외곽 방향으로 편심 배치한 이상적 비접촉 구조입니다.`;
  $('hcf-neff').textContent=f(m.neff,8);
  $('hcf-mfd').textContent=f(m.MFD_PII_um,2);
  $('hcf-aeff').textContent=f(m.Aeff_um2,1);
  $('hcf-d').textContent=f(m.D_ps_per_nm_km,3);
  $('hcf-coverage').textContent=f(g.angularCoverageFraction*100,1)+'%';
  $('hcf-gap').textContent=f(g.minAdjacentGapUm,2);
  const primary=s.resonances.find(x=>x.layer===1&&x.order===1&&x.kind==='resonance');
  $('hcf-res').textContent=primary?f(primary.wavelengthNm,0)+' nm':'범위 밖';
  $('hcf-loss').textContent='미계산';
  $('hcf-res-table').innerHTML=resultsTable(s.resonances,p);
  $('hcf-count-chart').innerHTML=coverageSvg(counts,p.tubeCount);
  const colors=['#60dfdb','#81a9ff','#d59dfb'];
  const layers=Array.from({length:p.nesting},(_,i)=>({name:'Tube t'+(i+1),color:colors[i],data:spec.map(x=>({x:x.lambdaNm,y:x.layerDetunings[i]}))}));
  $('hcf-phase-chart').innerHTML=plotSvg(layers,{title:'벽별 ARROW 공진 회피 지표',xmin:p.scanMinNm,xmax:p.scanMaxNm,ymin:0,ymax:1,pad:false,cursor:p.wavelengthNm,tick:x=>f(x,2)});
  $('hcf-d-chart').innerHTML=plotSvg([{name:'D ideal',color:'#5fdce3',data:spec.map(x=>({x:x.lambdaNm,y:x.D}))}],{title:'이상적 모세관 D(λ)',xmin:p.scanMinNm,xmax:p.scanMaxNm,cursor:p.wavelengthNm});
  updateMessage(`${p.nesting}중 · ${p.tubeCount}개 튜브 · 분석 ${f(p.wavelengthNm,0)} nm | ${g.autoFitted?'충돌 방지를 위해 D₁을 요청값 '+f(g.requestedOuterTubeDiameterUm,2)+'→'+f(g.outerTubeDiameterUm,2)+' µm로 자동 축소했습니다. ':'입력 치수가 유지되었습니다. '}N/튜브 비율은 단면·기하 지표를 바꾸지만, 이번 이상적 모세관 n_eff/Aeff/D 예측에는 반영되지 않습니다.`);
  // Input changes invalidate previously imported external loss and B comparison.
  $('hcf-external-answer').textContent='현재 입력에 대응하는 외부 Im(neff)를 넣으면 환산합니다.';
  if(snapshotA)$('hcf-compare-table').textContent='설계 A 저장됨. [A와 현재 B 비교] 버튼으로 새 결과를 확인하세요.';
 }catch(e){current=null;emptyOnError();updateMessage('입력 설계 오류: '+e.message,true);}
}
function saveBlob(blob,filename){const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=filename;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),500);}
function exportJson(){if(!current)return;const payload={...current.s,geometry:{...current.g,rings:current.g.rings},countSweep:current.counts,spectrum:current.spec};saveBlob(new Blob([JSON.stringify(payload,null,2)],{type:'application/json'}),'HCF_NestedTube_Design.json');}
function exportCsv(){if(!current)return;const lines=['lambda_nm,capillary_neff,capillary_D_ps_nm_km,wall1_detuning,wall2_detuning,wall3_detuning,confinement_loss_db_km'];for(const x of current.spec){lines.push([x.lambdaNm,x.neff,x.D,...[0,1,2].map(i=>x.layerDetunings[i]??''),''].join(','));}saveBlob(new Blob([lines.join('\n')],{type:'text/csv;charset=utf-8'}),'HCF_NestedTube_Spectrum.csv');}
function exportSvg(){if(!current)return;const svg=$('hcf-svg').querySelector('svg');if(svg)saveBlob(new Blob([svg.outerHTML],{type:'image/svg+xml'}),'HCF_NestedTube_CrossSection.svg');}
function showCompare(){if(!current||!snapshotA)return;const A=snapshotA,B=current;
 const rows=[['튜브 개수 N',A.p.tubeCount,B.p.tubeCount],['Nested 단계',A.p.nesting,B.p.nesting],['Core 직경 µm',f(A.p.coreDiameterUm,2),f(B.p.coreDiameterUm,2)],['실제 D₁ µm',f(A.g.outerTubeDiameterUm,3),f(B.g.outerTubeDiameterUm,3)],['커버율 %',f(A.g.angularCoverageFraction*100,2),f(B.g.angularCoverageFraction*100,2)],['인접 Gap µm',f(A.g.minAdjacentGapUm,3),f(B.g.minAdjacentGapUm,3)],['공진 λ₁ (nm)',f(A.s.resonances.find(x=>x.layer===1&&x.kind==='resonance'&&x.order===1)?.wavelengthNm,1),f(B.s.resonances.find(x=>x.layer===1&&x.kind==='resonance'&&x.order===1)?.wavelengthNm,1)],['n_eff (capillary)',f(A.s.mode.neff,8),f(B.s.mode.neff,8)],['MFD_PII µm',f(A.s.mode.MFD_PII_um,3),f(B.s.mode.MFD_PII_um,3)],['Aeff µm²',f(A.s.mode.Aeff_um2,3),f(B.s.mode.Aeff_um2,3)],['D ps/nm/km',f(A.s.mode.D_ps_per_nm_km,4),f(B.s.mode.D_ps_per_nm_km,4)],['Loss dB/km','미계산','미계산']];
 $('hcf-compare-table').innerHTML=`<table><thead><tr><th>비교 항목</th><th>저장 설계 A</th><th>현재 설계 B</th></tr></thead><tbody>${rows.map(x=>`<tr><td>${x[0]}</td><td>${x[1]}</td><td>${x[2]}</td></tr>`).join('')}</tbody></table><p class="hcf-muted">튜브 개수 차이는 모세관 근사 n_eff·D에 반영되지 않습니다. 손실 비교는 복소 FEM이 필요한 별도 과제입니다.</p>`;
}
// Each input directly and synchronously controls the visualization. No data is uploaded.
root.addEventListener('input',event=>{
 if(event.target.id==='hcf-external-im'){
  $('hcf-external-answer').textContent='외부값을 계산 버튼으로 환산하십시오.';return;
 }
 if(event.target.matches('input,select'))refresh();
});
root.addEventListener('change',event=>{if(event.target.matches('select,input'))refresh();});
root.addEventListener('click',event=>{
 const preset=event.target.closest('[data-preset]');if(preset){
  const depth=Number(preset.dataset.preset);
  const cfg=depth===3?{...DEFAULT,tubeCount:8,nesting:3,outerTubeRatio:.76,thirdRatio:.55}:depth===2?{...DEFAULT,tubeCount:6,nesting:2}:{...DEFAULT,tubeCount:6,nesting:1};
  setInputs(cfg);return;
 }
 const target=event.target.closest('button');if(!target)return;
 switch(target.id){
  case 'hcf-save-a':if(current){snapshotA=structuredClone(current);$('hcf-compare-table').textContent=`A 저장: ${snapshotA.p.tubeCount}개 / ${snapshotA.p.nesting}중. 현재 설계를 바꿔보세요.`;}break;
  case 'hcf-compare':showCompare();break;
  case 'hcf-export-json':exportJson();break;
  case 'hcf-export-csv':exportCsv();break;
  case 'hcf-export-svg':exportSvg();break;
  case 'hcf-external-calc':{
   try{if(!current)throw Error('유효한 설계가 없습니다.');const raw=$('hcf-external-im').value;if(raw==='')throw Error('Im(neff) 입력이 필요합니다.');const im=Number(raw);
   const loss=externalComplexNeffLoss(im,current.p.wavelengthNm);
   $('hcf-external-answer').textContent=`외부 복소 neff 환산: ${f(loss,6)} dB/km (이 웹의 예측 아님)`;
   }catch(e){$('hcf-external-answer').textContent='계산 오류: '+e.message;}
  }break;
 }
});
refresh();
