import {bentIndexCut,bendLossFromComplexNeff,assessPMLConvergence,VELAMURI_REFERENCE} from './bend_physics_bridge.mjs?v=20261103-velamuri-v1';
const $=id=>document.getElementById(id);
let lastResult=null;
const keys=['bendphys-wave','bendphys-diameter','bendphys-neffimag','bendphys-turns','bendphys-clad','bendphys-poisson','bendphys-stress'];
function validNumber(id){const el=$(id);if(!el||el.value.trim()==='')throw Error(id+' 입력값이 비어 있습니다.');let v=Number(el.value);if(!Number.isFinite(v))throw Error(id+' 유효한 숫자가 아닙니다.');return v;}
function currentModel(){
 const profile=$('profile').value;
 if(profile==='mcf')throw Error('이 2D 단면 진단에서는 MCF 보간 예제를 지원하지 않습니다.');
 let shells=[];if(profile==='multilayer'){
   shells=$('shells').value.trim().split(/\n+/).filter(Boolean).map(x=>x.split(',').map(Number));
   if(!shells.length || shells.some(x=>x.length!==2||x.some(z=>!Number.isFinite(z))))
     throw Error('다층 프로파일 입력을 확인하세요.');
 }
 return {profile,core_radius_um:validNumber('core_radius_um'),
   delta_n_core:validNumber('delta_n_core'),graded_alpha:validNumber('graded_alpha'),
   trench_start_um:validNumber('trench_start_um'),trench_end_um:validNumber('trench_end_um'),
   trench_delta_n:validNumber('trench_delta_n'),shells};
}
function drawCut(rows){
 const c=$('bendphys-profile-chart'),ctx=c.getContext('2d'),W=c.width,H=c.height;
 const P={left:76,right:24,top:37,bottom:56},w=W-P.left-P.right,h=H-P.top-P.bottom;
 ctx.clearRect(0,0,W,H);ctx.fillStyle='#0d192c';ctx.fillRect(0,0,W,H);
 const ys=rows.flatMap(r=>[r.nStraight,r.nGeometryOnly,r.nEquivalent]);
 const x0=rows[0].xUm,x1=rows.at(-1).xUm,min=Math.min(...ys),max=Math.max(...ys),pad=Math.max(1e-4,(max-min)*.09);
 const minY=min-pad,maxY=max+pad;
 const xp=x=>P.left+(x-x0)/(x1-x0)*w,yp=y=>P.top+(maxY-y)/(maxY-minY)*h;
 ctx.font='13px sans-serif';
 for(let i=0;i<=4;i++){
  const y=P.top+h*i/4,xx=P.left+w*i/4;
  ctx.strokeStyle='#304963';ctx.lineWidth=1;ctx.beginPath();ctx.moveTo(P.left,y);ctx.lineTo(W-P.right,y);ctx.stroke();
  ctx.fillStyle='#a6c1d8';ctx.textAlign='right';ctx.fillText((maxY-(maxY-minY)*i/4).toFixed(4),P.left-8,y+4);
  ctx.textAlign='center';ctx.fillText((x0+(x1-x0)*i/4).toFixed(1),xx,H-P.bottom+21);
 }
 ctx.strokeStyle='#496984';ctx.setLineDash([3,4]);ctx.beginPath();ctx.moveTo(xp(0),P.top);ctx.lineTo(xp(0),H-P.bottom);ctx.stroke();ctx.setLineDash([]);
 const draw=(key,color)=>{ctx.beginPath();ctx.strokeStyle=color;ctx.lineWidth=2.5;
  rows.forEach((v,i)=>i?ctx.lineTo(xp(v.xUm),yp(v[key])):ctx.moveTo(xp(v.xUm),yp(v[key])));ctx.stroke();};
 draw('nStraight','#6caeff');draw('nGeometryOnly','#f7be73');draw('nEquivalent','#55d9c9');
 ctx.fillStyle='#e8f4ff';ctx.textAlign='center';ctx.fillText('Cross-section x (µm) · positive toward outside of bend',P.left+w/2,H-8);
 ctx.save();ctx.translate(17,P.top+h/2);ctx.rotate(-Math.PI/2);ctx.fillText('n(x) and n_equivalent(x)',0,0);ctx.restore();
 let x=P.left+5;ctx.textAlign='left';
 for(const [name,color] of [['Straight','#6caeff'],['Geometry only','#f7be73'],['Geo + stress proxy','#55d9c9']]){
  ctx.fillStyle=color;ctx.fillRect(x,17,18,4);ctx.fillStyle='#e4f2ff';ctx.fillText(name,x+24,23);x+=name.length*8+74;
 }
}
function parsePML(text){
 return text.trim().split(/\n+/).filter(Boolean).map((line,i)=>{
  const cells=line.split(/[\s,;]+/).filter(Boolean).map(Number);
  if(cells.length!==2||cells.some(x=>!Number.isFinite(x)))throw Error('PML '+(i+1)+'행은 [두께(λ), Im(neff)] 형식입니다.');
  return {thicknessLambda:cells[0],neffImagMagnitude:cells[1]};
 });
}
function update(){
 try{
  const wavelengthNm=validNumber('bendphys-wave'),bendDiameterMm=validNumber('bendphys-diameter'),
    neffImagMagnitude=validNumber('bendphys-neffimag'),turns=validNumber('bendphys-turns'),
    claddingRadiusUm=validNumber('bendphys-clad'),poissonRatio=validNumber('bendphys-poisson'),
    stressMode=$('bendphys-stress').value,model=currentModel();
  const result=bendLossFromComplexNeff({wavelengthNm,bendDiameterMm,neffImagMagnitude,turns});
  const curve=bentIndexCut(model,{wavelengthNm,bendDiameterMm,claddingRadiusUm,stressMode,poissonRatio,samples:301});
  drawCut(curve.rows);
  $('bendphys-dbturn').textContent=result.lossDbPerTurn.toFixed(7)+' dB/turn';
  $('bendphys-dbperkm').textContent=(result.lossDbPerM*1000).toFixed(4)+' dB/km';
  $('bendphys-total').textContent=result.totalBendLossDb.toFixed(6)+' dB';
  $('bendphys-pml').textContent=(7*wavelengthNm/1000).toFixed(3)+' µm';
  $('bendphys-status').textContent='후처리 계산 완료 · 복소 n_eff는 입력값, 현재 FEM이 계산한 값이 아님';
  $('bendphys-status').className='eng-status success';
  const inputPML=$('bendphys-pml-data').value.trim();let pml=null;
  if(inputPML){
   pml=assessPMLConvergence({wavelengthNm,bendDiameterMm,rows:parsePML(inputPML),
     tolerancePercent:validNumber('bendphys-pml-tolerance')});
   $('bendphys-pmlstatus').textContent=
    '외부 FEM PML 두께 증가에 따른 손실 상대 변화: '+pml.relativeChangePercent.toFixed(3)+'% · '
    +(pml.passed?'기준 충족':'수렴 미확인')+' · 표본 '+pml.rows.length+'개';
  }else{
   $('bendphys-pmlstatus').textContent='외부 FEM PML 스윕 미입력 · PML 수렴 검증 미실시';
  }
  lastResult={external_complex_neff_postprocess:result,index_visualization:{
    wavelengthNm,bendDiameterMm,claddingRadiusUm,stressMode,poissonRatio,core_model:model,
    representation:'1D cut of equivalent-straight-guide index, not a 2D mode solver'},
    PML_convergence:pml,paper_reference:VELAMURI_REFERENCE,
    caveats:['Im(neff) supplied by user/another solver, not predicted here',
      'GEBT and complex-mode FEM/PML not implemented','Any example imaginary neff is illustrative or inverted from reference data']};
 }catch(error){
  $('bendphys-status').textContent='확인 필요: '+error.message;$('bendphys-status').className='eng-status error';
  ['bendphys-dbturn','bendphys-dbperkm','bendphys-total','bendphys-pml'].forEach(id=>$(id).textContent='—');
  $('bendphys-pmlstatus').textContent='검증 데이터 없음';
  lastResult=null;
 }
}
keys.forEach(id=>$(id).addEventListener('input',update));
$('bendphys-pml-data').addEventListener('input',update);
$('bendphys-pml-tolerance').addEventListener('input',update);
$('bendphys-run').addEventListener('click',update);
$('bendphys-from-radius').addEventListener('click',()=>{
 const radius=Number($('bend-radius').value),wave=Number($('bend-wave').value);
 if(Number.isFinite(radius)&&radius>0)$('bendphys-diameter').value=(2*radius).toString();
 if(Number.isFinite(wave)&&wave>0)$('bendphys-wave').value=wave.toString();
 update();
});
$('bendphys-example').addEventListener('click',()=>{
 $('bendphys-diameter').value=String(VELAMURI_REFERENCE.exampleBendDiameterMm);
 $('bendphys-wave').value=String(VELAMURI_REFERENCE.exampleWavelengthNm);
 $('bendphys-neffimag').value=String(
  VELAMURI_REFERENCE.reportedBifSimLossDbTurn/
  ((20/Math.log(10))*(2*Math.PI/(1550e-9))*(Math.PI*10/1000)));
 $('bendphys-turns').value='1';$('bendphys-stress').value='none';
 $('bendphys-pml-data').value='';
 update();
 $('bendphys-pmlstatus').textContent='논문 예제의 손실 0.0119 dB/turn에서 Im(neff)를 역산한 자기일관성 검사입니다. FEM의 독립 예측/검증이 아닙니다.';
});
$('bendphys-json').addEventListener('click',()=>{
 if(!lastResult)return;
 const blob=new Blob([JSON.stringify(lastResult,null,2)],{type:'application/json'}),uri=URL.createObjectURL(blob);
 const a=document.createElement('a');a.href=uri;a.download='velamuri_bend_physics_postprocessor.json';a.click();
 setTimeout(()=>URL.revokeObjectURL(uri),1000);
});
update();
