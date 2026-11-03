import {estimatePMD,maxwellPDF} from './pmd.mjs';
const $=id=>document.getElementById(id);
const ids=['pmd-coef','pmd-length','pmd-definition'];
function draw(out){
 const canvas=$('pmd-chart'),ctx=canvas.getContext('2d'),W=canvas.width,H=canvas.height,
   L=64,R=22,T=22,B=54,ww=W-L-R,hh=H-T-B;
 ctx.clearRect(0,0,W,H);ctx.fillStyle='#0d192c';ctx.fillRect(0,0,W,H);
 const maxX=Math.max(out.p99DgdPs*1.35,0.001),pts=[];
 for(let i=0;i<=150;i++){const x=maxX*i/150;pts.push([x,maxwellPDF(x,out.maxwellScalePs)]);}
 const maxY=Math.max(...pts.map(v=>v[1]),.001)*1.12;
 const xp=x=>L+x/maxX*ww,yp=y=>H-B-y/maxY*hh;
 ctx.font='13px system-ui';ctx.lineWidth=1;
 for(let j=0;j<=4;j++){const yy=T+j*hh/4;ctx.strokeStyle='#304865';ctx.beginPath();ctx.moveTo(L,yy);ctx.lineTo(W-R,yy);ctx.stroke();
 ctx.fillStyle='#a7c4d8';ctx.textAlign='right';ctx.fillText((maxY*(1-j/4)).toFixed(3),L-8,yy+4);}
 for(let j=0;j<=4;j++){ctx.fillStyle='#a7c4d8';ctx.textAlign='center';ctx.fillText((maxX*j/4).toFixed(3),L+ww*j/4,H-B+23);}
 ctx.beginPath();ctx.strokeStyle='#45d5d4';ctx.lineWidth=3;pts.forEach(([x,y],i)=>i?ctx.lineTo(xp(x),yp(y)):ctx.moveTo(xp(x),yp(y)));ctx.stroke();
 for(const [val,color] of [[out.meanDgdPs,'#f9b876'],[out.p95DgdPs,'#ff88a1']]){
  ctx.strokeStyle=color;ctx.lineWidth=2;ctx.setLineDash([6,4]);ctx.beginPath();ctx.moveTo(xp(val),T);ctx.lineTo(xp(val),H-B);ctx.stroke();ctx.setLineDash([]);
 }
 ctx.fillStyle='#edf7ff';ctx.textAlign='center';ctx.fillText('Differential Group Delay · DGD (ps)',L+ww/2,H-8);
 ctx.save();ctx.translate(17,T+hh/2);ctx.rotate(-Math.PI/2);ctx.fillText('Probability density (1/ps)',0,0);ctx.restore();
 ctx.fillStyle='#f9b876';ctx.fillRect(W-240,16,15,4);ctx.fillStyle='#e5f2ff';ctx.textAlign='left';ctx.fillText('Mean',W-218,21);
 ctx.fillStyle='#ff88a1';ctx.fillRect(W-126,16,15,4);ctx.fillStyle='#e5f2ff';ctx.fillText('P95',W-104,21);
}
function update(){
 try{
  if(ids.slice(0,2).some(id=>$(id).value.trim()===''))throw Error('PMD 입력을 확인하세요.');
  const arg={pmdCoefficientPsSqrtKm:Number($('pmd-coef').value),lengthKm:Number($('pmd-length').value),
   coefficientDefinition:$('pmd-definition').value};
  const out=estimatePMD(arg);
  $('pmd-mean').textContent=out.meanDgdPs.toFixed(5)+' ps';
  $('pmd-rms').textContent=out.rmsDgdPs.toFixed(5)+' ps';
  $('pmd-p95').textContent=out.p95DgdPs.toFixed(5)+' ps';
  $('pmd-p99').textContent=out.p99DgdPs.toFixed(5)+' ps';
  $('pmd-explain').textContent=arg.coefficientDefinition==='mean'?
   'PMD 계수를 평균 DGD/√km로 해석했습니다. 평균 DGD='+out.meanDgdPs.toFixed(5)+' ps.':
   'PMD 계수를 RMS DGD/√km로 해석했습니다. RMS DGD='+out.rmsDgdPs.toFixed(5)+' ps.';
  $('pmd-status').textContent='계산 완료 · 강한 랜덤결합 Maxwell 통계 가정';
  $('pmd-status').className='eng-status success';
  $('pmd-json-cache').value=JSON.stringify(out,null,2);draw(out);
 }catch(err){
  ['pmd-mean','pmd-rms','pmd-p95','pmd-p99'].forEach(id=>$(id).textContent='—');
  $('pmd-status').textContent='입력 오류: '+err.message;$('pmd-status').className='eng-status error';
  $('pmd-json-cache').value='';
 }
}
ids.forEach(id=>$(id).addEventListener('input',update));
$('pmd-run').addEventListener('click',update);
$('pmd-json').addEventListener('click',()=>{
 if(!$('pmd-json-cache').value)return;
 const blob=new Blob([$('pmd-json-cache').value],{type:'application/json'}),uri=URL.createObjectURL(blob),a=document.createElement('a');
 a.href=uri;a.download='pmd_dgd_statistics.json';a.click();setTimeout(()=>URL.revokeObjectURL(uri),1000);
});
update();
