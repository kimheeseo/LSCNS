import {estimateAttenuation} from './attenuation.mjs';
const el=id=>document.getElementById(id);
const ids=['att-alpha1550','att-alpha1625','att-wave','att-length','att-pin','att-connectors','att-connloss','att-splices','att-spliceloss'];
const keys=['alpha1550','alpha1625','wavelengthNm','lengthKm','launchDbm','connectorCount','connectorDb','spliceCount','spliceDb'];
function calculate(){
 const status=el('att-status');
 try{
  const obj=Object.fromEntries(ids.map((id,i)=>[keys[i],Number(el(id).value)]));
  if(ids.some(id=>el(id).value.trim()===''))throw Error('모든 입력값을 기재하세요.');
  const r=estimateAttenuation(obj);
  el('att-alpha-out').textContent=r.alphaDbKm.toFixed(5)+' dB/km';
  el('att-fiber-out').textContent=r.fiberLossDb.toFixed(3)+' dB';
  el('att-total-out').textContent=r.totalLossDb.toFixed(3)+' dB';
  el('att-pout-out').textContent=r.outputDbm.toFixed(3)+' dBm';
  el('att-breakdown').textContent=
    '섬유 '+r.fiberLossDb.toFixed(3)+' dB + 커넥터 '+r.connectorLossDb.toFixed(3)+
    ' dB + 융착 '+r.spliceLossDb.toFixed(3)+' dB = 총 '+r.totalLossDb.toFixed(3)+' dB';
  status.textContent='계산 완료 · 입력값 기반 보간 (실측 검증 아님)';
  status.className='eng-status success';
  const canvas=el('att-chart'),ctx=canvas.getContext('2d');
  const W=canvas.width,H=canvas.height,left=65,right=28,top=28,bottom=54,usableX=W-left-right,usableY=H-top-bottom;
  ctx.clearRect(0,0,W,H);ctx.fillStyle='#0d192c';ctx.fillRect(0,0,W,H);
  const maxX=Math.max(r.input.lengthKm,1),base=r.connectorLossDb+r.spliceLossDb,maxY=Math.max(r.totalLossDb,1);
  ctx.font='13px system-ui';ctx.lineWidth=1;
  for(let j=0;j<=4;j++){const yy=H-bottom-j*usableY/4;ctx.strokeStyle='#304865';ctx.beginPath();ctx.moveTo(left,yy);ctx.lineTo(W-right,yy);ctx.stroke();ctx.fillStyle='#9fbbd4';ctx.textAlign='right';ctx.fillText((j*maxY/4).toFixed(2),left-8,yy+5);}
  for(let j=0;j<=4;j++){const xx=left+j*usableX/4;ctx.fillStyle='#9fbbd4';ctx.textAlign='center';ctx.fillText((j*maxX/4).toFixed(1),xx,H-bottom+23);}
  ctx.strokeStyle='#42d5d1';ctx.lineWidth=3;ctx.beginPath();
  ctx.moveTo(left,H-bottom-base/maxY*usableY);
  ctx.lineTo(W-right,H-bottom-(base+r.alphaDbKm*maxX)/maxY*usableY);ctx.stroke();
  ctx.fillStyle='#e2f5ff';ctx.textAlign='center';ctx.fillText('Fiber length (km)',left+usableX/2,H-9);
  ctx.save();ctx.translate(19,top+usableY/2);ctx.rotate(-Math.PI/2);ctx.fillText('Total link loss (dB)',0,0);ctx.restore();
  el('att-result-json').value=JSON.stringify(r,null,2);
 }catch(e){
  status.textContent='입력 오류: '+e.message;status.className='eng-status error';
  ['att-alpha-out','att-fiber-out','att-total-out','att-pout-out'].forEach(id=>el(id).textContent='—');
  el('att-breakdown').textContent='입력값을 확인해 다시 계산하세요.';
  el('att-result-json').value='';
 }
}
ids.forEach(id=>el(id).addEventListener('input',calculate));
el('att-run').addEventListener('click',calculate);
el('att-json').addEventListener('click',()=>{
 const data=el('att-result-json').value;if(!data)return;
 const blob=new Blob([data],{type:'application/json'}),url=URL.createObjectURL(blob);
 const a=document.createElement('a');a.href=url;a.download='attenuation_estimator.json';a.click();
 setTimeout(()=>URL.revokeObjectURL(url),1000);
});
calculate();
