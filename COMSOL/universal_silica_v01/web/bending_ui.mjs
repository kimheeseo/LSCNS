import {estimateBending} from './bending_loss.mjs';
const $=id=>document.getElementById(id);
const keys=['radiusSmallMm','radiusLargeMm','lossSmall1550DbTurn','lossLarge1550DbTurn',
 'lossSmall1625DbTurn','lossLarge1625DbTurn','radiusMm','wavelengthNm','turns','straightLinkLossDb'];
const ids=['bend-rsmall','bend-rlarge','bend-1550small','bend-1550large','bend-1625small',
 'bend-1625large','bend-radius','bend-wave','bend-turns','bend-linkbase'];
function update(){
 try{
  if(ids.some(id=>$(id).value.trim()===''))throw Error('모든 굽힘손실 입력값을 기재하세요.');
  const input=Object.fromEntries(ids.map((v,i)=>[keys[i],Number($(v).value)]));
  const out=estimateBending(input);
  $('bend-perturn').textContent=out.lossPerTurnDb.toFixed(6)+' dB/turn';
  $('bend-added').textContent=out.bendLossDb.toFixed(5)+' dB';
  $('bend-linktotal').textContent=out.totalLinkLossDb.toFixed(5)+' dB';
  $('bend-length').textContent=out.bendLengthM.toFixed(4)+' m';
  $('bend-status').textContent='계산 완료 · 4개 입력 시험값 보간 (FEM 계산 아님)';
  $('bend-status').className='eng-status success';
  $('bend-json-cache').value=JSON.stringify(out,null,2);
  const cn=$('bend-chart'),ctx=cn.getContext('2d');
  const W=cn.width,H=cn.height,L=64,R=23,T=24,B=55,ww=W-L-R,hh=H-T-B;
  ctx.clearRect(0,0,W,H);ctx.fillStyle='#0d192c';ctx.fillRect(0,0,W,H);
  const curve=[];
  for(let i=0;i<=100;i++){
   const radius=input.radiusSmallMm+(input.radiusLargeMm-input.radiusSmallMm)*i/100;
   curve.push([radius,estimateBending({...input,radiusMm:radius,turns:1}).lossPerTurnDb]);
  }
  const minX=input.radiusSmallMm,maxX=input.radiusLargeMm,maxY=Math.max(...curve.map(q=>q[1]))*1.1;
  const x=v=>L+(v-minX)/(maxX-minX)*ww,y=v=>H-B-(v/maxY)*hh;
  ctx.lineWidth=1;ctx.font='13px system-ui';
  for(let i=0;i<=4;i++){
   const yy=T+hh*i/4;
   ctx.strokeStyle='#314a65';ctx.beginPath();ctx.moveTo(L,yy);ctx.lineTo(W-R,yy);ctx.stroke();
   ctx.fillStyle='#9ab7d2';ctx.textAlign='right';ctx.fillText((maxY*(1-i/4)).toFixed(4),L-7,yy+4);
   ctx.textAlign='center';ctx.fillText((minX+(maxX-minX)*i/4).toFixed(1),x(minX+(maxX-minX)*i/4),H-B+21);
  }
  ctx.beginPath();ctx.strokeStyle='#70b7ff';ctx.lineWidth=3;
  curve.forEach(([rx,loss],i)=>i?ctx.lineTo(x(rx),y(loss)):ctx.moveTo(x(rx),y(loss)));ctx.stroke();
  ctx.fillStyle='#ffc68d';ctx.beginPath();ctx.arc(x(input.radiusMm),y(out.lossPerTurnDb),6,0,2*Math.PI);ctx.fill();
  ctx.fillStyle='#edf7ff';ctx.textAlign='center';ctx.fillText('Bend radius (mm)',L+ww/2,H-9);
  ctx.save();ctx.translate(17,T+hh/2);ctx.rotate(-Math.PI/2);ctx.fillText('Bend loss (dB/turn)',0,0);ctx.restore();
 }catch(err){
  $('bend-status').textContent='입력 오류: '+err.message;
  $('bend-status').className='eng-status error';
  ['bend-perturn','bend-added','bend-linktotal','bend-length'].forEach(k=>$(k).textContent='—');
  $('bend-json-cache').value='';
 }
}
ids.forEach(id=>$(id).addEventListener('input',update));
$('bend-run').addEventListener('click',update);
$('bend-import-link').addEventListener('click',()=>{
 try{const x=JSON.parse(document.getElementById('att-result-json').value);$('bend-linkbase').value=x.totalLossDb;update();}
 catch(e){$('bend-status').textContent='감쇠 모듈 계산을 먼저 실행하세요.';$('bend-status').className='eng-status error';}
});
$('bend-json').addEventListener('click',()=>{
 if(!$('bend-json-cache').value)return;
 const blob=new Blob([$('bend-json-cache').value],{type:'application/json'}),uri=URL.createObjectURL(blob),a=document.createElement('a');
 a.href=uri;a.download='bend_loss_estimator.json';a.click();setTimeout(()=>URL.revokeObjectURL(uri),1000);
});
update();
