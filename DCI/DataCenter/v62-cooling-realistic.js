(function browserPatch(){
'use strict';
const I18N={
 ko:{title:'Cooling Architecture · 3D Engineering View',sub:'Facility Cooling → CDU → Rack Manifold → GPU/CPU Cold Plate',cold:'냉각수 공급',warm:'가열수 환수',note:'단일 고해상도 3D Cooling Architecture입니다. 구조 시각화는 이해를 위한 것이며, 냉각수 구조 참고도이며, 현재 서버의 냉각 방식 및 수량은 계산 BOM을 기준으로 합니다. Cold Plate 호환성을 뜻하지 않습니다.',facility:'Facility Cooling',cdu:'CDU',manifold:'Rack Manifold',rack:'Liquid-Cooled GPU Rack',coldPlate:'GPU / CPU Cold Plate'},
 en:{title:'Cooling Architecture · 3D Engineering View',sub:'Facility Cooling → CDU → Rack Manifold → GPU/CPU Cold Plate',cold:'Cold supply',warm:'Warm return',note:'Single high-resolution 3D Cooling Architecture. The visual explains the system topology; use the calculated BOM for the selected cooling mode and quantities. This reference does not establish cold-plate compatibility.',facility:'Facility Cooling',cdu:'CDU',manifold:'Rack Manifold',rack:'Liquid-Cooled GPU Rack',coldPlate:'GPU / CPU Cold Plate'},
 zh:{title:'Cooling Architecture · 3D Engineering View',sub:'Facility Cooling → CDU → Rack Manifold → GPU/CPU Cold Plate',cold:'冷却液供给',warm:'热回流',note:'单一高分辨率 3D 冷却架构视图。图像用于理解系统结构，实际设计值以下方动态 Rack Heat / ΔT / Flow / CDU 计算为准。',facility:'Facility Cooling',cdu:'CDU',manifold:'Rack Manifold',rack:'Liquid-Cooled GPU Rack',coldPlate:'GPU / CPU Cold Plate'},
 ja:{title:'Cooling Architecture · 3D Engineering View',sub:'Facility Cooling → CDU → Rack Manifold → GPU/CPU Cold Plate',cold:'冷却水供給',warm:'温水戻り',note:'単一の高解像度 3D Cooling Architecture です。図は構成理解用で、設計値は下部の動的 Rack Heat / ΔT / Flow / CDU 計算を基準にします。',facility:'Facility Cooling',cdu:'CDU',manifold:'Rack Manifold',rack:'Liquid-Cooled GPU Rack',coldPlate:'GPU / CPU Cold Plate'},
 de:{title:'Cooling Architecture · 3D Engineering View',sub:'Facility Cooling → CDU → Rack Manifold → GPU/CPU Cold Plate',cold:'Kaltwasser Vorlauf',warm:'Warmer Rücklauf',note:'Ein einzelnes hochauflösendes 3D-Kühlarchitekturdiagramm. Die Grafik erläutert die Topologie; für die Auslegung gelten die dynamischen Rack-Heat-/ΔT-/Flow-/CDU-Berechnungen unten.',facility:'Facility Cooling',cdu:'CDU',manifold:'Rack Manifold',rack:'Liquid-Cooled GPU Rack',coldPlate:'GPU / CPU Cold Plate'}
};
function lang(){
 const raw=((window.__dcBomUiLang||document.documentElement.getAttribute('data-dc-bom-ui-lang')||document.documentElement.lang||'ko')+'').toLowerCase();
 if(raw.includes('zh')||raw.includes('cn'))return'zh';
 if(raw.includes('ja')||raw.includes('jp'))return'ja';
 if(raw.includes('de'))return'de';
 if(raw.includes('en'))return'en';
 return'ko';
}
function T(){return I18N[lang()]||I18N.ko}
function rounded(ctx,x,y,w,h,r){
 r=Math.min(r,w/2,h/2);ctx.beginPath();ctx.moveTo(x+r,y);ctx.arcTo(x+w,y,x+w,y+h,r);ctx.arcTo(x+w,y+h,x,y+h,r);ctx.arcTo(x,y+h,x,y,r);ctx.arcTo(x,y,x+w,y,r);ctx.closePath();
}
function isoBox(ctx,x,y,w,h,d,front,top,side){
 ctx.fillStyle=top;ctx.beginPath();ctx.moveTo(x,y);ctx.lineTo(x+d,y-d*.48);ctx.lineTo(x+w+d,y-d*.48);ctx.lineTo(x+w,y);ctx.closePath();ctx.fill();
 ctx.fillStyle=side;ctx.beginPath();ctx.moveTo(x+w,y);ctx.lineTo(x+w+d,y-d*.48);ctx.lineTo(x+w+d,y+h-d*.48);ctx.lineTo(x+w,y+h);ctx.closePath();ctx.fill();
 ctx.fillStyle=front;ctx.fillRect(x,y,w,h);
}
function arrow(ctx,x1,y1,x2,y2,color,width){
 ctx.save();ctx.strokeStyle=color;ctx.fillStyle=color;ctx.lineWidth=width;ctx.lineCap='round';ctx.beginPath();ctx.moveTo(x1,y1);ctx.lineTo(x2,y2);ctx.stroke();
 const a=Math.atan2(y2-y1,x2-x1),s=13+width;ctx.beginPath();ctx.moveTo(x2,y2);ctx.lineTo(x2-s*Math.cos(a-.5),y2-s*Math.sin(a-.5));ctx.lineTo(x2-s*Math.cos(a+.5),y2-s*Math.sin(a+.5));ctx.closePath();ctx.fill();ctx.restore();
}
function label(ctx,text,x,y,size,align){
 ctx.save();ctx.font='700 '+size+'px system-ui,-apple-system,Segoe UI,sans-serif';ctx.textAlign=align||'center';ctx.textBaseline='middle';ctx.fillStyle='#e9f4ff';ctx.shadowColor='rgba(0,0,0,.65)';ctx.shadowBlur=5;ctx.fillText(text,x,y);ctx.restore();
}
function renderScene(canvas,scale){
 const dpr=Math.min(3,window.devicePixelRatio||1)*(scale||1);
 const W=1600,H=900;canvas.width=Math.round(W*dpr);canvas.height=Math.round(H*dpr);const ctx=canvas.getContext('2d');ctx.setTransform(dpr,0,0,dpr,0,0);
 const L=T();
 const bg=ctx.createLinearGradient(0,0,0,H);bg.addColorStop(0,'#07192b');bg.addColorStop(.56,'#081827');bg.addColorStop(1,'#030914');ctx.fillStyle=bg;ctx.fillRect(0,0,W,H);
 const glow=ctx.createRadialGradient(1060,320,30,1060,320,600);glow.addColorStop(0,'rgba(42,104,154,.21)');glow.addColorStop(1,'rgba(6,17,31,0)');ctx.fillStyle=glow;ctx.fillRect(0,0,W,H);

 // ceiling/light strips
 for(let i=0;i<5;i++){ctx.fillStyle='rgba(147,207,244,.13)';ctx.fillRect(690+i*170,48,105,7)}

 // raised-floor perspective
 ctx.strokeStyle='rgba(102,145,182,.13)';ctx.lineWidth=1;
 for(let y=620;y<900;y+=42){ctx.beginPath();ctx.moveTo(260,y);ctx.lineTo(1580,y);ctx.stroke()}
 for(let x=250;x<1600;x+=90){ctx.beginPath();ctx.moveTo(800+(x-800)*.38,500);ctx.lineTo(x,900);ctx.stroke()}

 // Facility cooling plant
 ctx.save();ctx.shadowColor='rgba(0,0,0,.55)';ctx.shadowBlur=18;ctx.shadowOffsetY=10;
 isoBox(ctx,80,250,250,260,45,'#173149','#234963','#102437');
 ctx.restore();
 ctx.fillStyle='#0b1a29';ctx.fillRect(105,300,200,110);
 for(let i=0;i<4;i++){ctx.fillStyle=i%2?'#31546b':'#203e55';ctx.fillRect(120+i*46,316,30,78)}
 ctx.strokeStyle='#7fb3d7';ctx.lineWidth=3;ctx.beginPath();ctx.arc(205,450,34,0,Math.PI*2);ctx.stroke();ctx.beginPath();ctx.arc(205,450,15,0,Math.PI*2);ctx.stroke();
 label(ctx,L.facility,205,215,22);

 // CDU cabinet
 ctx.save();ctx.shadowColor='rgba(0,0,0,.6)';ctx.shadowBlur=20;ctx.shadowOffsetY=12;isoBox(ctx,430,315,235,315,48,'#16283c','#294560','#0d1d2d');ctx.restore();
 ctx.fillStyle='#0a1521';rounded(ctx,460,355,175,94,10);ctx.fill();
 for(let i=0;i<5;i++){ctx.strokeStyle='rgba(97,173,220,.55)';ctx.lineWidth=2;ctx.beginPath();ctx.moveTo(478,376+i*14);ctx.lineTo(615,376+i*14);ctx.stroke()}
 ctx.fillStyle='#1c3f56';rounded(ctx,466,480,162,108,8);ctx.fill();
 ctx.strokeStyle='#58b6e8';ctx.lineWidth=4;ctx.beginPath();ctx.arc(515,532,26,0,Math.PI*2);ctx.stroke();ctx.beginPath();ctx.arc(579,532,26,0,Math.PI*2);ctx.stroke();
 ctx.fillStyle='#68d8ff';ctx.fillRect(592,600,11,11);label(ctx,L.cdu,548,280,24);

 // manifold
 ctx.save();ctx.shadowColor='rgba(0,0,0,.5)';ctx.shadowBlur=15;isoBox(ctx,735,605,610,65,26,'#132a3e','#244a63','#0c1a2a');ctx.restore();
 label(ctx,L.manifold,1040,700,20);

 // racks
 const rackXs=[800,1010,1220];
 rackXs.forEach((x,ri)=>{
   ctx.save();ctx.shadowColor='rgba(0,0,0,.62)';ctx.shadowBlur=22;ctx.shadowOffsetY=13;isoBox(ctx,x,205,165,385,36,'#172638','#263b51','#0d1724');ctx.restore();
   ctx.fillStyle='#0a111b';ctx.fillRect(x+14,225,136,340);
   for(let u=0;u<8;u++){
     const yy=240+u*39;const grad=ctx.createLinearGradient(x+20,yy,x+142,yy);
     grad.addColorStop(0,'#263b50');grad.addColorStop(.55,'#162635');grad.addColorStop(1,'#0f1b29');ctx.fillStyle=grad;rounded(ctx,x+20,yy,123,29,4);ctx.fill();
     ctx.fillStyle='#4bc0ff';ctx.fillRect(x+29,yy+10,8,8);
     for(let p=0;p<4;p++){ctx.fillStyle='rgba(119,159,187,.45)';ctx.fillRect(x+73+p*13,yy+11,8,6)}
   }
   ctx.fillStyle='#2c4a63';ctx.fillRect(x+18,572,130,8);
   ctx.fillStyle='#3ea9ea';ctx.fillRect(x+6,248,5,272);ctx.fillStyle='#ef6759';ctx.fillRect(x+153,248,5,272);
   label(ctx,(ri===1?L.rack:'GPU Rack '+(ri+1)),x+82,170,ri===1?19:16);
 });

 // cold-plate cutaway on the far right
 ctx.save();ctx.shadowColor='rgba(0,0,0,.55)';ctx.shadowBlur=16;isoBox(ctx,1420,325,120,175,26,'#193247','#2b526a','#0d1e2d');ctx.restore();
 ctx.fillStyle='#102230';rounded(ctx,1438,354,85,70,8);ctx.fill();ctx.strokeStyle='#49c7ff';ctx.lineWidth=5;ctx.beginPath();ctx.moveTo(1450,375);ctx.lineTo(1510,375);ctx.lineTo(1510,405);ctx.lineTo(1450,405);ctx.closePath();ctx.stroke();
 ctx.fillStyle='#5bd0ff';rounded(ctx,1454,379,52,22,5);ctx.fill();label(ctx,L.coldPlate,1483,290,16);

 // main chilled supply + warm return
 arrow(ctx,330,390,430,390,'#36b7ff',12);
 arrow(ctx,665,420,735,420,'#36b7ff',12);
 arrow(ctx,735,420,1350,420,'#36b7ff',12);
 arrow(ctx,1350,465,665,465,'#ff6d5c',12);
 arrow(ctx,430,465,330,465,'#ff6d5c',12);

 // risers to racks
 rackXs.forEach(x=>{arrow(ctx,x+8,615,x+8,515,'#36b7ff',7);arrow(ctx,x+158,515,x+158,615,'#ff6d5c',7)});
 // branch to cold plate
 arrow(ctx,1350,420,1418,390,'#36b7ff',7);arrow(ctx,1418,445,1350,465,'#ff6d5c',7);

 // floor return/supply bands
 ctx.save();ctx.globalAlpha=.28;ctx.fillStyle='#2aa7f2';ctx.fillRect(690,757,700,8);ctx.fillStyle='#ff6250';ctx.fillRect(690,781,700,8);ctx.restore();
 label(ctx,L.cold,810,742,15,'left');label(ctx,L.warm,810,812,15,'left');

 // flow labels
 ctx.fillStyle='rgba(6,18,32,.82)';rounded(ctx,370,360,250,70,10);ctx.fill();label(ctx,'Facility Water Loop',495,378,16);label(ctx,'Primary ↔ Secondary Heat Exchange',495,405,12);
 ctx.fillStyle='rgba(6,18,32,.82)';rounded(ctx,840,78,530,74,10);ctx.fill();label(ctx,'Technology Cooling Loop · Direct-to-Chip',1105,101,19);label(ctx,'CDU → Manifold → QD → Cold Plate → Return',1105,128,13);

 // small QD connectors
 rackXs.forEach(x=>{ctx.fillStyle='#c9e9ff';ctx.beginPath();ctx.arc(x+8,525,6,0,Math.PI*2);ctx.fill();ctx.fillStyle='#ffd0c8';ctx.beginPath();ctx.arc(x+158,525,6,0,Math.PI*2);ctx.fill()});
}
function lightbox(){
 let m=document.getElementById('v631-cooling-lightbox');if(m)return m;
 m=document.createElement('div');m.id='v631-cooling-lightbox';m.hidden=true;m.innerHTML='<div class="v631-lb-box"><canvas></canvas></div><button type="button" aria-label="Close">×</button>';document.body.appendChild(m);
 const close=()=>{m.hidden=true};m.querySelector('button').onclick=close;m.addEventListener('click',e=>{if(e.target===m)close()});document.addEventListener('keydown',e=>{if(e.key==='Escape')close()});
 return m;
}
function apply(){
 const panel=document.getElementById('v53-cooling-architecture');if(!panel)return false;
 const wrap=panel.querySelector('.v53-cool-svg-wrap');if(!wrap)return false;
 const L=T();
 if(wrap.dataset.v631canvas==='1'){
   const title=wrap.querySelector('[data-v631-title]'),sub=wrap.querySelector('[data-v631-sub]'),cold=wrap.querySelector('[data-v631-cold]'),warm=wrap.querySelector('[data-v631-warm]'),note=panel.querySelector('.v631-cool-note');
   if(title)title.textContent=L.title;if(sub)sub.textContent=L.sub;if(cold)cold.textContent=L.cold;if(warm)warm.textContent=L.warm;if(note)note.textContent=L.note;
   const canvas=wrap.querySelector('canvas');if(canvas)renderScene(canvas,1);
   return true;
 }
 wrap.innerHTML='<div class="v631-cool3d"><div class="v631-cool3d-head"><span data-v631-title>'+L.title+'</span><small data-v631-sub>'+L.sub+'</small></div><canvas aria-label="'+L.title+'"></canvas><div class="v631-cool3d-legend"><span class="v631-leg"><i class="v631-dot v631-cold"></i><span data-v631-cold>'+L.cold+'</span></span><span class="v631-leg"><i class="v631-dot v631-warm"></i><span data-v631-warm>'+L.warm+'</span></span></div></div>';
 wrap.dataset.v631canvas='1';
 let note=panel.querySelector('.v631-cool-note');if(!note){note=document.createElement('div');note.className='v631-cool-note';wrap.insertAdjacentElement('afterend',note)}note.textContent=L.note;
 const canvas=wrap.querySelector('canvas');renderScene(canvas,1);
 canvas.onclick=()=>{const m=lightbox();m.hidden=false;renderScene(m.querySelector('canvas'),1.45)};
 return true;
}
function start(){
 apply();
 let lastW=window.innerWidth;
 window.addEventListener('resize',()=>{if(Math.abs(window.innerWidth-lastW)>40){lastW=window.innerWidth;setTimeout(apply,100)}});
 document.addEventListener('change',()=>setTimeout(apply,80),true);
 document.addEventListener('click',e=>{if(e.target&&e.target.closest&&e.target.closest('[data-lang],[data-language],.lang,.language'))setTimeout(apply,80)},true);
 const mo=new MutationObserver(()=>setTimeout(apply,0));mo.observe(document.documentElement,{childList:true,subtree:true});
 setInterval(apply,1500);
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start,{once:true});else start();
})();