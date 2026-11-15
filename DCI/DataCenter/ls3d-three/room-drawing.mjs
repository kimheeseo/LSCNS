// Architectural-style plan generated from the tour model, not surveyed CAD.
export function roomSheet(zone,assets,mode){
 const size=mode==='server-room'?[26,24]:zone.size,[w,d]=size,[cx,cz]=zone.pos;
 const x=cx-w/2,y=cz-d/2,height=Math.max(d+28,assets.length*1.7+22);
 return {x,y,w,d,cx,cz,view:[x-9,y-10,w+36,height],sheet:{x:x-8,y:y-9,w:w+34,h:height-2}};
}
export function drawingCode(a,i){const t=a.type||'';return (/UPS|battery|PDU|transformer|switchgear|ATS|generator|busway/i.test(t)?'E':/ODF|MMR|core|spine|leaf|rack|carrier|optical/i.test(t)?'N':/CDU|chill|cool|pump|air handling/i.test(t)?'M':'A')+String(i+1).padStart(2,'0');}
export function drawRoomBase(svg,el,zone,assets,mode){
 const s=roomSheet(zone,assets,mode),{x,y,w,d,cx,cz,sheet}=s;
 const line=(x1,y1,x2,y2,attrs={})=>svg.append(el('line',{x1,y1,x2,y2,stroke:'#333','stroke-width':.07,...attrs}));
 const text=(x,y,value,size=.65,attrs={})=>svg.append(el('text',{x,y,fill:'#222','font-size':size,'font-family':'system-ui,sans-serif',...attrs},value));
 const rect=(x,y,width,height,attrs={})=>svg.append(el('rect',{x,y,width,height,fill:'none',stroke:'#222','stroke-width':.08,...attrs}));
 const defs=el('defs');const hatch=el('pattern',{id:'room-wall-hatch',width:.65,height:.65,patternUnits:'userSpaceOnUse',patternTransform:'rotate(45)'});hatch.append(el('line',{x1:0,y1:0,x2:0,y2:.65,stroke:'#777','stroke-width':.09}));defs.append(hatch);svg.append(defs);
 rect(sheet.x,sheet.y,sheet.w,sheet.h,{fill:'#fff','stroke-width':.15});
 // Raised floor joints are drawn lightly beneath equipment.
 for(let xx=x+3.2;xx<x+w;xx+=3.2)line(xx,y,xx,y+d,{stroke:'#ddd','stroke-width':.035});
 for(let yy=y+3.2;yy<y+d;yy+=3.2)line(x,yy,x+w,yy,{stroke:'#ddd','stroke-width':.035});
 const outdoor=['utility','cooling'].includes(zone.id),wall=.32,door=3.3;
 if(outdoor){rect(x,y,w,d,{'stroke-dasharray':'.45 .2','stroke-width':.15});}
 else{
  const wallRect=(xx,yy,ww,dd)=>rect(xx,yy,ww,dd,{fill:'url(#room-wall-hatch)','stroke-width':.13});
  wallRect(x-wall/2,y-wall/2,w+wall,wall);wallRect(x-wall/2,y,wall,d);wallRect(x+w-wall/2,y,wall,d);
  wallRect(x-wall/2,y+d-wall/2,w/2-door/2+wall/2,wall);wallRect(cx+door/2,y+d-wall/2,w/2-door/2+wall/2,wall);
  // Matches architecture.mjs: 3.3-unit double doorway at the center of +Z facade.
  const r=door/2,yy=y+d;
  line(cx-r,yy,cx-r,yy+r,{'stroke-width':.12});line(cx+r,yy,cx+r,yy+r,{'stroke-width':.12});
  svg.append(el('path',{d:`M ${cx-r} ${yy+r} A ${r} ${r} 0 0 0 ${cx} ${yy} M ${cx+r} ${yy+r} A ${r} ${r} 0 0 1 ${cx} ${yy}`,fill:'none',stroke:'#444','stroke-width':.06}));
  text(cx,yy+2.6,'D01 · 양개 출입문',.58,{'text-anchor':'middle'});
  for(const sx of [-1,1])rect(cx+sx*door/2-.22,yy-.22,.44,.44,{fill:'#222','stroke-width':.07});
 }
 // Structural axes and dimension chains: model coordinates (u), never millimeters.
 function dimH(a,b,yy,label){line(a,yy,b,yy);for(const xx of [a,b]){line(xx,y-.3,xx,yy-.45,{stroke:'#888','stroke-width':.04});line(xx-.18,yy+.2,xx+.18,yy-.2,{'stroke-width':.11});}text((a+b)/2,yy-.25,label,.58,{'text-anchor':'middle'});}
 function dimV(a,b,xx,label){line(xx,a,xx,b);for(const yy of [a,b]){line(x-.3,yy,xx-.45,yy,{stroke:'#888','stroke-width':.04});line(xx-.18,yy+.2,xx+.18,yy-.2,{'stroke-width':.11});}text(xx-.35,(a+b)/2,label,.58,{transform:`rotate(-90 ${xx-.35} ${(a+b)/2})`,'text-anchor':'middle'});}
 dimH(x,x+w,y-4.6,w.toFixed(1)+' u');dimV(y,y+d,x-4.6,d.toFixed(1)+' u');
 const xs=[x,cx,x+w],ys=[y,cz,y+d];
 for(let i=0;i<2;i++){dimH(xs[i],xs[i+1],y-2.7,(w/2).toFixed(1));dimV(ys[i],ys[i+1],x-2.7,(d/2).toFixed(1));}
 for(let i=0;i<3;i++){const xx=xs[i],yy=ys[i];line(xx,y-6,xx,y+d+.5,{stroke:'#aaa','stroke-width':.035,'stroke-dasharray':'.8 .18 .12 .18'});line(x-6,yy,x+w+.5,yy,{stroke:'#aaa','stroke-width':.035,'stroke-dasharray':'.8 .18 .12 .18'});svg.append(el('circle',{cx:xx,cy:y-6.6,r:.48,fill:'white',stroke:'#444','stroke-width':.07}));text(xx,y-6.4,String(i+1),.55,{'text-anchor':'middle'});svg.append(el('circle',{cx:x-6.6,cy:yy,r:.48,fill:'white',stroke:'#444','stroke-width':.07}));text(x-6.6,yy+.2,['A','B','C'][i],.55,{'text-anchor':'middle'});}
 if(!outdoor){
  // Overhead routes follow the architecture kit; dashed lines are tray outlines, not cable connectivity.
  for(const sx of [-1,1])rect(cx+sx*w*.29-.38,cz-d*.39,.76,d*.78,{'stroke-dasharray':'.35 .17',stroke:'#999','stroke-width':.05});
  rect(cx-w*.34,cz-d*.22-.38,w*.68,.76,{'stroke-dasharray':'.35 .17',stroke:'#999','stroke-width':.05});
 }
 // Equipment schedule, aligned with numbered symbols in the plan.
 const lx=x+w+3,ly=y-5,tw=20;
 text(lx,ly,'설비 일람표 / EQUIPMENT',.85,{'font-weight':700});
 line(lx,ly+.5,lx+tw,ly+.5);text(lx,ly+1.5,'번호',.58);text(lx+2.5,ly+1.5,'설비명 · 분류',.58);
 assets.forEach((a,i)=>{const yy=ly+3+i*1.7;const row=el('g',{'data-schedule-asset':a.id});row.append(el('text',{x:lx,y:yy,'font-size':.6,fill:'#222'},drawingCode(a,i)));const name=a.name.length>30?a.name.slice(0,29)+'…':a.name;row.append(el('text',{x:lx+2.5,y:yy,'font-size':.57,fill:'#222'},name));row.append(el('text',{x:lx+2.5,y:yy+.65,'font-size':.43,fill:'#666'},a.type));svg.append(row);line(lx,yy+.95,lx+tw,yy+.95,{stroke:'#bbb','stroke-width':.035});});
 const notesY=Math.max(y+d+6,ly+4+assets.length*1.7);
 text(x,notesY,'도면 기호',.7,{'font-weight':700});
 text(x,notesY+1.2,'해치: 벽 단면 / 호: 문 여닫이 / 점선: 상부 트레이·구조 기준선',.52);
 text(x,notesY+2.3,'설비 외곽은 모델 배치·크기 기준. 내부 기호는 유형별 개념 표시.',.52);
 const ty=sheet.y+sheet.h-5;
 rect(sheet.x,ty,sheet.w,5,{'stroke-width':.12});line(sheet.x+sheet.w*.62,ty,sheet.x+sheet.w*.62,ty+5);
 text(sheet.x+1,ty+1.8,(zone.ko||zone.n)+' · 상세 평면도',1,{'font-weight':700});text(sheet.x+1,ty+3.5,'모델 좌표 단위 u · 실측 mm/시공도면 아님',.65);
 text(sheet.x+sheet.w*.64,ty+1.6,'DC-'+zone.id.toUpperCase()+'-PLAN',.8,{'font-weight':700});text(sheet.x+sheet.w*.64,ty+3.3,'축척: 화면 맞춤 / X → · Z ↓',.58);
 return s.view;
}
export function drawRoomEquipment(g,el,a,index){
 const {x,z,w,d}=a,left=x-w/2,top=z-d/2,sw=.065,type=a.type||'';
 const add=(tag,attrs={},text)=>{const n=el(tag,{stroke:'#333','stroke-width':sw,fill:'none',...attrs},text);g.append(n);return n;};
 // The first rect is the shared selection outline used by the tour workspace.
 add('rect',{x:left,y:top,width:w,height:d,fill:'#fff','stroke-width':.16});
 add('rect',{x:left+.18,y:top+.18,width:Math.max(.1,w-.36),height:Math.max(.1,d-.36),'stroke-width':.045});
 if(/battery/i.test(type)){
  for(let ix=0;ix<4;ix++)for(let iy=0;iy<2;iy++){const cw=(w-.8)/4,ch=(d-.7)/2;add('rect',{x:left+.35+ix*cw,y:top+.25+iy*ch,width:cw-.15,height:ch-.12});add('line',{x1:left+.5+ix*cw,y1:top+.4+iy*ch,x2:left+.75+ix*cw,y2:top+.4+iy*ch});}
 }else if(/UPS|PDU|switchgear|ATS|busway/i.test(type)){
  const n=Math.max(2,Math.round(w/1.5));for(let i=1;i<n;i++)add('line',{x1:left+i*w/n,y1:top+.2,x2:left+i*w/n,y2:top+d-.2});
  for(let i=0;i<n;i++){add('rect',{x:left+(i+.25)*w/n,y:top+.45,width:w/n*.5,height:Math.min(.65,d*.22)});add('line',{x1:left+(i+.7)*w/n,y1:top+d*.5,x2:left+(i+.7)*w/n,y2:top+d*.72});}
 }else if(/rack|ODF|MMR|core|spine|leaf|carrier|optical/i.test(type)){
  for(let zz=top+.45;zz<top+d-.2;zz+=.38)add('line',{x1:left+.3,y1:zz,x2:left+w-.3,y2:zz,'stroke-width':.045});
  for(const side of [left+.2,left+w-.2])add('line',{x1:side,y1:top+.2,x2:side,y2:top+d-.2,'stroke-width':.1});
  for(let xx=left+.4;xx<left+w-.3;xx+=.38)add('circle',{cx:xx,cy:top+d-.32,r:.055,fill:'#333'});
 }else if(/CDU|chill|cool|pump|air handling/i.test(type)){
  const r=Math.min(w,d)*.28;add('circle',{cx:x,cy:z,r});for(let i=0;i<8;i++){const t=i*Math.PI/4;add('line',{x1:x,y1:z,x2:x+Math.cos(t)*r,y2:z+Math.sin(t)*r});}add('circle',{cx:x,cy:z,r:r*.22,fill:'#fff'});
 }else if(/transformer|generator/i.test(type)){
  add('circle',{cx:x-w*.15,cy:z,r:Math.min(w,d)*.23});add('circle',{cx:x+w*.15,cy:z,r:Math.min(w,d)*.23});
 }else{add('line',{x1:left+.3,y1:top+.3,x2:left+w-.3,y2:top+d-.3});add('line',{x1:left+w-.3,y1:top+.3,x2:left+.3,y2:top+d-.3});}
 const code=drawingCode(a,index),cy=top-.75;
 add('line',{x1:x,y1:top,x2:x,y2:cy+.35,stroke:'#444','stroke-width':.045});
 add('rect',{x:x-1.05,y:cy-.55,width:2.1,height:.85,fill:'white','stroke-width':.06});
 add('text',{x,y:cy+.02,'text-anchor':'middle','font-family':'system-ui,sans-serif','font-size':.56,fill:'#111',stroke:'none'},code);
 const short=a.name.split(' · ')[0].replace('Network Rack ','NET ').replace('Storage Rack ','STG ').replace('GPU Rack ','GPU ');
 add('text',{x,y:top+d+.7,'text-anchor':'middle','font-size':.48,fill:'#222',stroke:'none'},short.length>24?short.slice(0,23)+'…':short);
}

function illustratedColors(a){const t=a.type||'';if(/UPS|PDU|battery|busway|transformer|switchgear|generator|ATS/i.test(t))return ['#e9ae4d','#fff0cf','#80520d'];if(/CDU|cool|chill|pump|air handling/i.test(t))return ['#52b49b','#d9f5ed','#1d6556'];if(/storage/i.test(t))return ['#a790d0','#ede6fa','#574079'];return ['#4e96be','#dbeefa','#1e506c'];}
export function drawIllustratedBase(svg,el,zone,assets,mode){
 const s=roomSheet(zone,assets,mode),{x,y,w,d,cx,cz,sheet}=s;
 const add=(tag,attrs,text)=>{const n=el(tag,attrs,text);svg.append(n);return n;};
 add('rect',{x:sheet.x,y:sheet.y,width:sheet.w,height:sheet.h,rx:.8,fill:'#f5f9fc',stroke:'#c5d6e0','stroke-width':.12});
 add('text',{x,y:y-5.5,'font-size':1.5,'font-weight':700,fill:'#183f55'},(zone.ko||zone.n)+' · 그림형 배치도');
 add('text',{x,y:y-3.6,'font-size':.65,fill:'#547383'},'설비 모양·색상으로 구분하는 공간 안내 · 번호는 오른쪽 설비 목록과 대응');
 add('rect',{x:x+.45,y:y+.5,width:w,height:d,fill:'#c6d6df',rx:.25});
 add('rect',{x,y,width:w,height:d,fill:'#edf4f7',stroke:'#7795a7','stroke-width':.35,rx:.15});
 for(let xx=x+3.2;xx<x+w;xx+=3.2)add('line',{x1:xx,y1:y,x2:xx,y2:y+d,stroke:'#d7e4eb','stroke-width':.05});
 for(let yy=y+3.2;yy<y+d;yy+=3.2)add('line',{x1:x,y1:yy,x2:x+w,y2:yy,stroke:'#d7e4eb','stroke-width':.05});
 add('rect',{x:cx-1.65,y:y+d-.3,width:3.3,height:.6,fill:'#edf4f7'});
 add('path',{d:`M ${cx-1.65} ${y+d} L ${cx-1.65} ${y+d+1.65} M ${cx+1.65} ${y+d} L ${cx+1.65} ${y+d+1.65}`,fill:'none',stroke:'#36748a','stroke-width':.12});
 add('text',{x:cx,y:y+d+3,'text-anchor':'middle','font-size':.85,fill:'#325c72'},'출입구');
 const lx=x+w+3,ly=y-1;
 add('text',{x:lx,y:ly-1,'font-size':1,'font-weight':700,fill:'#234e66'},'공간 안의 설비');
 assets.forEach((a,i)=>{const yy=ly+1.5+i*1.7,[c,bg,ink]=illustratedColors(a);add('rect',{x:lx,y:yy-.8,width:2.3,height:1.05,rx:.2,fill:bg,stroke:c,'stroke-width':.07});add('text',{x:lx+1.15,y:yy-.08,'font-size':.58,'text-anchor':'middle',fill:ink,'font-weight':700},drawingCode(a,i));add('text',{x:lx+2.8,y:yy-.1,'font-size':.57,fill:'#24485f'},a.name.length>27?a.name.slice(0,26)+'…':a.name);add('text',{x:lx+2.8,y:yy+.57,'font-size':.45,fill:'#66808f'},a.type);});
 const by=sheet.y+sheet.h-5;
 add('line',{x1:sheet.x+1,y1:by,x2:sheet.x+sheet.w-1,y2:by,stroke:'#c1d4df','stroke-width':.08});
 add('text',{x:sheet.x+1.5,y:by+1.7,'font-size':.72,fill:'#315d74'},'파랑: 랙·네트워크 / 보라: 스토리지 / 노랑: 전력 / 초록: 냉각');
 add('text',{x:sheet.x+1.5,y:by+3.25,'font-size':.62,fill:'#607f91'},'배치·외곽 크기는 3D 모델 기준 · 설비 내부 그림은 기능을 설명하는 개념 표현');
 return s.view;
}
export function drawIllustratedEquipment(g,el,a,index){
 const {x,z,w,d}=a,l=x-w/2,t=z-d/2,[c,bg,ink]=illustratedColors(a),type=a.type||'';
 const add=(tag,attrs={},text)=>{const n=el(tag,{fill:'none',stroke:ink,'stroke-width':.065,...attrs},text);g.append(n);return n;};
 add('rect',{x:l,y:t,width:w,height:d,rx:.2,fill:bg,stroke:c,'stroke-width':.15});
 add('path',{d:`M ${l+.15} ${t+d-.35} H ${l+w-.15} V ${t+d-.1} H ${l+.15} Z`,fill:c,stroke:'none'});
 if(/battery/i.test(type)){for(let i=0;i<4;i++){const xx=l+.3+i*(w-.5)/4,cw=(w-.8)/4;add('rect',{x:xx,y:t+.6,width:cw-.1,height:Math.max(.4,d-1.3),rx:.08,fill:'#fff9e8'});add('rect',{x:xx+cw*.3,y:t+.4,width:cw*.35,height:.2,fill:ink,stroke:'none'});add('text',{x:xx+cw*.4,y:z+.15,'text-anchor':'middle','font-size':.65,fill:ink,stroke:'none'},'+');}}
 else if(/CDU|cool|chill|pump|air handling/i.test(type)){add('circle',{cx:x,cy:z,r:Math.min(w,d)*.29,fill:'#f6fffc'});for(let i=0;i<6;i++){const q=i*Math.PI/3,r=Math.min(w,d)*.26;add('path',{d:`M ${x} ${z} Q ${x+Math.cos(q+.4)*r} ${z+Math.sin(q+.4)*r} ${x+Math.cos(q)*r} ${z+Math.sin(q)*r}`,stroke:c,'stroke-width':.17});}add('circle',{cx:x,cy:z,r:.2,fill:c});}
 else if(/rack|ODF|MMR|spine|leaf|core|carrier|optical/i.test(type)){for(let yy=t+.4;yy<t+d-.5;yy+=.42){add('rect',{x:l+.3,y:yy,width:w-.6,height:.28,rx:.04,fill:'#f8fcff',stroke:c,'stroke-width':.04});for(let xx=l+.5;xx<l+w-.4;xx+=.4)add('circle',{cx:xx,cy:yy+.14,r:.045,fill:c,stroke:'none'});}add('rect',{x:l+.18,y:t+.2,width:.12,height:d-.5,fill:c,stroke:'none'});}
 else{const n=Math.max(2,Math.round(w/1.6));for(let i=0;i<n;i++){const xx=l+.2+i*(w-.4)/n,cw=(w-.4)/n;add('rect',{x:xx,y:t+.25,width:cw-.12,height:d-.7,rx:.05,fill:'#fffbf2',stroke:c,'stroke-width':.05});add('rect',{x:xx+.18,y:t+.5,width:Math.max(.2,cw-.5),height:.5,fill:ink,stroke:'none'});add('path',{d:`M ${xx+cw*.5} ${t+d*.4} l -.22 .55 h .35 l -.22 .55`,stroke:c,'stroke-width':.13});}}
 add('rect',{x:x-1.15,y:t-.95,width:2.3,height:.9,rx:.2,fill:bg,stroke:c,'stroke-width':.06});add('text',{x,y:t-.28,'text-anchor':'middle','font-size':.58,'font-weight':700,fill:ink,stroke:'none'},drawingCode(a,index));
}
