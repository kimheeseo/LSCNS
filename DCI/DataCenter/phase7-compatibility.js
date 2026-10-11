/* Evidence-aware packet validation. PASS means supplied evidence passed rules, not purchasing approval. */
(function(root,factory){const api=factory();if(typeof module==='object'&&module.exports)module.exports=api;if(root)root.LSDCPhase7Compatibility=api;})(typeof globalThis!=='undefined'?globalThis:this,function(){
'use strict';
const s=v=>String(v??'').trim(),n=v=>v==null||typeof v==='boolean'||s(v)===''?null:Number.isFinite(Number(v))?Number(v):null;
const protocol=v=>{const p=s(v).toLowerCase();return p==='ib'?'infiniband':p};
function speedG(v){const m=s(v).match(/^(\d+(?:\.\d+)?)\s*(g(?:bps|b\/s)?|t(?:bps|b\/s)?)?$/i);return m?Number(m[1])*(/^t/i.test(m[2]||'')?1000:1):null}
const optical=l=>!/\b(dac|aec|copper)\b/i.test(s(l.type)+' '+s(l.media||l.fiber));
const evidenceUrl=v=>/^https?:\/\/[^\s/]+\/.+/i.test(s(v));
function validate(packet){
 const blocking=[],warnings=[],info=[];
 const finish=()=>{const status=blocking.length?'FAIL':warnings.length?'UNKNOWN':'PASS';return {compatible:status==='PASS',nonBlocking:blocking.length===0,status,blocking,warnings,info,profile:packet?.facility?.profile||null,activeLinks:Array.isArray(packet?.optical?.topology?.links)?packet.optical.topology.links.filter(x=>!x.down).length:0}};
 if(!packet||packet.schemaVersion!=='lsdc-twin-bom/1.0'){blocking.push('지원 가능한 lsdc-twin-bom/1.0 packet이 아닙니다.');return finish()}
 const o=packet.optical||{},top=o.topology||{},links=Array.isArray(top.links)?top.links:[],nodes=Array.isArray(top.nodes)?top.nodes:[];
 const designProtocol=protocol(packet.network?.protocol||top.protocol||o.protocol||o.assumptions?.protocol),mode=s(o.cpoMode).toLowerCase();
 const speed=top.portSpeed||o.assumptions?.portSpeed,designSpeed=speedG(speed);
 if(!['ethernet','infiniband'].includes(designProtocol))warnings.push('설계 네트워크 프로토콜이 없거나 지원 여부가 미확정입니다.');
 if(!['cpo','pluggable'].includes(mode))warnings.push('CPO/Pluggable 모드가 없거나 지원 여부가 미확정입니다.');
 if(![400,800,1600].includes(designSpeed))warnings.push('포트 속도 정합성을 확인할 수 없습니다: '+s(speed));
 const products=Array.isArray(packet.products)?packet.products:Object.values(packet.products||{});
 const used=new Map(),active=links.filter(x=>!x.down);
 for(const l of active){
  const label='링크 '+s(l.id)+': ',endpoints=[l.from,l.to].map(id=>nodes.find(x=>x.id===id));
  if(endpoints.some(x=>!x))blocking.push(label+'endpoint 노드가 없습니다.');
  for(const ep of endpoints.filter(Boolean)){
   used.set(ep.id,(used.get(ep.id)||0)+1);
   const p=protocol(ep.protocol||l.protocol);
   if(!p)warnings.push(label+'endpoint '+ep.id+' 프로토콜이 미확정입니다.');
   else if(!['ethernet','infiniband'].includes(p))warnings.push(label+'endpoint 프로토콜 지원 여부가 미확정입니다.');
   else if(designProtocol&&p!==designProtocol)blocking.push(label+'endpoint 프로토콜과 설계 프로토콜이 다릅니다.');
  }
  const protocols=endpoints.filter(Boolean).map(x=>protocol(x.protocol||l.protocol)).filter(Boolean);
  if(new Set(protocols).size>1)blocking.push(label+'양 끝 프로토콜이 다릅니다.');
  const linkSpeed=speedG(l.speed);
  if(linkSpeed===null)warnings.push(label+'링크 속도 표기가 미확정입니다.');
  else if(designSpeed!==null&&linkSpeed<designSpeed)blocking.push(label+'링크 속도가 설계 속도보다 낮습니다 ('+s(l.speed)+' / '+s(speed)+').');
  const sku=s(l.sku||l.productSku||l.productId),product=products.find(p=>p&&[p.sku,p.productSku,p.id].some(id=>s(id)===sku)&&sku)||{};
  if(!sku)warnings.push(label+'제품 SKU가 없어 제품 단위 호환성은 UNKNOWN입니다.');
  if(!evidenceUrl(l.datasheetUrl||l.officialUrl||product.datasheetUrl||product.officialUrl))warnings.push(label+'제품 공식 데이터시트 URL 근거가 없습니다.');
  const length=n(l.lengthM),reach=n(l.maxReachM??l.supportedReachM??l.reachM??product.maxReachM);
  if(length===null||length<0)warnings.push(label+'유효한 길이 정보가 없습니다.');
  if(reach===null||reach<0)warnings.push(label+'제품 reach 정보가 없어 거리 호환성을 확인할 수 없습니다.');
  else if(length!==null&&length>reach)blocking.push(label+'길이 '+length+' m가 제품 reach '+reach+' m를 초과합니다.');
  const media=s(l.media||l.fiber).toUpperCase(),productMedia=s(product.media||product.fiber).toUpperCase();
  if(!/^(SMF|MMF|OS2|OM3|OM4|OM5|SINGLE MODE|MULTIMODE|DAC|AEC|COPPER|AOC)$/.test(media))warnings.push(label+'매체 분류가 미확정입니다.');
  const family=v=>/SMF|OS2|SINGLE MODE/.test(v)?'smf':/MMF|OM[345]|MULTIMODE/.test(v)?'mmf':v;
  if(productMedia&&family(media)!==family(productMedia))blocking.push(label+'제품과 설계 매체가 다릅니다.');
  if(optical(l)){
   const connector=s(l.connectorType||l.connector||product.connectorType),fec=s(l.fec||product.fec);
   if(!connector)warnings.push(label+'connector 정보가 없어 접속 호환성이 미확정입니다.');
   if(!fec)warnings.push(label+'FEC 정보가 없어 송수신 호환성이 미확정입니다.');
   for(const ep of endpoints.filter(Boolean)){
    if(!s(ep.connectorType))warnings.push(label+'endpoint '+ep.id+' connector 정보가 미확정입니다.');
    else if(connector&&s(ep.connectorType).toLowerCase()!==connector.toLowerCase())blocking.push(label+'endpoint connector가 링크와 다릅니다.');
    if(!s(ep.fec))warnings.push(label+'endpoint '+ep.id+' FEC 정보가 미확정입니다.');
    else if(fec&&s(ep.fec).toLowerCase()!==fec.toLowerCase())blocking.push(label+'endpoint FEC가 링크와 다릅니다.');
   }
  }
  const productSpeed=speedG(product.speed||product.portSpeed);
  if(productSpeed!==null&&linkSpeed!==null&&linkSpeed>productSpeed)blocking.push(label+'링크 속도가 제품 지원 속도를 초과합니다.');
  const productProtocol=protocol(product.protocol);
  if(productProtocol&&designProtocol&&productProtocol!==designProtocol)blocking.push(label+'제품 프로토콜과 설계 프로토콜이 다릅니다.');
  if(mode==='cpo'&&/pluggable|transceiver/i.test(s(l.type))&&/(switch|spine|leaf|core)/i.test(endpoints.filter(Boolean).map(x=>x.layer+' '+x.name).join(' ')))blocking.push(label+'CPO switch side와 별도 pluggable 광모듈이 중복 계상될 수 있습니다.');
 }
 for(const [id,count] of used){const ep=nodes.find(x=>x.id===id),capacity=n(ep?.portCount);if(capacity===null||capacity<0||!Number.isInteger(capacity))warnings.push('노드 '+id+': 물리 port/cage 수용량이 미확정입니다.');else if(count>capacity)blocking.push('노드 '+id+': 사용 포트 '+count+'개가 수용량 '+capacity+'개를 초과합니다.');}
 const explicitPorts=n(top.portCapacity??top.availablePorts??o.assumptions?.portCount),activeOptical=active.filter(optical).length;
 if(explicitPorts!==null&&activeOptical>explicitPorts)blocking.push('활성 광 링크 '+activeOptical+'개가 입력 포트 수용량 '+explicitPorts+'개를 초과합니다.');
 if(!links.length)warnings.push('토폴로지 링크가 없어 광 연결 호환성을 확인할 수 없습니다.');
 info.push('PASS는 입력된 제품 근거·포트·접속 규칙 검사 결과입니다. URL 원문 및 실제 장비의 독립 검증·구매 승인은 별도입니다.');
 return finish();
}
return {validate,speedG};
});
