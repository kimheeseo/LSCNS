/* Phase 7-B · evidence-aware compatibility checks; missing product data stays unknown. */
(function(root,factory){
  const api=factory();
  if(typeof module==='object'&&module.exports)module.exports=api;
  if(root)root.LSDCPhase7Compatibility=api;
})(typeof globalThis!=='undefined'?globalThis:this,function(){
  'use strict';
  const s=v=>String(v??'').trim(),n=v=>Number.isFinite(Number(v))?Number(v):null;
  const optical=l=>!/\b(dac|aec|copper)\b/i.test(s(l.type)+' '+s(l.media||l.fiber));
  function validate(packet){
    const blocking=[],warnings=[],info=[];
    if(!packet||packet.schemaVersion!=='lsdc-twin-bom/1.0')return {compatible:false,blocking:['지원 가능한 lsdc-twin-bom/1.0 packet이 아닙니다.'],warnings:[],info:[]};
    const f=packet.facility||{},o=packet.optical||{},top=o.topology||{},links=Array.isArray(top.links)?top.links:[],nodes=Array.isArray(top.nodes)?top.nodes:[];
    const protocol=s(packet.network?.protocol||top.protocol||o.protocol||o.assumptions?.protocol);
    const mode=s(o.cpoMode).toLowerCase(),speed=s(top.portSpeed||o.assumptions?.portSpeed||links.find(x=>x.speed)?.speed);
    if(!protocol)warnings.push('네트워크 프로토콜이 packet에 없어 Ethernet/InfiniBand 조합을 확정할 수 없습니다.');
    if(protocol&&!/^(ethernet|infiniband|ib)$/i.test(protocol))warnings.push('알 수 없는 프로토콜 값: '+protocol);
    for(const l of links.filter(x=>!x.down)){
      const ep=[l.from,l.to].map(id=>nodes.find(x=>x.id===id)).filter(Boolean);
      const epProtocols=ep.map(x=>s(x.protocol)).filter(Boolean);
      if(protocol&&epProtocols.some(x=>x.toLowerCase()!==protocol.toLowerCase()&&!(protocol.toLowerCase()==='infiniband'&&x.toLowerCase()==='ib')))blocking.push('링크 '+l.id+': endpoint 프로토콜과 설계 프로토콜이 다릅니다.');
      const linkSpeed=n(String(l.speed||'').match(/[\d.]+/)?.[0]),designSpeed=n(String(speed).match(/[\d.]+/)?.[0]);
      const scale=(v)=>/t/i.test(v)?1000:1;
      if(linkSpeed!==null&&designSpeed!==null&&linkSpeed*scale(s(l.speed))<designSpeed*scale(speed))blocking.push('링크 '+l.id+': 링크 속도가 설계 속도보다 낮습니다 ('+l.speed+' / '+speed+').');
      const sku=s(l.sku||l.productSku||l.productId);
      const maxReach=n(l.maxReachM??l.supportedReachM??l.reachM);
      const length=n(l.lengthM);
      if(length!==null&&maxReach!==null&&length>maxReach)blocking.push('링크 '+l.id+': 길이 '+length+' m가 입력된 제품 reach '+maxReach+' m를 초과합니다.');
      else if(length!==null&&maxReach===null)warnings.push('링크 '+l.id+': '+(sku?'제품 SKU의 공식 reach 정보 확인 필요':'SKU/reach 정보가 없어 거리 호환성을 확인할 수 없습니다.')+' ('+length+' m)');
      const media=s(l.media||l.fiber).toUpperCase();
      if(media&&!['SMF','MMF','SINGLE MODE','MULTIMODE','OS2','OM3','OM4','OM5'].some(x=>media.includes(x))&&!/DAC|AEC|COPPER|AOC/i.test(media))warnings.push('링크 '+l.id+': media 값이 표준 SMF/MMF/DAC/AEC/AOC 분류로 확인되지 않습니다.');
      if(mode==='cpo'&&/pluggable|transceiver/i.test(s(l.type))&&/(switch|spine|leaf|core)/i.test(ep.map(x=>x.layer+' '+x.name).join(' ')))blocking.push('링크 '+l.id+': CPO switch side와 별도 pluggable 광모듈이 중복 계상될 수 있습니다.');
    }
    if(!mode)warnings.push('CPO/Pluggable 모드가 없어 광 모듈 계수 방식을 확인할 수 없습니다.');
    else if(!['cpo','pluggable'].includes(mode))warnings.push('알 수 없는 광 모드: '+mode);
    else info.push(mode==='cpo'?'CPO 설정: switch-side 내장 광 엔진과 별도 pluggable 수량을 분리 검토합니다.':'Pluggable 설정: 양 끝 광모듈 수량은 port/SKU 정보가 없어 미확정입니다.');
    if(!speed)warnings.push('포트 속도 정보가 없어 400G/800G/1.6T 정합성을 확정할 수 없습니다.');
    else if(!/^(400|800|1600)\s*(g|gbps|gb\/s|t|tbps|tb\/s)?$/i.test(speed.replace(/\s/g,'')))warnings.push('지원 여부를 판정하지 못한 링크 속도 표기: '+speed);
    const active=links.filter(x=>!x.down&&optical(x)).length,explicitPorts=n(top.portCapacity??top.availablePorts??o.assumptions?.portCount);
    if(explicitPorts!==null&&active>explicitPorts)blocking.push('활성 광 링크 '+active+'개가 입력 포트 수용량 '+explicitPorts+'개를 초과합니다.');
    else if(active&&!nodes.some(x=>n(x.portCount)!==null))warnings.push('스위치별 물리 cage/port 수가 없어 링크와 포트 수용량을 비교할 수 없습니다.');
    if(!links.length)warnings.push('토폴로지 링크가 없어 광 연결 호환성을 확인할 수 없습니다.');
    if(!packet.products&&!links.some(x=>x.sku||x.productSku||x.productId))warnings.push('제품 SKU 및 공식 데이터시트 연결이 없어 제품 단위 호환성은 UNKNOWN입니다.');
    return {compatible:blocking.length===0,blocking,warnings,info,status:blocking.length?'FAIL':warnings.length?'UNKNOWN':'PASS',profile:f.profile||null,activeLinks:links.filter(x=>!x.down).length};
  }
  return {validate};
});
