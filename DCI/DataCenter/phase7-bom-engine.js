/* Phase 7-A · shared summary for lsdc-twin-bom/1.0 packets. */
(function(root,factory){
  const api=factory();
  if(typeof module==='object'&&module.exports)module.exports=api;
  if(root)root.LSDCPhase7BOM=api;
})(typeof globalThis!=='undefined'?globalThis:this,function(){
  'use strict';
  const finite=(v)=>Number.isFinite(Number(v))?Number(v):null;
  const count=(v)=>{const n=finite(v);return n===null?null:Math.max(0,Math.round(n))};
  const norm=(v)=>String(v??'').trim().toLowerCase();
  function calculate(packet){
    if(!packet||packet.schemaVersion!=='lsdc-twin-bom/1.0')throw new TypeError('Expected lsdc-twin-bom/1.0 packet');
    const f=packet.facility||{},opt=packet.optical||{},top=opt.topology||{},ass=opt.assumptions||{},links=Array.isArray(top.links)?top.links:[],nodes=Array.isArray(top.nodes)?top.nodes:[];
    const compute=packet.compute||packet.network?.compute||{},network=packet.network||{},rackCount=count(f.rackCount);
    const ratio=finite(compute.gpuServerRatio??f.gpuServerRatio);
    const gpuServers=count(compute.gpuServerCount??compute.serverCount??(rackCount!==null&&ratio!==null?rackCount*ratio:null));
    const gpuPerServer=count(compute.gpuPerServer??compute.acceleratorsPerServer??ass.gpuPerServer??8);
    const gpuCount=count(compute.gpuCount??compute.acceleratorCount??(gpuServers!==null?gpuServers*gpuPerServer:null));
    const nicPerServer=count(network.nicPerServer??compute.nicPerServer??ass.nicPerServer??1);
    const nicCount=count(network.nicCount??compute.nicCount??(gpuServers!==null?gpuServers*nicPerServer:null));
    const layer=(n)=>norm(n.layer??n.type??n.role);
    const layerCount=(names)=>nodes.filter(n=>names.some(x=>layer(n)===x||layer(n).includes(x))).length;
    const leafCount=count(network.leafCount??layerCount(['leaf','tor']));
    const spineCount=count(network.spineCount??layerCount(['spine']));
    const coreCount=count(network.coreCount??layerCount(['core']));
    const opticalLinks=links.filter(l=>!l.down&&!/\b(dac|aec|copper)\b/i.test(String(l.type||'')+' '+String(l.media||l.fiber||'')));
    const opticalLinkCount=opticalLinks.length;
    const totalOpticalLengthM=opticalLinks.reduce((s,l)=>s+(finite(l.lengthM)??0),0);
    const gpuAssumption=(compute.gpuCount==null&&compute.acceleratorCount==null)?'기본 가정: GPU 서버당 8개':'입력 packet 명시';
    const nicAssumption=(network.nicCount==null&&compute.nicCount==null)?'기본 가정: GPU 서버당 NIC 1개':'입력 packet 명시';
    const rows=[
      {id:'rack',category:'Facility',part:'서버 랙',quantity:rackCount,unit:'대',basis:'facility.rackCount',source:'input'},
      {id:'gpu-server',category:'Compute',part:'GPU 서버',quantity:gpuServers,unit:'대',basis:ratio===null?'명시 수량 또는 GPU 서버 비율 필요':'rackCount × gpuServerRatio',source:ratio===null?'unknown':'input'},
      {id:'gpu',category:'Compute',part:'GPU / 가속기',quantity:gpuCount,unit:'개',basis:gpuAssumption,source:gpuAssumption.startsWith('기본')?'assumption':'input'},
      {id:'nic',category:'Network',part:'NIC',quantity:nicCount,unit:'개',basis:nicAssumption,source:nicAssumption.startsWith('기본')?'assumption':'input'},
      {id:'leaf',category:'Network',part:'Leaf / ToR 스위치',quantity:leafCount,unit:'대',basis:'optical.topology.nodes의 layer/type',source:'topology'},
      {id:'spine',category:'Network',part:'Spine 스위치',quantity:spineCount,unit:'대',basis:'optical.topology.nodes의 layer/type',source:'topology'},
      {id:'core',category:'Network',part:'Core 스위치',quantity:coreCount,unit:'대',basis:'optical.topology.nodes의 layer/type',source:'topology'},
      {id:'optical-link',category:'Optical',part:'활성 광 링크',quantity:opticalLinkCount,unit:'링크',basis:'활성 topology.links 중 DAC/AEC/Copper 제외',source:'topology'},
      {id:'optical-length',category:'Optical',part:'광 링크 총 길이',quantity:Number(totalOpticalLengthM.toFixed(2)),unit:'m',basis:'활성 광 링크 lengthM 합계 · 여유 길이 미포함',source:'topology'}
    ];
    return {schema:'lsdc/phase7-bom-engine/1.0',inputSchema:packet.schemaVersion,revision:packet.revision??packet.createdAt??null,counts:{rackCount,gpuServerCount:gpuServers,gpuCount,nicCount,leafCount,spineCount,coreCount,opticalLinkCount,totalOpticalLengthM:Number(totalOpticalLengthM.toFixed(2))},rows,assumptions:[gpuAssumption,nicAssumption,'시각화 객체·SKU 수량·광 모듈/코어 수량은 이 요약에서 별도로 확정하지 않음.']};
  }
  return {calculate};
});
