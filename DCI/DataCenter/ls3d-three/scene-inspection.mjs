// Shared product descriptions and visibility rules for the real scene picker.
export function isVisiblePickObject(object){
  for(let node=object;node;node=node.parent)if(node.visible===false)return false;
  return !object.userData?.noPick&&object.name!=='selection-ring';
}

export function describeEquipment(asset,kind){
  if(!asset||kind==='rack')return null;
  const busway=kind==='busway';
  return {
    id:'equipment:'+asset.id,assetId:asset.id,zoneId:asset.zone,kind:'equipment',
    label:busway?'부스덕트 · 버스웨이 (Busduct / Busway)':asset.name,
    description:busway?
      '금속 하우징 안의 도체로 전력을 전달·분배하는 설비입니다. 전력 계통에서 공급받은 전기를 탭오프 분기 장치를 통해 하위 분전 설비나 랙 전원으로 보냅니다. 화면의 금속 외함은 보호 하우징, 내부 금색 부재는 도체를 나타냅니다. 케이블을 받치는 트레이와 달리 내부 도체가 전력을 전달합니다. 이 모델은 교육용 개념도이며 LS전선의 특정 제품·정격전류·단락내량을 지정하지 않습니다.':
      [asset.desc,asset.details,'교육용 개념 모델이며 실제 제품 사양과 배치는 별도 확인이 필요합니다.'].filter(Boolean).join(' ')
  };
}
