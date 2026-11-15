// Present the actual movable device in each rack, rather than a generic optical drawer.
export function mechanicalButtons(type,m={},rails={}){
 const door={code:'front-door',label:m.targetDoor?'랙 문 닫기':'랙 문 열기',pressed:!!m.targetDoor};
 const slide=noun=>({code:'server-slide',label:noun+(m.targetSlide?' 넣기':' 빼기'),pressed:!!m.targetSlide});
 if(/storage/i.test(type))return [door,slide('스토리지 쉘프'),
  {code:'drive-toggle',label:rails.driveTarget?'드라이브 넣기':'드라이브 빼기',pressed:!!rails.driveTarget},
  {code:'next-drive',label:'다음 드라이브 선택'}];
 if(/gpu|compute|server|co-location/i.test(type))return [door,slide('서버 트레이'),
  {code:'previous-server',label:'이전 서버 선택'},{code:'next-server',label:'다음 서버 선택'}];
 return [door,{code:'cassette-slide',label:m.targetCassette?'광 카세트 넣기':'광 카세트 빼기',pressed:!!m.targetCassette}];
}
export function pinnedMechanicalHeight(rows=1){return Number(rows)>1?'184px':'132px';}
