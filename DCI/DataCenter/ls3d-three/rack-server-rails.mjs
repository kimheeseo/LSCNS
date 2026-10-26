// STEP 12 — separately sliding individual server sleds, guide rails and visible rear NIC/ports.
// Changes visual scene only; GPU/BOM/NIC electrical capacity are not recalculated.
export function createRackServerRails(THREE,{cube}) {
  const sessions=new Map(),stats={enhanced:0,sleds:0,nics:0,ports:0,updates:0};
  function box(group,name,x,y,z,w,h,d,color){
    const v=cube(group,x,y,z,w,h,d,color,.48,.41);v.name='rack-server:'+name;return v;
  }
  function decorate(asset,built,high){
    const id=asset.id,w=asset.w,d=asset.d,h=asset.h,serverGroups=built.serverGroups||[];
    const groups=new Map();
    const unit=(h-.46)/built.rackUnits;
    for(const group of serverGroups){
      const info=group.userData.rackPart;if(!info)continue;
      groups.set(info.id,group);
      const y=.23+(info.uStart-1)*unit,unitH=info.uHeight*unit;
      // Rails remain attached to the frame; only the selected sled moves.
      for(const side of [-1,1]){
        box(high,'fixed-guide-'+info.id+'-'+side,side*w*.365,y+.06,-d*.13,w*.044,.052,d*.76,'#b4c6d1');
        box(group,'sled-rail-'+info.id+'-'+side,side*w*.338,y+.07,-d*.11,w*.035,.047,d*.67,'#7898ac');
        box(group,'rail-carriage-'+info.id+'-'+side,side*w*.35,y+.09,d*.14,w*.064,.06,d*.08,'#dfba75');
      }
      // Rear NIC plate: one physical NIC card + independent 4 duplex port shells.
      box(group,'nic-board-'+info.id,0,y+unitH*.21,d*.16,w*.52,.057,d*.25,'#2b8b78');
      const nic=box(group,'nic-frontplate-'+info.id,0,y+unitH*.22,d*.33,w*.51,.23,.075,'#8aa7bd');
      for(let port=0;port<4;port++){
        const xx=-w*.18+port*w*.12;
        box(group,'nic-cage-'+info.id+'-'+port,xx,y+unitH*.26,d*.382,w*.082,.095,.065,'#2a4c69');
        box(group,'nic-latch-'+info.id+'-'+port,xx,y+unitH*.33,d*.396,w*.054,.028,.06,'#ecc784');
        stats.ports++;
      }
      const led=box(group,'server-health-'+info.id,w*.25,y+unitH*.36,-d*.547,.10,.064,.025,'#65e4b9');
      for(const mesh of [nic,led]){mesh.userData.assetId=id;mesh.userData.partId=info.id;}
      stats.nics++;stats.sleds++;
    }
    const gpuGroups=new Map(),gpuOrder=[];
    for(const gpu of built.gpuGroups||[]){
      const partId=gpu.userData.rackPart?.id;
      if(!partId)continue;
      gpuGroups.set(partId,{group:gpu,serverPartId:gpu.userData.serverPartId,target:0,progress:0});
      gpuOrder.push(partId);
    }
    const driveGroups=new Map(),driveOrder=[];
    for(const drive of built.storageDrives||[]){
      const partId=drive.userData.rackPart?.id;
      if(!partId)continue;
      driveGroups.set(partId,{group:drive,shelfPartId:drive.userData.storageShelfId,target:0,progress:0});
      driveOrder.push(partId);
    }
    const session={id,groups,selected:groups.keys().next().value||null,current:0,
      gpuGroups,gpuOrder,selectedGpu:gpuOrder[0]||null,
      driveGroups,driveOrder,selectedDrive:driveOrder[0]||null,depth:d};
    sessions.set(id,session);stats.enhanced++;
    return session;
  }
  function select(assetId,partId){
    const s=sessions.get(assetId);if(!s||!s.groups.has(partId))return null;
    s.selected=partId;
    const first=s.gpuOrder.find(id=>s.gpuGroups.get(id).serverPartId===partId);
    if(first&&s.gpuGroups.get(s.selectedGpu)?.serverPartId!==partId)s.selectedGpu=first;
    const firstDrive=s.driveOrder.find(id=>s.driveGroups.get(id).shelfPartId===partId);
    if(firstDrive&&s.driveGroups.get(s.selectedDrive)?.shelfPartId!==partId)s.selectedDrive=firstDrive;
    return partId;
  }
  function selectGpu(assetId,partId){
    const s=sessions.get(assetId),gpu=s?.gpuGroups.get(partId);
    if(!gpu)return null;
    for(const [id,other] of s.gpuGroups)if(id!==partId)other.target=0;
    s.selected=gpu.serverPartId;s.selectedGpu=partId;
    return {id:partId,serverPartId:gpu.serverPartId,slot:s.gpuOrder.indexOf(partId)+1,total:s.gpuOrder.length};
  }
  function cycleGpu(assetId,step=1){
    const s=sessions.get(assetId);if(!s?.gpuOrder.length)return null;
    const index=s.gpuOrder.indexOf(s.selectedGpu);
    return selectGpu(assetId,s.gpuOrder[(index+step+s.gpuOrder.length)%s.gpuOrder.length]);
  }
  function toggleGpu(assetId){
    const s=sessions.get(assetId),gpu=s?.gpuGroups.get(s.selectedGpu);
    if(!gpu)return null;
    gpu.target=gpu.target?0:1;
    return {id:s.selectedGpu,target:gpu.target};
  }
  function selectDrive(assetId,partId){
    const s=sessions.get(assetId),drive=s?.driveGroups.get(partId);
    if(!drive)return null;
    for(const [id,other] of s.driveGroups)if(id!==partId)other.target=0;
    s.selected=drive.shelfPartId;s.selectedDrive=partId;
    return {id:partId,shelfPartId:drive.shelfPartId,slot:s.driveOrder.indexOf(partId)+1,total:s.driveOrder.length};
  }
  function cycleDrive(assetId,step=1){
    const s=sessions.get(assetId);if(!s?.driveOrder.length)return null;
    const index=s.driveOrder.indexOf(s.selectedDrive);
    return selectDrive(assetId,s.driveOrder[(index+step+s.driveOrder.length)%s.driveOrder.length]);
  }
  function toggleDrive(assetId){
    const s=sessions.get(assetId),drive=s?.driveGroups.get(s.selectedDrive);
    if(!drive)return null;
    drive.target=drive.target?0:1;
    return {id:s.selectedDrive,target:drive.target};
  }
  function update(assetId,moving,legacy,dt=.016) {
    const session=sessions.get(assetId);
    if(!session)return null;
    const pull=Math.max(0,Math.min(1,Number(legacy?.slide)||0))*(moving?.reach||1);
    // Do not move every server: keep the original shared parent stationary.
    if(moving?.slide)moving.slide.position.z=0;
    for(const [key,group] of session.groups)group.position.z=key===session.selected?-pull:0;
    const amount=Math.min(1,Math.max(0,dt)*7);
    for(const gpu of session.gpuGroups.values()){
      gpu.progress+=(gpu.target-gpu.progress)*amount;
      gpu.group.position.z=gpu.progress?-gpu.progress*session.depth*.56:0;
    }
    for(const drive of session.driveGroups.values()){
      drive.progress+=(drive.target-drive.progress)*amount;
      drive.group.position.z=drive.progress?-drive.progress*session.depth*.32:0;
    }
    session.current=pull;stats.updates++;
    return {assetId,selected:session.selected,pull,serverCount:session.groups.size};
  }
  function state(id){
    const s=sessions.get(id);
    return s?{selected:s.selected,pull:s.current,serverCount:s.groups.size,
      selectedGpu:s.selectedGpu,gpuCount:s.gpuOrder.length,
      gpuSlot:s.gpuOrder.indexOf(s.selectedGpu)+1,
      gpuTarget:s.gpuGroups.get(s.selectedGpu)?.target||0,
      gpuProgress:s.gpuGroups.get(s.selectedGpu)?.progress||0,
      selectedDrive:s.selectedDrive,driveCount:s.driveOrder.length,
      driveSlot:s.driveOrder.indexOf(s.selectedDrive)+1,
      driveTarget:s.driveGroups.get(s.selectedDrive)?.target||0}:null;
  }
  function drop(id){sessions.delete(id);}
  function reset(){sessions.clear();}
  return {decorate,select,selectGpu,cycleGpu,toggleGpu,selectDrive,cycleDrive,toggleDrive,update,state,drop,reset,stats};
}
