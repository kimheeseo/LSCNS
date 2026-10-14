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
    const session={id,groups,selected:groups.keys().next().value||null,current:0};
    sessions.set(id,session);stats.enhanced++;
    return session;
  }
  function select(assetId,partId){
    const s=sessions.get(assetId);if(!s||!s.groups.has(partId))return null;
    s.selected=partId;return partId;
  }
  function update(assetId,moving,legacy) {
    const session=sessions.get(assetId);
    if(!session)return null;
    const pull=Math.max(0,Math.min(1,Number(legacy?.slide)||0))*(moving?.reach||1);
    // Do not move every server: keep the original shared parent stationary.
    if(moving?.slide)moving.slide.position.z=0;
    for(const [key,group] of session.groups)group.position.z=key===session.selected?-pull:0;
    session.current=pull;stats.updates++;
    return {assetId,selected:session.selected,pull,serverCount:session.groups.size};
  }
  function state(id){
    const s=sessions.get(id);
    return s?{selected:s.selected,pull:s.current,serverCount:s.groups.size}:null;
  }
  function drop(id){sessions.delete(id);}
  function reset(){sessions.clear();}
  return {decorate,select,update,state,drop,reset,stats};
}
