// STEP 07 — deterministic grid A* navigation for the Three.js digital twin.
// Original worker routes, simulation time, facility profiles, and emergency scenarios stay authoritative.
// This is a conceptual collision planner, not a life-safety egress solver or certified NavMesh.
export function createNpcNavigation(T, options={}) {
  const cell=options.cellSize || 1.25, radius=options.agentRadius || .39, dwell=2.3;
  const plans=new Map(), lastPose=new Map();
  const metrics={planned:0,failedLegs:0,expanded:0,queries:0,cacheHits:0,
    corrections:0,stops:0,frames:0,doorCrossings:0};
  const clamp=(v,lo,hi)=>Math.max(lo,Math.min(hi,v));
  const dist=(a,b)=>Math.hypot(a[0]-b[0],a[1]-b[1]);
  const eps=.0001;
  function snapshot(n) {
    const route=T.routeForWorker ? T.routeForWorker(n) : n.route;
    const points=(Array.isArray(route)?route:[]).filter(p=>Array.isArray(p)&&p.length>=2&&Number.isFinite(p[0])&&Number.isFinite(p[1]));
    const assets=(T.visibleAssets?.()||[]).filter(a=>a.zone===n.zone || (n.zone==='campus'&&a.zone==='campus'));
    const building=(T.Z||[]).find(z=>z.id===n.zone);
    const zones=n.zone==='campus' ? (T.visibleZones?.()||T.Z||[]).filter(z=>!['campus','utility','cooling'].includes(z.id)) : [];
    const mode=T.facilityMode || 'aidc';
    // Signature deliberately independent of currentZone so navigating to another zone
    // doesn't reset or teleport the route. It does react to source/asset geometry changes.
    const sig=JSON.stringify([mode,n.zone,points,
      assets.map(a=>[a.id,a.x,a.z,a.w,a.d,a.hidden]),
      zones.map(z=>[z.id,z.pos,z.size]),building?.pos,building?.size]);
    return {route:points,assets,zones,building,mode,sig};
  }
  function gridFor(n,snap) {
    const route=snap.route,b=snap.building;
    const bounds=n.zone==='campus'?[-69,69,-55,55]:
      (b?[b.pos[0]-b.size[0]/2-4,b.pos[0]+b.size[0]/2+4,b.pos[1]-b.size[1]/2-4,b.pos[1]+b.size[1]/2+4]:
        [-69,69,-55,55]);
    for(const p of route) {
      bounds[0]=Math.min(bounds[0],p[0]-3);bounds[1]=Math.max(bounds[1],p[0]+3);
      bounds[2]=Math.min(bounds[2],p[1]-3);bounds[3]=Math.max(bounds[3],p[1]+3);
    }
    const left=Math.floor(bounds[0]/cell)*cell,top=Math.floor(bounds[2]/cell)*cell;
    const nx=Math.max(3,Math.ceil((bounds[1]-left)/cell)+1),nz=Math.max(3,Math.ceil((bounds[3]-top)/cell)+1);
    const size=nx*nz,blocked=new Uint8Array(size);
    const inside=(x,z)=>x>=0&&z>=0&&x<nx&&z<nz;
    const index=(x,z)=>z*nx+x;
    const pos=(x,z)=>[left+x*cell,top+z*cell];
    const nearRect=(x,z,cx,cz,w,d,extra)=>Math.abs(x-cx)<w/2+extra&&Math.abs(z-cz)<d/2+extra;
    const building=snap.building;
    function blockedAt(x,z) {
      // World/site extents. In "server-room", outside the single hall is not navigable.
      if(snap.mode==='server-room'&&building){
        const w=26,d=24;
        if(Math.abs(x-building.pos[0])>w/2-radius || Math.abs(z-building.pos[1])>d/2-radius)return true;
      }
      for(const a of snap.assets){
        // Site marker/gate are not solid walls; other equipment footprints are.
        if(n.zone==='campus'&&/site-sign|gate/i.test(a.id||''))continue;
        if(nearRect(x,z,a.x,a.z,+a.w||2,+a.d||2,radius+.12))return true;
      }
      if(n.zone==='campus'){
        for(const z0 of snap.zones){
          if(nearRect(x,z,z0.pos[0],z0.pos[1],z0.size[0],z0.size[1],radius+.4))
            return true; // exterior paths do not cut through occupied buildings
        }
      }else if(building&&snap.mode!=='server-room'&&
        !['utility','cooling'].includes(n.zone)){
        const [cx,cz]=building.pos,[w,d]=building.size,half=w/2,dep=d/2;
        const outerX=Math.abs(x-cx),outerZ=Math.abs(z-cz);
        const inStrip=outerX>=half-(radius+.24)&&outerX<=half+(radius+.24);
        const inBack=outerZ>=dep-(radius+.24)&&outerZ<=dep+(radius+.24);
        const front=z>cz,portal=Math.abs(x-cx)<1.65-radius*.15;
        // Only the front entrance has a legal opening through the envelope.
        if((inStrip&&outerZ<=dep+radius+.24) ||
          (inBack&&outerX<=half+radius+.24&&!(front&&portal)))return true;
      }
      return false;
    }
    for(let iz=0;iz<nz;iz++)for(let ix=0;ix<nx;ix++){
      const [x,z]=pos(ix,iz);if(blockedAt(x,z))blocked[index(ix,iz)]=1;
    }
    function pointBlocked(p){return blockedAt(p[0],p[1]);}
    function nodeFor(p){
      const x0=clamp(Math.round((p[0]-left)/cell),0,nx-1),z0=clamp(Math.round((p[1]-top)/cell),0,nz-1);
      // Start/goal can be in apparatus; pick nearest safe point instead of walking through it.
      for(let r=0;r<Math.max(nx,nz);r++){
        let best=null,bestD=Infinity;
        for(let dz=-r;dz<=r;dz++)for(let dx=-r;dx<=r;dx++){
          if(Math.max(Math.abs(dx),Math.abs(dz))!==r)continue;
          const xx=x0+dx,zz=z0+dz;
          if(!inside(xx,zz)||blocked[index(xx,zz)])continue;
          const d=(xx-x0)**2+(zz-z0)**2;
          if(d<bestD){bestD=d;best=index(xx,zz);}
        }
        if(best!==null)return best;
      }
      return -1;
    }
    // Sample each piece of straight travel; never use a diagonal shortcut across the corner of a rack.
    function lineClear(a,b) {
      const len=dist(a,b),steps=Math.max(1,Math.ceil(len/(cell*.35)));
      for(let i=0;i<=steps;i++){
        const f=i/steps,x=a[0]+(b[0]-a[0])*f,z=a[1]+(b[1]-a[1])*f;
        if(blockedAt(x,z))return false;
      }
      return true;
    }
    return {nx,nz,size,blocked,index,pos,nodeFor,lineClear,pointBlocked,left,top};
  }
  class Heap {
    constructor(){this.items=[];}
    push(key,priority){
      const a=this.items;let i=a.length;a.push([key,priority]);
      while(i>0){const p=(i-1)>>1;if(a[p][1]<=priority)break;a[i]=a[p];i=p;}a[i]=[key,priority];
    }
    pop(){
      const a=this.items,head=a[0],last=a.pop();
      if(a.length){let i=0;while(true){let c=i*2+1;if(c>=a.length)break;
        if(c+1<a.length&&a[c+1][1]<a[c][1])c++;
        if(a[c][1]>=last[1])break;a[i]=a[c];i=c;}a[i]=last;}
      return head;
    }
    get length(){return this.items.length;}
  }
  function findPath(grid,a,b){
    metrics.queries++;
    const start=grid.nodeFor(a),goal=grid.nodeFor(b);
    if(start<0||goal<0)return null;
    if(start===goal)return [grid.pos(start%grid.nx,Math.floor(start/grid.nx))];
    const N=grid.size,cost=new Float64Array(N),prev=new Int32Array(N),closed=new Uint8Array(N);
    cost.fill(Infinity);prev.fill(-1);cost[start]=0;
    const h=i=>Math.hypot((i%grid.nx)-(goal%grid.nx),Math.floor(i/grid.nx)-Math.floor(goal/grid.nx));
    const heap=new Heap();heap.push(start,h(start));
    const dirs=[[-1,0],[1,0],[0,-1],[0,1],[-1,-1],[-1,1],[1,-1],[1,1]];
    let found=false,steps=0;
    while(heap.length && steps++<N*5){
      const [i]=heap.pop();if(closed[i])continue;closed[i]=1;metrics.expanded++;
      if(i===goal){found=true;break;}
      const x=i%grid.nx,z=(i/grid.nx)|0;
      for(const [dx,dz] of dirs){
        const xx=x+dx,zz=z+dz;
        if(xx<0||zz<0||xx>=grid.nx||zz>=grid.nz)continue;
        const ni=grid.index(xx,zz);
        if(grid.blocked[ni]||closed[ni])continue;
        if(dx&&dz&&(grid.blocked[grid.index(x+dx,z)]||grid.blocked[grid.index(x,z+dz)]))continue;
        const tentative=cost[i]+(dx&&dz?Math.SQRT2:1);
        if(tentative+1e-8<cost[ni]){
          cost[ni]=tentative;prev[ni]=i;heap.push(ni,tentative+h(ni));
        }
      }
    }
    if(!found)return null;
    const raw=[];for(let k=goal;k!==-1;k=prev[k]){
      raw.push(grid.pos(k%grid.nx,Math.floor(k/grid.nx)));if(k===start)break;
    }
    raw.reverse();
    // Prune intermediate waypoints with a real clearance-aware line test.
    const simplified=[raw[0]];let k=0;
    while(k<raw.length-1){
      let next=raw.length-1;while(next>k+1&&!grid.lineClear(raw[k],raw[next]))next--;
      simplified.push(raw[next]);k=next;
    }
    return simplified;
  }
  function legFor(grid,a,b){
    const source=grid.pointBlocked(a)?grid.posFromNode?.(grid.nodeFor(a)):a;
    const target=grid.pointBlocked(b)?grid.posFromNode?.(grid.nodeFor(b)):b;
    // The optional helpers may be absent: use grid coordinates for safe snapped endpoints.
    const safe=n=>n<0?null:grid.pos(n%grid.nx,(n/grid.nx)|0);
    const A=source||safe(grid.nodeFor(a)),B=target||safe(grid.nodeFor(b));
    if(!A||!B)return null;
    if(grid.lineClear(A,B))return [A,B];
    const route=findPath(grid,A,B);if(!route)return null;
    if(dist(route[0],A)>eps&&grid.lineClear(A,route[0]))route.unshift(A);
    if(dist(route[route.length-1],B)>eps&&grid.lineClear(route[route.length-1],B))route.push(B);
    return route;
  }
  function length(points) {
    let total=0;for(let i=1;i<points.length;i++)total+=dist(points[i-1],points[i]);return total;
  }
  function getPlan(n) {
    const snap=snapshot(n);
    const cached=plans.get(n.id);
    if(cached?.sig===snap.sig){metrics.cacheHits++;return cached;}
    const grid=gridFor(n,snap),legs=[];
    for(let i=1;i<snap.route.length;i++){
      const a=snap.route[i-1],b=snap.route[i];
      const points=legFor(grid,a,b);
      if(!points){metrics.failedLegs++;legs.push({points:[],distance:0,blocked:true});continue;}
      // Exactly one valid door opening is left in the enclosing façade.
      let crosses=false;
      if(snap.building&&n.zone!=='campus'){
        const front=snap.building.pos[1]+snap.building.size[1]/2;
        for(let j=1;j<points.length;j++)if((points[j-1][1]-front)*(points[j][1]-front)<=0)crosses=true;
      }
      if(crosses)metrics.doorCrossings++;
      legs.push({points,distance:length(points),blocked:false});
    }
    const plan={sig:snap.sig,grid,legs,route:snap.route,mode:snap.mode,zone:n.zone,
      total:legs.reduce((sum,l)=>sum+l.distance+dwell,0),
      fails:legs.filter(l=>l.blocked).length};
    plans.set(n.id,plan);metrics.planned++;lastPose.delete(n.id);return plan;
  }
  function poseFromPlan(plan,n,simTime,original) {
    if(!plan.legs.length)return original;
    const rate=Math.max(.05,(Number(n.speed)||2.5)/8);
    const cycle=Math.max(eps,plan.total);
    let clock=((Math.max(0,simTime)*rate%cycle)+cycle)%cycle;
    for(let i=0;i<plan.legs.length;i++){
      const leg=plan.legs[i];
      if(leg.blocked||leg.points.length<2)continue;
      if(clock<leg.distance){
        let remaining=clock;
        for(let j=1;j<leg.points.length;j++){
          const a=leg.points[j-1],b=leg.points[j],d=dist(a,b);
          if(remaining<=d){
            const f=d?remaining/d:1;
            return {x:a[0]+(b[0]-a[0])*f,z:a[1]+(b[1]-a[1])*f,
              heading:Math.atan2(b[0]-a[0],b[1]-a[1]),action:'walk'};
          }
          remaining-=d;
        }
      }
      clock-=leg.distance;
      const end=leg.points[leg.points.length-1],start=leg.points[leg.points.length-2];
      if(clock<dwell){
        return {x:end[0],z:end[1],heading:Math.atan2(end[0]-start[0],end[1]-start[1]),
          action:i%2?'inspect':'pause'};
      }
      clock-=dwell;
    }
    return original;
  }
  function avoidAgents(n,proposed,plan,simTime){
    if(!plan.grid)return proposed;
    const p=[proposed.x,proposed.z],prev=lastPose.get(n.id);
    const others=[];
    for(const p0 of T.npcs||[]){
      if(p0.id===n.id)continue;
      if(p0.point && (p0.zone===n.zone || n.zone==='campus'))others.push(p0.point);
      const old=lastPose.get(p0.id);
      if(old && old.mode===plan.mode && old.zone===n.zone && Math.abs(old.time-simTime)<4)
        others.push([old.pose.x,old.pose.z]);
    }
    const safe=q=>!plan.grid.pointBlocked(q)&&
      others.every(other=>dist(q,other)>=radius*2.4);
    if(safe(p))return proposed;
    const dirs=[0,1,-1,2,-2,3,-3,4].map(i=>proposed.heading+i*Math.PI/4);
    for(let r=radius*2.6;r<=radius*5;r+=radius*.8)for(const ang of dirs){
      const q=[p[0]+Math.cos(ang)*r,p[1]+Math.sin(ang)*r];
      if(!safe(q))continue;
      if(!plan.grid.lineClear(p,q) && (prev && !plan.grid.lineClear([prev.pose.x,prev.pose.z],q)))continue;
      metrics.corrections++;
      return {...proposed,x:q[0],z:q[1],heading:Math.atan2(q[0]-(prev?.pose.x??p[0]),q[1]-(prev?.pose.z??p[1]))};
    }
    metrics.stops++;
    if(prev && prev.mode===plan.mode && safe([prev.pose.x,prev.pose.z]))
      return {...prev.pose,action:'pause'};
    return {...proposed,action:'pause'};
  }
  function pose(n,original,simTime=T.simTime||0) {
    if(!n||!original||!Number.isFinite(original.x)||!Number.isFinite(original.z))return original;
    if(n.point)return original; // guards remain at their original posts
    const plan=getPlan(n);
    if(!plan.legs.length || plan.fails)return original;
    let candidate=poseFromPlan(plan,n,simTime,original);
    candidate=avoidAgents(n,candidate,plan,simTime);
    const previous=lastPose.get(n.id),speed=Math.max(.05,(Number(n.speed)||2.5)/8);
    if(previous&&previous.mode===plan.mode&&simTime>=previous.time&&simTime-previous.time<2){
      // Clamp the maximum advance to physical walk speed, preserving pause and frame continuity.
      const a=[previous.pose.x,previous.pose.z],b=[candidate.x,candidate.z],distance=dist(a,b);
      const maxStep=Math.max(.055,speed*(simTime-previous.time)*1.5+.055);
      if(distance>maxStep&&plan.grid.lineClear(a,b)){
        const f=maxStep/distance;
        candidate={...candidate,x:a[0]+(b[0]-a[0])*f,z:a[1]+(b[1]-a[1])*f,
          heading:Math.atan2(b[0]-a[0],b[1]-a[1]),action:'walk'};
      }
    }
    lastPose.set(n.id,{mode:plan.mode,zone:n.zone,time:simTime,pose:candidate});
    metrics.frames++;
    return candidate;
  }
  function diagnostic(){
    return {cellSize:cell,agentRadius:radius,plans:plans.size,metrics:{...metrics},
      routes:[...plans.entries()].map(([id,p])=>({id,zone:p.zone,legs:p.legs.length,
        failed:p.fails,routeLength:Number(p.total.toFixed(2))}))};
  }
  function reset(){plans.clear();lastPose.clear();}
  return {pose,getPlan,diagnostic,reset};
}
