(function applySeptember20Mats(){
'use strict';
if(window.__hcfDay20Mats)return;
const day='2026-09-20',revision='v28.2';
const copy=x=>JSON.parse(JSON.stringify(x));
const isSunday=()=>currentStaffDate()===day;
const chairs=copy(scene.objects.filter(o=>o.kind==='chair'));
const zones=copy(scene.seat_zones),groups=copy(scene.mat_groups);
const countKeys=['chairs','temporary_chairs','mats','mat_groups','vvip_chairs','reserved_chairs','vvip','row_counts_by_zone'];
const savedCounts=Object.fromEntries(countKeys.map(k=>[k,copy(scene.counts[k]??null)]));
const center=p=>p.reduce((a,q)=>[a[0]+q[0]/p.length,a[1]+q[1]/p.length],[0,0]);
const rectangle=(x,y,w,h)=>[[x,y],[x+w,y],[x+w,y+h],[x,y+h]];
const bounds=p=>({x:Math.min(...p.map(q=>q[0])),y:Math.min(...p.map(q=>q[1])),right:Math.max(...p.map(q=>q[0])),bottom:Math.max(...p.map(q=>q[1]))});
const newMats=[],footprints=[];
// Build only inside the former chair blocks. Row breaks and the tower opening remain clear.
for(const zone of zones){
 const rowMap=new Map();
 for(const o of chairs.filter(o=>(o.zone_id||o.block)===zone.id)){
  const b=bounds(o.poly),key=b.y.toFixed(3);if(!rowMap.has(key))rowMap.set(key,[]);rowMap.get(key).push(b);
 }
 const rows=[...rowMap.values()].sort((a,b)=>a[0].y-b[0].y),bands=[];
 rows.forEach((row,i)=>{
  row.sort((a,b)=>a.x-b.x);const runs=[];
  for(const b of row){const last=runs[runs.length-1];if(last&&b.x-last.right<.25){last.right=Math.max(last.right,b.right);last.bottom=Math.max(last.bottom,b.bottom);}else runs.push({...b});}
  const next=rows[i+1]?.[0]?.y;
  for(const run of runs){if(next&&next-run.y<1.15)run.bottom=next;
   const last=bands.find(b=>Math.abs(b.x-run.x)<.001&&Math.abs(b.right-run.right)<.001&&Math.abs(b.bottom-run.y)<.001);
   if(last)last.bottom=run.bottom;else bands.push(run);
  }
 });
 bands.forEach((b,index)=>{
  const poly=rectangle(b.x,b.y,b.right-b.x,b.bottom-b.y);
  footprints.push({id:'DAY20-'+zone.id+'-'+index,zone_id:zone.id,poly});
  // Texture subdivisions are drawing units, not a confirmed mat quantity or audience capacity.
  const nx=Math.max(1,Math.round((b.right-b.x)/2)),ny=Math.max(1,Math.round((b.bottom-b.y)/2));
  const w=(b.right-b.x)/nx,h=(b.bottom-b.y)/ny;
  for(let r=0;r<ny;r++)for(let c=0;c<nx;c++)newMats.push({id:'DAY20-MAT-'+zone.id+'-'+index+'-'+r+'-'+c,kind:'mat',name:zone.id+' 돗자리 관람구역',poly:rectangle(b.x+c*w+.02,b.y+r*h+.02,w-.04,h-.04),height:.035,width:w-.04,depth:h-.04,color:(c+r)%2?'#81b573':'#8abd7e',group:'DAY20-'+zone.id,zone_id:zone.id,day20_mat:true,dimension_status:'돗자리 구역 표현용 분할. 실제 규격·수량·수용인원 산정값 아님.',layout_rotated:true});
 });
}
const chairToMat=t=>String(t||'').replaceAll('에어바운스','워터슬라이드').replaceAll('관람석','돗자리 관람구역').replaceAll('좌석 구역','돗자리 구역').replaceAll('좌석','돗자리 구역').replaceAll('복도','통로');
const newF07={u:15.8,v:6.5,location_label:'메인무대 우측 옆',role:'무대 옆 진행팀 연락·공연 전환 및 출연진 이동 지원·무대 주변 통행로 확보·혼잡 및 요청사항을 현장 총괄에게 전달'};
const flow='통로에 멈추거나 모여서 관람하지 않도록 이동 방향 안내·계속 통행하도록 유도·혼잡 시 현장 총괄에게 보고';
const beforeRoster=applyStaffRoster;
applyStaffRoster=function(){
 beforeRoster();if(!isSunday())return;
 for(const phase of ['ceremony','concert'])for(const p of scene.staff_phases[phase]||[]){
  p.location_label=chairToMat(p.location_label);p.role=chairToMat(p.role);
  if(p.id==='F07'){Object.assign(p,newF07,{remote:false,group:'towers'});delete p.external_area;}
  if(p.id==='F12')p.role='박경득(F22)과 구간 분담·'+flow;
  if(p.id==='F22')p.role='한상일(F12) 옆 현장 운영 지원·'+flow;
  if(p.sunday_post)p.sunday_post={...p.sunday_post,u:p.u,v:p.v,location_label:p.location_label,role:p.role};
  p.assignment=p.role;
 }
};
const beforeWorld=createWorld;
let lastDate=null,oldMatsChecked=true;
createWorld=function(s){
 const sunday=isSunday();
 const keep=s.objects.filter(o=>o.kind!=='chair'&&!o.day20_mat&&(!sunday||!String(o.kind).includes('chair')));
 s.objects=keep.concat(copy(sunday?newMats:chairs));
 s.seat_zones=copy(zones).map(z=>sunday?{...z,count:0,vip:0,vvip:0,seating_type:'mat'}:z);
 s.mat_groups=copy(groups).concat(sunday?copy(footprints):[]);
 for(const k of countKeys)if(savedCounts[k]!==null)s.counts[k]=copy(savedCounts[k]);
 if(sunday)Object.assign(s.counts,{chairs:0,temporary_chairs:0,candidate_temporary_chairs:0,vvip_chairs:0,reserved_chairs:0,vvip:0,row_counts_by_zone:{},mats:s.objects.filter(o=>o.kind==='mat').length,mat_groups:s.mat_groups.length});
 if(s.installation_seating)s.installation_seating.rendered_chair_count=sunday?0:chairs.length+(s.counts.temporary_chairs||0);
 const actualDate=currentStaffDate();if(lastDate!==actualDate){const m=$('mats');if(m){if(sunday){oldMatsChecked=m.checked;m.checked=true;}else if(lastDate===day)m.checked=oldMatsChecked;}lastDate=actualDate;}
 s.day20_seating={active:sunday,type:sunday?'전 구역 돗자리':'기존 배치',actual_mat_quantity:null,capacity:null};
 return beforeWorld(s);
};
const beforeGround=drawZoneGrounds;
drawZoneGrounds=function(){if(!isSunday())return beforeGround();ctx.save();for(const g of footprints)dashedPoly(g.poly,'#65856a',.7);ctx.restore();};
const beforeLabels=drawZoneCounts;
drawZoneCounts=function(){
 if(!isSunday())return beforeLabels();if(!$('labels').checked||!['all','seats'].includes(focus))return;
 for(const z of zones){if(canvas.clientWidth<700&&z.id==='A1 옆')continue;const p=z.id==='B2'?[-4.8,33.7]:z.center;boxLabel([p[0],p[1],el===90?0:1.2],z.id+' 돗자리','#2e5670',canvas.clientWidth<700?10:12);}
};
const beforeSummary=updateLayoutSummary;
updateLayoutSummary=function(){
 beforeSummary();const sunday=isSunday();
 const title=$('layoutSummary'),count=$('layoutCount');if(sunday&&title)title.textContent='9.20. 전 구역 돗자리';if(sunday&&count)count.textContent='';
 const stamp=document.querySelector('.title span')||document.getElementById('layoutBadge');if(stamp)stamp.textContent='현장 배치 · '+(sunday?revision:'v28.1');
 const options=$('phase')?.options;if(options?.[0])options[0].textContent=sunday?'공연 · 전 구역 돗자리':'18시 이후·공연';
};
window.__hcfDay20Mats={version:revision,date:day,stageSide:newF07.location_label,footprints:copy(footprints),matQuantityConfirmed:false,getState:()=>({date:currentStaffDate(),chairs:scene.objects.filter(o=>String(o.kind).includes('chair')).length,mats:scene.objects.filter(o=>o.kind==='mat').length,F07:copy(scene.staff_phases.concert.find(p=>p.id==='F07')||null)})};
// The existing roster-sync module performs the initial rebuild after this patch.
})();
