(function applySeptember20Operations(){
'use strict';
if(window.__hcfDay20Operations)return;
const day='2026-09-20',version='v28.3';
const flow='통로에 멈추거나 모여서 관람하지 않도록 이동 방향 안내·계속 통행하도록 유도·혼잡 시 현장 총괄에게 보고';
const assignments={
 F07:{location_label:'시스템부스 앞',role:'시스템부스 담당자·진행팀 연락·공연 진행 요청사항 전달·시스템부스 주변 통행로 확보·혼잡 및 요청사항을 현장 총괄에게 전달'},
 F22:{location_label:'메인무대 오른편 · 관객 기준',role:'무대 오른편 현장 운영 지원·주변 통행로 확보·정지 관람객 이동 안내·혼잡 및 요청사항을 현장 총괄에게 전달'}
};
function systemFront(){
 const booth=scene.objects.find(o=>o.id==='SYSTEM-1');
 if(!booth?.poly?.length)throw new Error('시스템부스 기준 시설을 찾을 수 없습니다.');
 const a=booth.poly,i=booth.opening_side??2;
 const c=a.reduce((p,q)=>[p[0]+q[0]/a.length,p[1]+q[1]/a.length],[0,0]);
 const e=[(a[i][0]+a[(i+1)%a.length][0])/2,(a[i][1]+a[(i+1)%a.length][1])/2];
 const d=Math.hypot(e[0]-c[0],e[1]-c[1])||1;
 return {u:e[0]+1.1*(e[0]-c[0])/d,v:e[1]+1.1*(e[1]-c[1])/d};
}
function audienceRight(){
 const stage=scene.objects.find(o=>o.id==='STAGE');
 if(!stage?.poly?.length)throw new Error('메인무대 기준 시설을 찾을 수 없습니다.');
 // Front edge runs from the audience-left corner [3] to audience-right [2].
 const a=stage.poly[3],b=stage.poly[2],w=Math.hypot(b[0]-a[0],b[1]-a[1]);
 const right=[(b[0]-a[0])/w,(b[1]-a[1])/w],front=[-right[1],right[0]];
 // Existing clear position outside the right side tower; does not move the stage.
 return {u:b[0]+7.65*right[0]+1.5*front[0],v:b[1]+7.65*right[1]+1.5*front[1]};
}
const previous=applyStaffRoster;
applyStaffRoster=function(){
 previous();
 if(currentStaffDate()!==day)return;
 const positions={F07:systemFront(),F22:audienceRight()};
 for(const phase of ['ceremony','concert'])for(const p of scene.staff_phases[phase]||[]){
  if(assignments[p.id]){
   Object.assign(p,positions[p.id],assignments[p.id],{remote:false,group:'towers'});
   p.assignment=p.role;
   delete p.external_area;
   if(p.sunday_post)p.sunday_post={...p.sunday_post,...positions[p.id],...assignments[p.id]};
  }
  if(p.id==='F12'){p.role=flow;p.assignment=flow;if(p.sunday_post)p.sunday_post={...p.sunday_post,role:flow};}
 }
 // Refresh the visible selector and CSV rows after positions and duties are changed.
 if(window.__hcfRoster281?.refresh)window.__hcfRoster281.refresh();
};
const previousSummary=updateLayoutSummary;
updateLayoutSummary=function(){
 previousSummary();
 if(currentStaffDate()===day){
  const label=document.querySelector('.title span')||document.getElementById('layoutBadge');
  if(label)label.textContent='현장 배치 · '+version;
 }
};
window.__hcfDay20Operations={version,defaultDate:day,assignments,getState:()=>({
 date:currentStaffDate(),phase:$('phase').value,
 posts:scene.staff_phases.concert.filter(p=>['F07','F12','F22'].includes(p.id)).map(p=>({id:p.id,u:p.u,v:p.v,location:p.location_label,role:p.role})),
 chairs:scene.objects.filter(o=>String(o.kind).includes('chair')).length
})};
// Set the launch default once; users remain free to select 19 September afterwards.
const date=$('staffDate');
if(date){const option=Array.from(date.options).find(o=>o.value===day);if(option){date.insertBefore(option,date.firstChild);for(const o of date.options)o.defaultSelected=o.value===day;}date.value=day;}
const phase=$('phase');if(phase)phase.value='concert';
rebuildOperationalWorld();
requestRender();
})();
