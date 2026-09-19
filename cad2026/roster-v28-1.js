(function synchronizeFieldRoster(){
'use strict';
if(window.__hcfRosterV281)return;
const original=JSON.parse(document.getElementById('data').textContent);
const copy=x=>JSON.parse(JSON.stringify(x));
const sourceById=Object.fromEntries(Object.values(original.staff_registry).flat().map(x=>[x.id,x]));
const commandIds=new Set(scene.city_latest_instruction.command_person_ids);
const clean=t=>String(t||'').replaceAll('촬영대','화이트타워').replaceAll('후방 천막 3동','후방 천막 5동').replaceAll('뒤편 천막 3동','뒤편 천막 5동').replaceAll('에어바운스','워터슬라이드');
const set=(p,location,role)=>{p.location_label=location;p.role=role;p.assignment=role;};
const flow='통로에 멈추거나 모여서 관람하지 않도록 이동 방향 안내·계속 통행하도록 유도·혼잡 지속 시 현장 총괄에게 보고';
const previous=applyStaffRoster;
applyStaffRoster=function(){
 previous();const date=currentStaffDate(),sun=date==='2026-09-20';
 for(const phase of ['ceremony','concert'])for(const p of scene.staff_phases[phase]){
  p.location_label=clean(p.location_label);p.role=clean(p.role);
  if(commandIds.has(p.person_id))p.role=p.role.replaceAll('합동부스','안전부스');
  if(p.id==='S30'){
   const q=sourceById.S30;p.u=q.u;p.v=q.v;
   set(p,'B1 왼쪽 외곽 · 프레스 옆','프레스와 B1 사이 보행 통로 확보·통로 관람객 이동 안내·천막 출입구 앞 적치 방지');
  }
  if(p.id==='S33')p.role=p.role.replaceAll('합동부스','안전부스');
  if(p.id==='A03')p.location_label='A1 옆 좌석·프레스 연결 통로';
  if(sun&&p.id==='S13')p.role='무대 우측 외곽 통로 확보·천막 방향 무단 접근 안내·통로 관람객 이동 안내·인근 업체 안전요원 및 현장 총괄과 상황 공유';
  if(p.id==='G02'&&(!sun&&phase==='ceremony'))p.role='일반 관객의 우측 계단 진입 차단·진행팀 확인 후 수상자 통행 지원·G03과 계단 주변 상황 공유';
  if(p.id==='G02'&&sun)p.role='일반 관객의 우측 계단 진입 차단·진행팀 확인 후 출연진 통행 지원·현장 총괄과 상황 공유';
  if(p.id==='G05'&&sun)p.role='무대 후면 제한구역 출입 확인·일반 관객 우회 안내·현장 총괄 및 출연진 진행담당과 상황 공유';
  if(p.id==='A14'&&sun){p.location_label='10번 천막-B3 연결 휀스 동측 끝';p.role=flow+'·34번은 20일 명단에 없으므로 반대편 지원 배치는 현장 총괄 확인';}
  if(!sun&&['F15','F16'].includes(p.id)){
   const q=sourceById[p.id];p.u=q.u-1.7;p.v=q.v+9.35;
   set(p,'A3 옆으로 이동한 수상자석 주변 · 20:00 이후 합류','20:00 합류 후 이동한 좌석 주변 통로 확보·잔여 안내 및 관람객 이동 지원·17:00~18:00 시상 지원 인력에는 포함하지 않음');
  }
  if(p.id==='S03'&&!sun&&phase==='ceremony')p.confirmation_note='기존 명단은 공연 안전관리로 기재. 내빈·수상 지원 추가 지정 여부 확인 필요. 지원 담당으로 확인되면 09번 옆으로 변경.';
  if(sun&&p.org==='하남시청')p.confirmation_note='20일 근무시간이 원자료에 미기재. 집결·종료시간 확인 필요.';
  if(/^[AG]\d{2}$/.test(p.id))p.confirmation_note=(p.confirmation_note?p.confirmation_note+' / ':'')+'업체 실명·근무일·시간 미제공. 번호 기준 배치계획이며 실제 투입 확정 필요.';
  p.assignment=p.role;
 }
 if(!sun)scene.staff_phases.ceremony=scene.staff_phases.ceremony.filter(p=>{
  if(!/^F\d{2}$/.test(p.id))return true;
  const r=scene.operation_roster.foundation.find(x=>x.person_id===p.person_id);
  return (r?.shifts||[]).some(s=>s.date===date&&s.start<'18:00'&&s.end>'17:00');
 });
 const ph=document.getElementById('phase')?.value||'concert';
 const active=scene.staff_phases[ph]||[];
 const select=document.getElementById('staffSelect');
 if(select){const old=select.value;select.innerHTML='<option value="all">전체</option>'+active.map(p=>'<option value="'+p.id+'">'+p.id+' · '+p.org+' · '+p.role+'</option>').join('');select.value=active.some(p=>p.id===old)?old:'all';}
};
window.__hcfRosterV281={version:'v28.1-roster',sourceRevision:'v28.0-field',unconfirmed:['시청 20일 4명 근무시간','업체 21개 번호의 실명·근무일·시간','03번 장희진 내빈·수상 지원 추가 지정 여부','20일 A14 반대편 지원 편성'],note:'근무일별 전체 배치와 17~18시 시상 배치를 구분. 실제 현장 투입 및 변경 승인 여부는 별도 확인.'};
if(typeof rebuildOperationalWorld==='function')rebuildOperationalWorld();else applyStaffRoster();
})();

(function applyRequestedFoundationAssignments(){
'use strict';
if(window.__hcfFoundationAssignments)return;
const clone=x=>JSON.parse(JSON.stringify(x));
const identityList=scene.operation_roster.foundation;
const registry=scene.staff_registry.foundation;
const neighborIdentity=identityList.find(p=>p.post_id==='F12'||p.person_id==='F12');
const neighbor=registry.find(p=>p.id==='F12');
if(!neighbor||!neighborIdentity)throw new Error('한상일(F12) 기준 배치를 찾을 수 없습니다.');
const assignment='한상일(F12) 옆 현장 운영 지원';
const location='B3 오른쪽 통로 · 한상일(F12) 옆';
const known=identityList.find(p=>p.name==='박경득');
if(known&&known.person_id!=='F22')throw new Error('박경득 명단의 기존 번호를 확인해야 합니다.');
const park=known||{order:22,person_id:'F22',post_id:'F22',name:'박경득',group:'지역문화그룹',grade:'부장',shift:'근무시간 미기재',shifts:[],mobile:false,duty_excluded:false,command_only:false,escort:false};
Object.assign(park,{assignment,role:assignment,location_label:location,sunday_assignment:assignment,sunday_location:location});
if(!known)identityList.push(park);
let parkPost=registry.find(p=>p.id==='F22');
if(!parkPost){parkPost=clone(neighbor);registry.push(parkPost);}
Object.assign(parkPost,{id:'F22',person_id:'F22',name:'박경득',org:'하남문화재단',grade:'부장',role:assignment,assignment,location_label:location,remote:false,command_only:false,shirt_color:neighbor.shirt_color||'#dc8b38'});
delete parkPost.confirmation_note;delete parkPost.sunday_post;
// Program duties have no invented coordinates, dates, shifts, or attendance count.
const program=scene.operation_roster.program_staff||[];
for(const entry of [
 {person_id:'PROGRAM-KIM-JINSEONG',name:'김진성',organization:'하남문화재단',grade:'',assignment:'버스투어 담당',location_label:'버스투어 운영 구역',work_date:'',working_time:'',map_post_id:null},
 {person_id:'PROGRAM-CHOI-HYEONJU',name:'최현주',organization:'도시관광그룹',grade:'차장',assignment:'홍보 담당',location_label:'행사 홍보 업무 구역',work_date:'',working_time:'',map_post_id:null}
]){const old=program.find(p=>p.name===entry.name);if(old){old.assignment=entry.assignment;old.location_label=entry.location_label;}else program.push(entry);}
scene.operation_roster.program_staff=program;
const previous=applyStaffRoster;
applyStaffRoster=function(){
 previous();
 for(const key of ['ceremony','concert']){
  const entries=scene.staff_phases[key];if(!Array.isArray(entries))continue;
  const reference=entries.find(p=>p.id==='F12');
  scene.staff_phases[key]=entries.filter(p=>p.id!=='F22');
  if(!reference)continue;
  const positioned={...clone(parkPost),u:reference.u+1.2,v:reference.v,role:assignment,assignment,location_label:location,working_time:'근무시간 미기재'};
  scene.staff_phases[key].push(positioned);
 }
};
window.__hcfFoundationAssignments={version:'2026-09-19-F22',post:'F22',neighbor:'F12',name:'박경득',programStaff:clone(program)};
if(typeof rebuildOperationalWorld==='function')rebuildOperationalWorld();else applyStaffRoster();
})();
