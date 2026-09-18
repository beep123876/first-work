(function syncFieldRoster281(){
'use strict';
if(window.__hcfRoster281)return;
const version='v28.1';
const clone=x=>JSON.parse(JSON.stringify(x));
const input=JSON.parse(document.getElementById('data').textContent);
const source30=input.staff_registry.city.find(p=>p.id==='S30');
const source=scene.operation_roster;
const identities=[...source.city_hall.people,...source.foundation,...(source.vendors||[])];
const identity=p=>identities.find(q=>q.person_id===p.person_id)||p;
const escape=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const renamed=s=>String(s||'').replace(/천막\s*3동/g,'천막 5동').replace(/촬영대/g,'화이트타워');
function workingTime(p,date){
 const person=identity(p);
 if(p.id.startsWith('F')){const a=(person.shifts||[]).filter(x=>x.date===date);return a.length?a.map(x=>x.start+'~'+x.end).join(' / '):'시간 미기재';}
 if(p.id.startsWith('S'))return source.city_hall.hours?.[date]||'시간 미기재';
 return '근무일·시간 미제공';
}
function attention(p,date,phase){
 const notes=[];const record=identity(p);
 if(p.id.startsWith('F')){
  const shifts=(record.shifts||[]).filter(s=>s.date===date);
  if(date==='2026-09-19'){
   const start=phase==='ceremony'?'17:00':'18:00',end=phase==='ceremony'?'18:00':'22:00';
   if(!shifts.some(s=>s.start<end&&s.end>start))notes.push('이 단계 근무시간 외');
   if(shifts.length&&shifts.every(s=>s.start>start))notes.push(shifts.map(s=>s.start).sort()[0]+'부터 합류');
   if(shifts.length&&shifts.every(s=>s.end<=start))notes.push('이 단계 시작 전 근무 종료');
  }
 }
 if(date==='2026-09-20'&&p.id==='F15')notes.push('14:00 투입 전 후방 통로 담당자 확인 필요');
 if(date==='2026-09-20'&&p.id==='A14')notes.push('34번은 19일 명단에만 있음: 반대편 요원 확인 필요');
 if(date==='2026-09-20'&&p.id==='F03')notes.push('수상자석 미운영: 해당 위치의 공연 지원 임무 확인 필요');
 if(date==='2026-09-20'&&p.id.startsWith('S'))notes.push('시청 20일 근무시간 미기재');
 if(/^[AG]/.test(p.id))notes.push('업체 성명·근무시간 확인 필요');
 return notes.join(' / ');
}
function normalize(){
 const date=currentStaffDate();
 for(const phase of ['ceremony','concert'])for(const p of scene.staff_phases[phase]){
  p.location_label=renamed(p.location_label);p.role=renamed(p.role);
  // S30 is the passage post beside the former joint booth, not an indoor command post.
  if(p.id==='S30'&&source30){Object.assign(p,{u:source30.u,v:source30.v,group:source30.group,command_only:false,location_label:'B1 왼쪽 외곽 · 프레스 부스 옆',role:'프레스 부스와 B1 사이 보행 통로 확보·정지 관람객 이동 안내·천막 출입구 앞 적치 방지'});}
  if(date==='2026-09-19'&&phase==='ceremony'&&p.id==='G02')p.role=p.role.replace(/S03/g,'G03');
  if(date==='2026-09-20'&&p.id==='S13')p.role=p.role.replace(/S20과 우측 접근 상황 공유/g,'인근 경호·안전요원과 우측 접근 상황 공유');
  if(date==='2026-09-20'&&p.id==='A14')p.location_label='10번 천막–B3 연결 휀스 동측 끝';
  p.working_time=workingTime(p,date);p.attendance_note=attention(p,date,phase);
 }
}
function makeRows(date,phase){
 return scene.staff_phases[phase].map(p=>{
  const q=identity(p),row={date,phase,post_id:p.id,number:p.id.startsWith('S')?p.id.slice(1):p.id,person_id:p.person_id,name:q.name||'성명 미제공',organization:/^[SF]/.test(p.id)?(q.group||p.org||''):(p.org||''),grade:q.grade||'',working_time:p.working_time,location:p.location_label||'',role:p.role||'',note:p.attendance_note||'',u:p.u,v:p.v};
  if(!q.field_posts)q.field_posts={};q.field_posts[date+'/'+phase]={location:row.location,role:row.role,u:p.u,v:p.v,working_time:row.working_time,note:row.note};
  return row;
 });
}
function refresh(){
 normalize();const date=currentStaffDate(),phase=date==='2026-09-20'?'concert':($('phase').value||'concert');
 const byPhase=Object.fromEntries(['ceremony','concert'].map(k=>[k,makeRows(date,k)]));
 scene.field_roster={version,date,phase,by_phase:byPhase,rows:byPhase[phase],notice:'해당일 배치 계획입니다. 근무시간 외 인원을 실제 투입 인원으로 집계하지 마세요.'};
 const select=$('staffSelect'),selected=select.value;
 select.innerHTML='<option value="all">전체</option>'+byPhase[phase].map(r=>'<option value="'+escape(r.post_id)+'">'+escape(r.number+' · '+r.name+' · '+r.location+' · '+r.role)+'</option>').join('');
 select.value=byPhase[phase].some(r=>r.post_id===selected)?selected:'all';
}
const before=applyStaffRoster;
applyStaffRoster=function(){before();refresh();};
function exportCSV(){
 const data=scene.field_roster;
 const fields=['number','name','organization','grade','working_time','location','role','note'];
 const safe=v=>'"'+String(v??'').replace(/^[=+@-]/,"'$&").replace(/"/g,'""')+'"';
 const text='\ufeff'+[['번호','성명','소속','직급','근무시간','배치 장소','담당 업무','확인사항'].map(safe).join(','),...data.rows.map(r=>fields.map(k=>safe(r[k])).join(','))].join('\r\n');
 const url=URL.createObjectURL(new Blob([text],{type:'text/csv;charset=utf-8'}));const link=document.createElement('a');link.href=url;link.download='이성산성문화제_'+data.date+'_'+data.phase+'_명단_v28.1.csv';link.click();setTimeout(()=>URL.revokeObjectURL(url),30000);
}
window.__hcfRoster281={version,refresh,exportCSV,getRows:()=>clone(scene.field_roster.rows)};
rebuildOperationalWorld();
const toolbar=document.querySelector('.field-tools')||document.querySelector('header');
if(toolbar&&!document.getElementById('fieldRosterCSV')){const button=document.createElement('button');button.id='fieldRosterCSV';button.type='button';button.textContent='명단';button.title='선택한 날짜·행사 단계의 명단 CSV 저장';button.setAttribute('aria-label',button.title);button.addEventListener('click',exportCSV);toolbar.append(button);}
const badge=document.querySelector('.title span')||document.getElementById('layoutBadge');if(badge)badge.textContent='현장 배치 · '+version;
})();
